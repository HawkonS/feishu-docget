import os
import stat
import json
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from src.app import (
    app, config, _resolve_template_path, _script_json, paginate_items,
    _is_system_admin_session,
)
from src.core.stats import _mask_url
from src.core.sqlite_store import migrate_legacy_data, list_download_stats, list_users
from src.core.user_store import SYSTEM_ADMIN_NAME, SYSTEM_ADMIN_OPEN_ID


class SecurityRegressionTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    def _csrf(self):
        with self.client.session_transaction() as session:
            return session['_csrf_token']

    def test_sensitive_config_is_owner_only(self):
        mode = stat.S_IMODE(os.stat('feishu-docget.properties').st_mode)
        self.assertEqual(mode, 0o600)

    def test_sensitive_config_reveal_requires_explicit_admin_request(self):
        with self.client.session_transaction() as session:
            session['is_admin'] = True
            session['_csrf_token'] = 'test-csrf-token'

        config_response = self.client.get('/api/config')
        self.assertEqual(config_response.status_code, 200)
        app_secret = next(item for item in config_response.get_json() if item['key'] == 'feishu.app_secret')
        self.assertTrue(app_secret['masked'])
        self.assertEqual(app_secret['value'], '******')

        self.assertEqual(
            self.client.post('/api/config/reveal', json={'key': 'feishu.app_secret'}).status_code,
            403,
        )
        response = self.client.post(
            '/api/config/reveal',
            json={'key': 'feishu.app_secret'},
            headers={'X-CSRF-Token': 'test-csrf-token'},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()['value'], config.get('feishu.app_secret'))
        self.assertEqual(response.headers.get('Cache-Control'), 'no-store')

        invalid = self.client.post(
            '/api/config/reveal',
            json={'key': 'page.title'},
            headers={'X-CSRF-Token': 'test-csrf-token'},
        )
        self.assertEqual(invalid.status_code, 400)

        with self.client.session_transaction() as session:
            session.pop('is_admin', None)
            session['user'] = {'open_id': 'ou_feishu_admin', 'name': '飞书管理员'}
        self.assertEqual(self.client.get('/api/config').status_code, 403)
        self.assertEqual(
            self.client.post(
                '/api/config/reveal',
                json={'key': 'feishu.app_secret'},
                headers={'X-CSRF-Token': 'test-csrf-token'},
            ).status_code,
            403,
        )

    def test_csrf_required_for_state_changes(self):
        self.client.get('/admin')
        self.assertEqual(
            self.client.post('/api/admin/login', json={'password': config.get('admin.password')}).status_code,
            403,
        )
        self.assertEqual(self.client.post('/api/admin/logout').status_code, 403)
        token = self._csrf()
        login = self.client.post(
            '/api/admin/login',
            json={'password': config.get('admin.password')},
            headers={'X-CSRF-Token': token},
        )
        self.assertEqual(login.status_code, 200)
        self.assertEqual(login.get_json()['status'], 'ok')
        self.assertEqual(
            self.client.post('/api/admin/logout', headers={'X-CSRF-Token': token}).status_code,
            200,
        )
        with self.client.session_transaction() as session:
            self.assertNotIn('job_tokens', session)

    def test_template_path_cannot_escape_template_directory(self):
        self.assertIsNone(_resolve_template_path('../feishu-docget.properties'))
        self.assertIsNone(_resolve_template_path('/etc/passwd'))
        self.assertTrue(_resolve_template_path(config.get('template.default', 'template.docx')))
        template_dir = os.path.realpath(config.get('template.dir', 'template'))
        link_path = os.path.join(template_dir, '.security-link.docx')
        try:
            os.symlink(os.path.realpath('feishu-docget.properties'), link_path)
            self.assertIsNone(_resolve_template_path(os.path.basename(link_path)))
        finally:
            if os.path.lexists(link_path):
                os.unlink(link_path)

    def test_stats_only_keeps_allowed_https_hosts(self):
        self.assertEqual(_mask_url('javascript:alert(1)'), '')
        self.assertEqual(_mask_url('https://evil.example/x'), '')
        self.assertEqual(_mask_url('https://foo.feishu.cn/wiki/abc?token=secret'), 'https://foo.feishu.cn/wiki/abc')

    def test_script_json_escapes_html_breakout_characters(self):
        encoded = _script_json('</script><img src=x>')
        self.assertNotIn('</script>', encoded.lower())
        self.assertIn('\\u003c', encoded)

    def test_admin_data_endpoints_return_server_pagination(self):
        with self.client.session_transaction() as session:
            session['is_admin'] = True
        for path, collection_key in (
            ('/api/admin/projects?page=1&page_size=1', 'items'),
            ('/api/admin/templates?page=1&page_size=1', 'templates'),
            ('/api/admin/users?page=1&page_size=1', 'users'),
            ('/api/stats?page=1&page_size=1', 'items'),
        ):
            response = self.client.get(path)
            self.assertEqual(response.status_code, 200)
            payload = response.get_json()
            self.assertIn(collection_key, payload)
            self.assertIn('page', payload)
            self.assertIn('page_size', payload)
            self.assertIn('total_pages', payload)
            self.assertLessEqual(len(payload[collection_key]), 1)

    def test_paginate_items_does_not_return_other_pages(self):
        result = paginate_items([{'id': i} for i in range(5)], page=2, page_size=2)
        self.assertEqual([item['id'] for item in result['items']], [2, 3])
        self.assertEqual(result['total'], 5)
        self.assertTrue(result['has_more'])

    def test_legacy_json_migrates_users_and_deduplicated_stats(self):
        with tempfile.TemporaryDirectory() as workspace:
            log_dir = os.path.join(workspace, 'logs')
            os.makedirs(log_dir)
            with open(os.path.join(log_dir, 'users.json'), 'w', encoding='utf-8') as handle:
                json.dump({
                    'ou_test': {
                        'name': '迁移用户',
                        'access_token': 'access-secret',
                        'refresh_token': 'refresh-secret',
                        'is_admin': True,
                    }
                }, handle)
            with open(os.path.join(log_dir, 'download_stats.jsonl'), 'w', encoding='utf-8') as handle:
                handle.write(json.dumps({'id': 'job-1', 'status': '下载中', 'ts': 1}) + '\n')
                handle.write(json.dumps({'id': 'job-1', 'status': '已完成', 'ts': 2, 'title': '文档'}) + '\n')

            migration_config = {'workspace.dir': workspace, 'log.dir': 'logs'}
            result = migrate_legacy_data(workspace, migration_config, backup=True)
            self.assertEqual(result['users_imported'], 1)
            self.assertEqual(result['stats_imported'], 1)
            self.assertEqual(len(result['backups']), 2)

            users = list_users(workspace, migration_config, page=1, page_size=10)
            self.assertEqual(users['total'], 1)
            self.assertTrue(users['items'][0]['is_admin'])
            stats = list_download_stats(workspace, migration_config, page=1, page_size=10)
            self.assertEqual(stats['total'], 1)
            self.assertEqual(stats['items'][0]['status'], '已完成')
            self.assertEqual(stats['items'][0]['title'], '文档')
            self.assertEqual(
                list_users(workspace, migration_config, page=1, page_size=10, query='迁移')['total'],
                1,
            )

    def test_system_admin_is_fixed_and_separate_from_feishu_identity(self):
        from src.core import user_store

        self.assertTrue(user_store.ensure_system_admin())
        system = user_store.get_user(SYSTEM_ADMIN_OPEN_ID)
        self.assertEqual(system['name'], SYSTEM_ADMIN_NAME)
        self.assertFalse(system['is_admin'])
        self.assertTrue(system['is_system_admin'])
        self.assertFalse(system['disabled'])
        self.assertTrue(system['is_system'])
        self.assertFalse(user_store.set_disabled(SYSTEM_ADMIN_OPEN_ID, True))
        self.assertFalse(user_store.set_admin(SYSTEM_ADMIN_OPEN_ID, False))
        system = user_store.get_user(SYSTEM_ADMIN_OPEN_ID)
        self.assertFalse(system['disabled'])
        self.assertFalse(system['is_admin'])
        self.assertTrue(system['is_system_admin'])

    def test_feishu_admin_keeps_real_name_on_homepage(self):
        previous_login_enabled = config.get('login.enabled')
        config['login.enabled'] = 'true'
        try:
            with self.client.session_transaction() as session:
                session['user'] = {'open_id': 'ou_real_name', 'name': '真实姓名'}
            with patch('src.app.user_store.is_admin', return_value=True):
                response = self.client.get('/')
            self.assertEqual(response.status_code, 200)
            html = response.get_data(as_text=True)
            self.assertIn('真实姓名', html)
            self.assertNotIn('user-card-name" id="userCardName">管理员</span>', html)
        finally:
            config['login.enabled'] = previous_login_enabled

    def test_admin_login_page_shows_feishu_entry_only_when_login_enabled(self):
        previous_login_enabled = config.get('login.enabled')
        try:
            config['login.enabled'] = 'true'
            enabled_html = self.client.get('/admin').get_data(as_text=True)
            self.assertIn('/auth/feishu/authorize?next=admin', enabled_html)
            self.assertNotIn('<div hidden>', enabled_html)

            config['login.enabled'] = 'false'
            disabled_html = self.client.get('/admin').get_data(as_text=True)
            self.assertIn('<div hidden>', disabled_html)
        finally:
            config['login.enabled'] = previous_login_enabled

    def test_dashboard_uses_system_and_operator_role_labels_only(self):
        with self.client.session_transaction() as session:
            session['is_admin'] = True
            session['user'] = {'open_id': SYSTEM_ADMIN_OPEN_ID, 'name': SYSTEM_ADMIN_NAME}
        response = self.client.get('/admin')
        html = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn('系统管理员', html)
        self.assertIn('运营管理员', html)
        self.assertIn('普通用户', html)
        self.assertIn('/api/admin/users/set-role', html)
        self.assertIn('>角色</th>', html)
        self.assertNotIn('toggleUserAdmin', html)
        self.assertNotIn('toggleSystemAdmin', html)
        self.assertNotIn('最高权限', html)
        self.assertNotIn('绑定最高权限', html)
        self.assertNotIn('/api/admin/users/bind-system-admin', html)
        self.assertIn('no-store', response.headers.get('Cache-Control', ''))

    def test_admin_modals_use_bootstrap_structure_and_lifecycle(self):
        with self.client.session_transaction() as session:
            session['is_admin'] = True
            session['user'] = {'open_id': SYSTEM_ADMIN_OPEN_ID, 'name': SYSTEM_ADMIN_NAME}

        html = self.client.get('/admin').get_data(as_text=True)

        self.assertIn('id="uploadTemplateModal" class="modal fade admin-modal"', html)
        self.assertIn('id="renameTemplateModal" class="modal fade admin-modal"', html)
        self.assertIn('modal-dialog-centered modal-dialog-scrollable', html)
        self.assertGreaterEqual(html.count('data-bs-dismiss="modal"'), 4)
        self.assertIn('/static/vendor/bootstrap.min.css', html)
        self.assertIn('/static/vendor/bootstrap.bundle.min.js', html)
        self.assertIn('bootstrap.Modal.getOrCreateInstance', html)
        self.assertNotIn('modal.style.display = "flex"', html)

    def test_config_cards_use_consistent_title_stack_and_tooltip_notes(self):
        with self.client.session_transaction() as session:
            session['is_admin'] = True
            session['user'] = {'open_id': SYSTEM_ADMIN_OPEN_ID, 'name': SYSTEM_ADMIN_NAME}

        html = self.client.get('/admin').get_data(as_text=True)

        self.assertIn('className = \'cfg-item-head\'', html)
        self.assertIn('className = \'cfg-title-stack\'', html)
        self.assertIn('data-bs-toggle', html)
        self.assertIn('data-bs-custom-class', html)
        self.assertIn('function refreshConfigTooltips()', html)
        self.assertIn("new bootstrap.Tooltip(el", html)
        self.assertIn("[item.label, item.key, item.desc, item.hint, item.preview]", html)
        self.assertIn('.cfg-item { min-height: 118px;', html)
        self.assertIn('margin-top: 7px', html)
        self.assertIn('height: 43px; min-height: 43px; flex: 0 0 43px;', html)
        self.assertNotIn('margin-top: auto; margin-bottom: 4px', html)
        self.assertNotIn("wrap.appendChild(descEl)", html)

    def test_frontend_modals_use_local_bootstrap_structure_and_lifecycle(self):
        previous_login_enabled = config.get('login.enabled')
        config['login.enabled'] = 'false'
        try:
            html = self.client.get('/').get_data(as_text=True)
        finally:
            config['login.enabled'] = previous_login_enabled

        self.assertIn('/static/vendor/bootstrap.min.css', html)
        self.assertIn('/static/vendor/bootstrap.bundle.min.js', html)
        self.assertIn('id="uploadModal" class="modal fade frontend-modal"', html)
        self.assertIn('id="advancedModal" class="modal fade frontend-modal"', html)
        self.assertGreaterEqual(html.count('modal-dialog-centered modal-dialog-scrollable'), 2)
        self.assertGreaterEqual(html.count('data-bs-dismiss="modal"'), 5)
        self.assertIn('bootstrap.Modal.getOrCreateInstance', html)
        self.assertNotIn('document.getElementById("uploadModal").style.display', html)
        self.assertNotIn('document.getElementById("advancedModal").style.display', html)
        self.assertNotIn('cdn.jsdelivr.net/npm/bootstrap', html)

        css_response = self.client.get('/static/vendor/bootstrap.min.css?v=5.3.3')
        js_response = self.client.get('/static/vendor/bootstrap.bundle.min.js?v=5.3.3')
        try:
            self.assertEqual(css_response.status_code, 200)
            self.assertEqual(js_response.status_code, 200)
            self.assertEqual(css_response.headers.get('Cache-Control'), 'public, max-age=31536000, immutable')
            self.assertEqual(js_response.headers.get('Cache-Control'), 'public, max-age=31536000, immutable')
            self.assertNotIn('Set-Cookie', css_response.headers)
            self.assertNotIn('Set-Cookie', js_response.headers)
        finally:
            css_response.close()
            js_response.close()

    def test_admin_oauth_target_is_preserved_and_requires_admin_role(self):
        previous_login_enabled = config.get('login.enabled')
        config['login.enabled'] = 'true'
        try:
            with patch('src.app.feishu_oauth.build_authorize_url', return_value='https://example.com/auth'):
                response = self.client.get('/auth/feishu/authorize?next=admin')
            self.assertEqual(response.status_code, 302)
            with self.client.session_transaction() as session:
                self.assertEqual(session['oauth_target'], 'admin')
                state = session['oauth_state']

            tokens = {
                'access_token': 'access', 'refresh_token': 'refresh', 'expires_in': 7200,
                'refresh_token_expires_in': 3600, 'scope': 'docx:document',
            }
            info = {'open_id': 'ou_not_admin', 'name': '普通用户'}
            with patch('src.app.feishu_oauth.exchange_code', return_value=tokens), \
                    patch('src.app.feishu_oauth.get_oauth_user_info', return_value=info), \
                    patch('src.app.user_store.get_user', return_value=None), \
                    patch('src.app.user_store.upsert_user', return_value=True), \
                    patch('src.app.user_store.is_admin', return_value=False):
                response = self.client.get(f'/auth/feishu/callback?state={state}&code=test')
            self.assertEqual(response.status_code, 302)
            self.assertTrue(response.location.endswith('/admin?error=not_admin'))
            with self.client.session_transaction() as session:
                self.assertEqual(session['user']['name'], '普通用户')

        finally:
            config['login.enabled'] = previous_login_enabled

    def test_feishu_system_admin_gets_system_rights_and_keeps_name(self):
        previous_login_enabled = config.get('login.enabled')
        config['login.enabled'] = 'true'
        try:
            with self.client.session_transaction() as session:
                session['user'] = {'open_id': 'ou_bound_admin', 'name': '张三'}
            with patch('src.app.user_store.is_system_admin_role', return_value=True), \
                    patch('src.app.user_store.is_admin', return_value=True), \
                    patch('src.app.user_store.is_disabled', return_value=False), \
                    patch('src.app.user_store.get_user', return_value={
                        'open_id': 'ou_bound_admin', 'name': '张三', 'avatar': '',
                        'disabled': False, 'is_system_admin': True,
                    }):
                with app.test_request_context('/'):
                    app.preprocess_request()
                    session['user'] = {'open_id': 'ou_bound_admin', 'name': '张三'}
                    self.assertTrue(_is_system_admin_session())
                self.assertEqual(self.client.get('/api/config').status_code, 200)
                html = self.client.get('/').get_data(as_text=True)
            self.assertIn('user-card-name" id="userCardName">张三</span>', html)
            self.assertNotIn('user-card-name" id="userCardName">管理员</span>', html)
        finally:
            config['login.enabled'] = previous_login_enabled

    def test_system_admin_role_allows_multiple_users_and_preserves_user_name(self):
        from src.core import sqlite_store

        with tempfile.TemporaryDirectory() as workspace:
            migration_config = {'workspace.dir': workspace, 'log.dir': 'logs'}
            for open_id, name in (('ou_first', '张三'), ('ou_second', '李四')):
                sqlite_store.upsert_user(workspace, migration_config, {
                    'open_id': open_id, 'name': name, 'disabled': False,
                })

            self.assertTrue(sqlite_store.set_system_admin(
                workspace, migration_config, 'ou_first', True,
            ))
            self.assertTrue(sqlite_store.set_system_admin(
                workspace, migration_config, 'ou_second', True,
            ))
            first = sqlite_store.get_user(workspace, migration_config, 'ou_first')
            second = sqlite_store.get_user(workspace, migration_config, 'ou_second')
            self.assertTrue(first['is_system_admin'])
            self.assertTrue(second['is_system_admin'])
            self.assertEqual(second['name'], '李四')

    def test_operator_admin_cannot_manage_system_admin(self):
        with self.client.session_transaction() as session:
            session['user'] = {'open_id': 'ou_operator', 'name': '运营'}
            session['_csrf_token'] = 'operator-csrf'
        with patch('src.app.user_store.is_admin', return_value=True), \
                patch('src.app.user_store.is_system_admin_role', side_effect=lambda open_id: open_id == 'ou_system'), \
                patch('src.app.user_store.get_user', return_value={'open_id': 'ou_system'}), \
                patch('src.app.user_store.is_system_admin', return_value=False):
            response = self.client.post(
                '/api/admin/users/set-role',
                json={'open_id': 'ou_system', 'role': 'operator_admin'},
                headers={'X-CSRF-Token': 'operator-csrf'},
            )
        self.assertEqual(response.status_code, 403)
        self.assertIn('运营管理员不能管理系统管理员', response.get_json()['message'])

    def test_operator_admin_cannot_grant_system_admin_role(self):
        with self.client.session_transaction() as session:
            session['user'] = {'open_id': 'ou_operator', 'name': '运营'}
            session['_csrf_token'] = 'operator-csrf'
        with patch('src.app.user_store.is_admin', return_value=True), \
                patch('src.app.user_store.is_system_admin_role', return_value=False), \
                patch('src.app.user_store.get_user', return_value={
                    'open_id': 'ou_user', 'disabled': False,
                }), \
                patch('src.app.user_store.is_system_admin', return_value=False):
            response = self.client.post(
                '/api/admin/users/set-role',
                json={'open_id': 'ou_user', 'role': 'system_admin'},
                headers={'X-CSRF-Token': 'operator-csrf'},
            )
        self.assertEqual(response.status_code, 403)
        self.assertIn('仅系统管理员可设置系统管理员', response.get_json()['message'])

    def test_operator_admin_cannot_access_system_controls(self):
        with self.client.session_transaction() as session:
            session['user'] = {'open_id': 'ou_operator', 'name': '运营'}
            session['_csrf_token'] = 'operator-csrf'
        with patch('src.app.user_store.is_admin', return_value=True), \
                patch('src.app.user_store.is_system_admin_role', return_value=False):
            response = self.client.post(
                '/api/admin/system',
                json={'action': 'status'},
                headers={'X-CSRF-Token': 'operator-csrf'},
            )
        self.assertEqual(response.status_code, 403)

    def test_system_admin_role_allows_multiple_users(self):
        from src.core import sqlite_store

        with tempfile.TemporaryDirectory() as workspace:
            migration_config = {'workspace.dir': workspace, 'log.dir': 'logs'}
            for open_id in ('ou_system_one', 'ou_system_two'):
                sqlite_store.upsert_user(workspace, migration_config, {
                    'open_id': open_id, 'name': open_id, 'disabled': False,
                })
            self.assertTrue(sqlite_store.set_system_admin(
                workspace, migration_config, 'ou_system_one', True,
            ))
            self.assertTrue(sqlite_store.set_system_admin(
                workspace, migration_config, 'ou_system_two', True,
            ))
            self.assertTrue(sqlite_store.get_user(
                workspace, migration_config, 'ou_system_one',
            )['is_system_admin'])
            self.assertTrue(sqlite_store.get_user(
                workspace, migration_config, 'ou_system_two',
            )['is_system_admin'])
            self.assertFalse(sqlite_store.get_user(
                workspace, migration_config, 'ou_system_one',
            )['is_admin'])

    def test_system_admin_inherits_operator_admin_permissions(self):
        from src.core import user_store

        previous_workspace = config.get('workspace.dir')
        previous_log_dir = config.get('log.dir')
        with tempfile.TemporaryDirectory() as workspace:
            try:
                config['workspace.dir'] = workspace
                config['log.dir'] = 'logs'
                self.assertTrue(user_store.upsert_user({
                    'open_id': 'ou_inherited', 'name': '继承权限用户',
                }))
                self.assertTrue(user_store.set_role(
                    'ou_inherited', user_store.SYSTEM_ADMIN_ROLE,
                ))
                self.assertTrue(user_store.is_system_admin_role('ou_inherited'))
                self.assertTrue(user_store.is_admin('ou_inherited'))
            finally:
                config['workspace.dir'] = previous_workspace
                config['log.dir'] = previous_log_dir

    def test_old_single_binding_schema_migrates_to_multi_system_admin_roles(self):
        from src.core import sqlite_store

        with tempfile.TemporaryDirectory() as workspace:
            log_dir = os.path.join(workspace, 'logs')
            os.makedirs(log_dir)
            db_path = os.path.join(log_dir, sqlite_store.DATABASE_FILENAME)
            connection = sqlite3.connect(db_path)
            connection.executescript(
                """
                CREATE TABLE users (
                    open_id TEXT PRIMARY KEY,
                    union_id TEXT NOT NULL DEFAULT '', user_id TEXT NOT NULL DEFAULT '',
                    name TEXT NOT NULL DEFAULT '', avatar TEXT NOT NULL DEFAULT '',
                    disabled INTEGER NOT NULL DEFAULT 0, is_admin INTEGER NOT NULL DEFAULT 0,
                    system_admin_bound INTEGER NOT NULL DEFAULT 0,
                    access_token TEXT NOT NULL DEFAULT '', refresh_token TEXT NOT NULL DEFAULT '',
                    token_expire_at INTEGER NOT NULL DEFAULT 0,
                    refresh_token_expire_at INTEGER NOT NULL DEFAULT 0,
                    scope TEXT NOT NULL DEFAULT '', token_invalid INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT '', last_login_at TEXT NOT NULL DEFAULT ''
                );
                CREATE UNIQUE INDEX idx_users_single_system_admin_bound
                    ON users(system_admin_bound) WHERE system_admin_bound = 1;
                INSERT INTO users(open_id, name, system_admin_bound)
                    VALUES('ou_old_system', '旧系统管理员', 1);
                INSERT INTO users(open_id, name) VALUES('ou_new_system', '新系统管理员');
                """
            )
            connection.commit()
            connection.close()

            migration_config = {'workspace.dir': workspace, 'log.dir': 'logs'}
            sqlite_store.initialize_database(workspace, migration_config)
            old_system = sqlite_store.get_user(
                workspace, migration_config, 'ou_old_system',
            )
            self.assertTrue(old_system['is_system_admin'])
            self.assertEqual(old_system['system_admin_bound'], 0)
            self.assertTrue(sqlite_store.set_system_admin(
                workspace, migration_config, 'ou_new_system', True,
            ))
            self.assertTrue(sqlite_store.get_user(
                workspace, migration_config, 'ou_new_system',
            )['is_system_admin'])


if __name__ == '__main__':
    unittest.main()
