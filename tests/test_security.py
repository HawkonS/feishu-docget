import os
import stat
import json
import tempfile
import unittest
from unittest.mock import patch

from src.app import app, config, _resolve_template_path, _script_json, paginate_items
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
        self.assertTrue(system['is_admin'])
        self.assertFalse(system['disabled'])
        self.assertTrue(system['is_system'])
        self.assertFalse(user_store.set_disabled(SYSTEM_ADMIN_OPEN_ID, True))
        self.assertFalse(user_store.set_admin(SYSTEM_ADMIN_OPEN_ID, False))
        system = user_store.get_user(SYSTEM_ADMIN_OPEN_ID)
        self.assertFalse(system['disabled'])
        self.assertTrue(system['is_admin'])

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


if __name__ == '__main__':
    unittest.main()
