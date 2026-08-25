import os
import stat
import unittest

from src.app import app, config, _resolve_template_path, _script_json
from src.core.stats import _mask_url


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


if __name__ == '__main__':
    unittest.main()
