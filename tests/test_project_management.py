import os
import shutil
import tempfile
import unittest
from unittest.mock import patch

from src.app import (
    _cleanup_broken_project_links,
    app,
    config,
    list_project_files,
    list_project_summaries,
)


class ProjectManagementRegressionTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    def _create_linked_projects(self, workspace):
        output_dir = os.path.join(workspace, 'output')
        master_dir = os.path.join(output_dir, 'document')
        master_img_dir = os.path.join(master_dir, 'img')
        referenced_dir = os.path.join(output_dir, 'document-1')
        os.makedirs(master_img_dir)
        os.makedirs(referenced_dir)
        document_path = os.path.join(referenced_dir, 'document.docx')
        with open(document_path, 'wb') as handle:
            handle.write(b'docx')
        img_link = os.path.join(referenced_dir, 'img')
        try:
            os.symlink('../document/img', img_link)
        except OSError as error:
            self.skipTest(f'当前环境无法创建软链接: {error}')
        return output_dir, master_dir, referenced_dir, img_link

    def test_historical_broken_img_link_does_not_leave_1970_entry(self):
        with tempfile.TemporaryDirectory() as workspace:
            output_dir, master_dir, referenced_dir, img_link = \
                self._create_linked_projects(workspace)
            shutil.rmtree(master_dir)

            files = list_project_files(referenced_dir)
            self.assertEqual([item['name'] for item in files], ['document.docx'])
            self.assertNotIn('1970-01-01', ''.join(item['ctime'] for item in files))

            with patch('src.app.base_dir', workspace), \
                    patch.dict(config, {'output.dir': 'output'}):
                summaries = list_project_summaries()
            self.assertEqual(summaries['total'], 1)
            self.assertEqual(summaries['items'][0]['file_count'], 1)

            removed = _cleanup_broken_project_links(output_dir)
            self.assertEqual(removed, [img_link])
            self.assertFalse(os.path.lexists(img_link))

    def test_delete_project_endpoint_removes_broken_img_links(self):
        with tempfile.TemporaryDirectory() as workspace:
            _, master_dir, referenced_dir, img_link = \
                self._create_linked_projects(workspace)
            with self.client.session_transaction() as session:
                session['is_admin'] = True
                session['_csrf_token'] = 'test-csrf-token'

            with patch('src.app.base_dir', workspace), \
                    patch.dict(config, {'output.dir': 'output'}):
                response = self.client.post(
                    '/api/admin/delete_project',
                    json={'path': master_dir},
                    headers={'X-CSRF-Token': 'test-csrf-token'},
                )

            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.get_json()['status'], 'ok')
            self.assertFalse(os.path.exists(master_dir))
            self.assertFalse(os.path.lexists(img_link))
            self.assertEqual(
                [item['name'] for item in list_project_files(referenced_dir)],
                ['document.docx'],
            )


if __name__ == '__main__':
    unittest.main()
