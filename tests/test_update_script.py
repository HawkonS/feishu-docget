import os
import shutil
import stat
import subprocess
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
UPDATE_SCRIPT = PROJECT_ROOT / 'tools' / 'update.sh'
GIT = shutil.which('git')


def run_git(args, cwd):
    return subprocess.run(
        [GIT, *args],
        cwd=cwd,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


class UpdateScriptTests(unittest.TestCase):
    def _create_repository(self, workspace):
        repo = workspace / 'repo'
        repo.mkdir()
        (repo / 'tools').mkdir()
        shutil.copy2(UPDATE_SCRIPT, repo / 'tools' / 'update.sh')
        (repo / 'tools' / 'update.sh').chmod(
            (repo / 'tools' / 'update.sh').stat().st_mode | stat.S_IXUSR
        )

        run_git(['init', '-b', 'main'], repo)
        run_git(['config', 'user.email', 'test@example.com'], repo)
        run_git(['config', 'user.name', 'Update Script Test'], repo)
        (repo / 'version.txt').write_text('initial\n', encoding='utf-8')
        run_git(['add', 'version.txt'], repo)
        run_git(['commit', '-m', 'initial version'], repo)

        remote = workspace / 'remote.git'
        run_git(['init', '--bare', remote], workspace)
        run_git(['remote', 'add', 'origin', str(remote)], repo)
        run_git(['push', '-u', 'origin', 'main'], repo)
        return repo, remote

    def _write_command(self, path, content):
        path.write_text(content, encoding='utf-8')
        path.chmod(path.stat().st_mode | stat.S_IXUSR)

    def _environment(self, workspace, bin_dir, **extra):
        env = os.environ.copy()
        env.update({
            'PATH': f'{bin_dir}{os.pathsep}{env["PATH"]}',
            'REAL_GIT': GIT,
        })
        env.update(extra)
        return env

    def test_fetch_failure_does_not_reset_or_restart(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workspace = Path(temp_dir)
            repo, _ = self._create_repository(workspace)
            bin_dir = workspace / 'bin'
            bin_dir.mkdir()
            calls = workspace / 'git-calls.log'
            restart_marker = workspace / 'restart.marker'

            self._write_command(
                bin_dir / 'git',
                '#!/bin/sh\n'
                'printf "%s\\n" "$*" >> "$GIT_CALLS"\n'
                'if [ "$1" = "fetch" ]; then exit 1; fi\n'
                'exec "$REAL_GIT" "$@"\n',
            )
            self._write_command(
                repo / 'stop.sh',
                f'#!/bin/sh\ntouch "{restart_marker}"\n',
            )

            before = run_git(['rev-parse', 'HEAD'], repo).stdout.strip()
            result = subprocess.run(
                ['bash', str(repo / 'tools' / 'update.sh'), '--yes'],
                cwd=repo,
                env=self._environment(workspace, bin_dir, GIT_CALLS=str(calls)),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )

            self.assertNotEqual(result.returncode, 0)
            self.assertIn('远程代码拉取失败', result.stdout)
            self.assertIn('服务未重启', result.stdout)
            self.assertFalse(restart_marker.exists())
            self.assertEqual(run_git(['rev-parse', 'HEAD'], repo).stdout.strip(), before)
            self.assertIn('fetch --prune origin main', calls.read_text(encoding='utf-8'))
            self.assertNotIn('reset --hard', calls.read_text(encoding='utf-8'))

    def test_successful_update_logs_applied_version_and_restarts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workspace = Path(temp_dir)
            repo, remote = self._create_repository(workspace)
            bin_dir = workspace / 'bin'
            bin_dir.mkdir()
            restart_marker = workspace / 'restart.marker'

            upstream = workspace / 'upstream'
            run_git(['clone', str(remote), upstream], workspace)
            run_git(['config', 'user.email', 'test@example.com'], upstream)
            run_git(['config', 'user.name', 'Update Script Test'], upstream)
            (upstream / 'version.txt').write_text('updated\n', encoding='utf-8')
            run_git(['add', 'version.txt'], upstream)
            run_git(['commit', '-m', 'updated version'], upstream)
            target = run_git(['rev-parse', 'HEAD'], upstream).stdout.strip()
            run_git(['push', 'origin', 'main'], upstream)

            self._write_command(
                bin_dir / 'systemctl',
                '#!/bin/sh\n'
                'case "$1" in\n'
                '  list-unit-files) printf "feishu-docget.service enabled\\n" ;;\n'
                f'  restart) touch "{restart_marker}" ;;\n'
                '  status) ;;\n'
                'esac\n',
            )
            self._write_command(bin_dir / 'sudo', '#!/bin/sh\nexec "$@"\n')

            result = subprocess.run(
                ['bash', str(repo / 'tools' / 'update.sh'), '--yes'],
                cwd=repo,
                env=self._environment(workspace, bin_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )

            self.assertEqual(result.returncode, 0, result.stdout)
            self.assertEqual(run_git(['rev-parse', 'HEAD'], repo).stdout.strip(), target)
            self.assertTrue(restart_marker.exists())
            self.assertIn('当前版本:', result.stdout)
            self.assertIn('目标版本:', result.stdout)
            self.assertIn('代码已更新到版本:', result.stdout)


if __name__ == '__main__':
    unittest.main()
