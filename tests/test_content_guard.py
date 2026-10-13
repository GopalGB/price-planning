import hashlib
import os
import shutil
import shlex
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import content_guard as cg


class CheckTests(unittest.TestCase):
    def run_check(self, files, deny=(), allow=None):
        return cg.check(list(files), lambda p: files[p], list(deny), allow or {})

    def test_spreadsheet_extension_blocked(self):
        self.assertTrue(self.run_check({'a/Plan.XLSM': b'x'}))

    def test_big_file_blocked_unless_allowlisted(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'vendor/wheels/x.whl': big}))
        allow = {'vendor/wheels/x.whl': hashlib.sha256(big).hexdigest()}
        self.assertEqual(self.run_check({'vendor/wheels/x.whl': big}, allow=allow), [])

    def test_allowlist_hash_mismatch_blocked(self):
        big = b'0' * (cg.MAX_BYTES + 1)
        self.assertTrue(self.run_check({'v.whl': big}, allow={'v.whl': 'deadbeef'}))

    def test_denylisted_string_in_txt_blocked(self):
        self.assertTrue(self.run_check({'notes.txt': b'Group 7 BU Recap'}, deny=['bu recap']))

    def test_clean_files_pass(self):
        self.assertEqual(self.run_check({'app/x.py': b'print(1)'}, deny=['secret label']), [])


class HookTests(unittest.TestCase):
    def git_shim(self, repo, staged=''):
        """Model read-only Git queries so hook tests never create or stage a repo."""
        shim = repo / 'bin'
        shim.mkdir()
        git = shim / 'git'
        git.write_text('#!/bin/sh\ncase "$1" in\n'
                       f'  rev-parse) printf "%s\\n" {shlex.quote(str(repo))} ;;\n'
                       f'  diff) printf "%s" {shlex.quote(staged)}'
                       + ('; printf "\\0"' if staged else '') + ' ;;\n'
                       '  *) exit 2 ;;\nesac\n')
        git.chmod(0o755)
        return shim

    def test_hook_runs_bash_maxhq_hook_when_present(self):
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d) / 'r'
            (repo / 'ops').mkdir(parents=True)
            (repo / '.githooks').mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json', '.githooks/pre-commit'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            hook_dir = Path(d) / 'security' / 'tools'
            hook_dir.mkdir(parents=True)
            (hook_dir / 'pre-commit-hook.sh').write_text(
                '#!/bin/bash\nset -euo pipefail\nitems=(ok)\n[[ ${items[0]} == ok ]]\n')
            # macOS sh accepts Bash syntax; model a strictly POSIX-only sh.
            shim = self.git_shim(repo)
            (shim / 'sh').write_text('#!/bin/sh\nexit 77\n')
            (shim / 'sh').chmod(0o755)
            env = dict(os.environ, PATH=str(shim) + os.pathsep + os.environ['PATH'])
            result = subprocess.run([shutil.which('sh'), '.githooks/pre-commit'], cwd=repo,
                                    env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_hook_runs_without_maxhq(self):
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d) / 'r'
            (repo / 'ops').mkdir(parents=True)
            (repo / '.githooks').mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json', '.githooks/pre-commit'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            (repo / 'bad.xlsx').write_bytes(b'x')
            shim = self.git_shim(repo, 'bad.xlsx')
            env = dict(os.environ, PATH=str(shim) + os.pathsep + os.environ['PATH'])
            result = subprocess.run(['sh', '.githooks/pre-commit'], cwd=repo, env=env,
                                    capture_output=True, text=True)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertIn('BLOCKED bad.xlsx', result.stdout)


if __name__ == '__main__':
    unittest.main()


class StubInterpreterTests(unittest.TestCase):
    def test_hook_skips_broken_python3_stub(self):
        import shutil, stat
        with tempfile.TemporaryDirectory() as d:
            repo, bindir = Path(d) / 'r', Path(d) / 'bin'
            subprocess.run(['git', 'init', '-q', str(repo)], check=True)
            (repo / 'ops').mkdir(); (repo / '.githooks').mkdir(); bindir.mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json', '.githooks/pre-commit'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            stub = bindir / 'python3'
            stub.write_text('#!/bin/sh\nexit 9009\n'); stub.chmod(stub.stat().st_mode | stat.S_IEXEC)
            (bindir / 'python').symlink_to(sys.executable)
            for tool in ('git', 'sh', 'dirname'):
                found = shutil.which(tool)
                if found:
                    (bindir / tool).symlink_to(found)
            (repo / 'bad.xlsx').write_bytes(b'x')
            subprocess.run(['git', '-C', str(repo), 'add', '-A'], check=True)
            env = {'PATH': str(bindir), 'HOME': d}
            r = subprocess.run(['/bin/sh', '.githooks/pre-commit'], cwd=repo, env=env, capture_output=True, text=True)
            self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
            self.assertIn('BLOCKED bad.xlsx', r.stdout)


class AllModeTests(unittest.TestCase):
    def test_all_skips_tracked_files_deleted_in_worktree(self):
        # A rebuild deletes old hashed bundle files before they are staged; --all must not crash.
        with tempfile.TemporaryDirectory() as d:
            repo = Path(d)
            (repo / 'ops').mkdir()
            for rel in ('ops/content_guard.py', 'ops/large_file_allowlist.json'):
                (repo / rel).write_bytes((ROOT / rel).read_bytes())
            (repo / 'present.txt').write_text('clean\n')
            shim = repo / 'bin'
            shim.mkdir()
            (shim / 'git').write_text('#!/bin/sh\nprintf "gone.js\\0present.txt\\0"\n')
            (shim / 'git').chmod(0o755)
            env = dict(os.environ, PATH=str(shim) + os.pathsep + os.environ['PATH'])
            result = subprocess.run([sys.executable, 'ops/content_guard.py', '--all'], cwd=repo,
                                    env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
