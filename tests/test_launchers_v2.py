import os, shutil, sqlite3, subprocess, sys, tempfile, unittest
from contextlib import closing
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent


class LaunchersV2(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='pp launcher ')
        self.repo = Path(self.tmp.name)
        shutil.copy(ROOT / 'run_v2.sh', self.repo)
        shutil.copy(ROOT / 'run_v2.py', self.repo)
        for pkg in ('pp', 'app'):
            shutil.copytree(ROOT / pkg, self.repo / pkg, ignore=shutil.ignore_patterns('__pycache__'))
        shutil.copytree(ROOT / 'ops', self.repo / 'ops', ignore=shutil.ignore_patterns('__pycache__'))
        # stub installer: the launcher only needs the venv interpreter path on stdout
        (self.repo / 'ops' / 'install_env.py').write_text('import sys\nprint(sys.executable)\n')
        from ops.fixtures import make_semantic
        (self.repo / 'template').mkdir()
        make_semantic(self.repo / 'template' / 'synthetic.xlsm')
        (self.repo / 'settings.json').write_text('{"demo_mode": true, "template_path": "template/synthetic.xlsm"}')

    def tearDown(self):
        self.tmp.cleanup()

    def sh(self, *args):
        environment = {**os.environ, 'PATH': str(Path(sys.executable).parent) + os.pathsep + os.environ['PATH']}
        for name in ('DATA_DIR', 'PP_DATABASE_PATH', 'PP_TEMPLATE_PATH'):
            environment.pop(name, None)
        return subprocess.run(['sh', str(self.repo / 'run_v2.sh'), *args], capture_output=True, text=True,
                              env=environment)

    @unittest.skipIf(os.name == 'nt', 'POSIX launcher')
    def test_sh_starts_app_with_args(self):
        res = self.sh('--import-template')
        self.assertEqual(res.returncode, 0, res.stderr)
        snapshots = list((self.repo / 'data' / 'templates').glob('*/meta.json'))
        self.assertEqual(len(snapshots), 1, res.stdout)

    @unittest.skipIf(os.name == 'nt', 'POSIX launcher')
    def test_sh_backup_and_restore_route_to_cli(self):
        db = self.repo / 'data' / 'planning.db'
        db.parent.mkdir()
        with closing(sqlite3.connect(db)) as con:
            con.execute('CREATE TABLE t (x)'); con.execute('INSERT INTO t VALUES (1)'); con.commit()
        made = self.sh('--backup')
        self.assertEqual(made.returncode, 0, made.stderr)
        out = Path(made.stdout.strip().splitlines()[-1])
        self.assertEqual(out.parent, (self.repo / 'backups').resolve())
        back = self.sh('--restore', str(out), '--dest', str(self.repo / 'restored'))
        self.assertEqual(back.returncode, 0, back.stderr)
        self.assertTrue((self.repo / 'restored' / 'planning.db').is_file())

    def test_bat_mirrors_legacy_resolution_and_routes(self):
        bat = (ROOT / 'run_v2.bat').read_text(encoding='utf-8')
        for must in ('python\\python.exe', 'py -3', 'where python', 'ops\\install_env.py', 'run_v2.py', '-m pp.backup',
                     '--backup', '--restore', 'sys.version_info >= (3, 10)'):
            self.assertIn(must, bat)
        self.assertNotIn('run.py', bat.replace('run_v2.py', ''))
        self.assertIn('\r\n', (ROOT / 'run_v2.bat').read_bytes().decode())


class DeployDocV2(unittest.TestCase):
    def test_deploy_doc_covers_v2(self):
        text = (ROOT / 'docs' / 'DEPLOY.md').read_text(encoding='utf-8')
        for must in ('template\\', 'PP_TEMPLATE_PATH', 'run_v2.bat', 'run_v2.bat --backup',
                     'run_v2.bat --restore', 'python -m pp.backup'):
            self.assertIn(must, text)

    def test_template_folder_ignored(self):
        res = subprocess.run(['git', 'check-ignore', '-q', 'template/notes.txt'], cwd=ROOT)
        self.assertEqual(res.returncode, 0)


if __name__ == '__main__':
    unittest.main()
