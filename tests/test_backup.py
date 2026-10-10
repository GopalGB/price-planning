import errno, hashlib, json, os, sqlite3, subprocess, sys, tempfile, unittest
from unittest import mock
from contextlib import closing
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from pp import backup as bk


def rows(path):
    with closing(sqlite3.connect(path)) as con:
        return [r[0] for r in con.execute('SELECT v FROM item ORDER BY id')]


class Backup(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)
        self.db = self.d / 'data' / 'planning.db'
        self.db.parent.mkdir()
        with closing(sqlite3.connect(self.db)) as con:
            con.execute('PRAGMA journal_mode=WAL')
            con.execute('PRAGMA user_version=7')
            con.execute('CREATE TABLE parent (id INTEGER PRIMARY KEY)')
            con.execute('CREATE TABLE item (id INTEGER PRIMARY KEY, v TEXT, p INTEGER REFERENCES parent(id))')
            con.execute("INSERT INTO item (v) VALUES ('a'), ('b')")
            con.commit()

    def tearDown(self):
        self.tmp.cleanup()

    def test_backup_during_open_write_transaction(self):
        writer = sqlite3.connect(self.db, isolation_level=None)
        writer.execute("INSERT INTO item (v) VALUES ('c')")  # committed (autocommit)
        writer.execute('BEGIN IMMEDIATE')
        writer.execute("INSERT INTO item (v) VALUES ('uncommitted')")
        try:
            out = bk.backup(self.db, self.d / 'backups')
        finally:
            writer.execute('ROLLBACK'); writer.close()
        self.assertEqual(rows(out), ['a', 'b', 'c'])
        meta = json.loads(out.with_suffix('.json').read_text())
        self.assertEqual(meta['schema_version'], 7)
        self.assertEqual(meta['sha256'], hashlib.sha256(out.read_bytes()).hexdigest())
        self.assertEqual(meta['source'], str(self.db.resolve()))
        self.assertIn('created_at', meta)
        self.assertFalse(Path(str(out) + '-wal').exists())

    def test_backup_refuses_broken_foreign_keys(self):
        with closing(sqlite3.connect(self.db)) as con:
            con.execute("INSERT INTO item (v, p) VALUES ('orphan', 99)"); con.commit()
        with self.assertRaises(ValueError):
            bk.backup(self.db, self.d / 'backups')
        self.assertEqual(list((self.d / 'backups').glob('*.db')), [])

    def test_restore_round_trip_and_refuses_non_empty_dir(self):
        out = bk.backup(self.db, self.d / 'backups')
        restored = bk.restore(out, self.d / 'new')
        self.assertEqual(restored, self.d / 'new' / 'planning.db')
        self.assertEqual(rows(restored), ['a', 'b'])
        busy = self.d / 'busy'; busy.mkdir(); (busy / 'x').write_text('x')
        with self.assertRaises(ValueError):
            bk.restore(out, busy)

    def test_corrupted_backup_refused(self):
        out = bk.backup(self.db, self.d / 'backups')
        data = bytearray(out.read_bytes()); data[-100] ^= 0xFF
        out.write_bytes(bytes(data))
        with self.assertRaises(ValueError):
            bk.restore(out, self.d / 'new1')
        # matching checksum but broken pages: integrity check must still refuse
        data = bytearray(out.read_bytes()); data[4096:8192] = b'\x07' * 4096
        out.write_bytes(bytes(data))
        meta = json.loads(out.with_suffix('.json').read_text())
        meta['sha256'] = hashlib.sha256(bytes(data)).hexdigest()
        out.with_suffix('.json').write_text(json.dumps(meta))
        with self.assertRaises(ValueError):
            bk.restore(out, self.d / 'new2')
        self.assertFalse((self.d / 'new2' / 'planning.db').exists())

    def test_restore_copy_failure_leaves_nothing_and_retry_works(self):
        out = bk.backup(self.db, self.d / 'backups')
        def boom(src, dst):
            Path(dst).write_bytes(b'partial')
            raise OSError(errno.ENOSPC, 'No space left on device')
        with mock.patch.object(bk.shutil, 'copyfile', side_effect=boom):
            with self.assertRaises(OSError):
                bk.restore(out, self.d / 'new')
        self.assertEqual(list((self.d / 'new').iterdir()), [])
        self.assertEqual(rows(bk.restore(out, self.d / 'new')), ['a', 'b'])
        self.assertEqual([p.name for p in (self.d / 'new').iterdir()], ['planning.db'])

    def test_second_restore_refuses_and_keeps_existing(self):
        out = bk.backup(self.db, self.d / 'backups')
        target = bk.restore(out, self.d / 'new')
        before = target.read_bytes()
        with self.assertRaises(ValueError):
            bk.restore(out, self.d / 'new')
        self.assertEqual(target.read_bytes(), before)
        # publish never overwrites, even if the target appears after the emptiness check
        real_copy = bk.shutil.copyfile
        def racer(src, dst):
            real_copy(src, dst)
            (self.d / 'race' / 'planning.db').write_bytes(b'someone else')
        with mock.patch.object(bk.shutil, 'copyfile', side_effect=racer):
            with self.assertRaises(ValueError):
                bk.restore(out, self.d / 'race')
        self.assertEqual((self.d / 'race' / 'planning.db').read_bytes(), b'someone else')
        self.assertEqual([p.name for p in (self.d / 'race').iterdir()], ['planning.db'])

    def test_default_db_follows_server_precedence(self):
        configured = self.d / 'configured'
        settings = self.d / 'settings.json'
        settings.write_text(json.dumps({'data_dir': str(configured)}))
        env = {k: v for k, v in os.environ.items() if k != 'DATA_DIR'}
        with mock.patch.dict(os.environ, env, clear=True):
            self.assertEqual(bk.default_db(settings), configured.resolve() / 'planning.db')
            self.assertEqual(bk.default_db(self.d / 'missing.json'), (bk.ROOT / 'data').resolve() / 'planning.db')
            settings.write_text(json.dumps({'data_dir': 'rel-data'}))
            self.assertEqual(bk.default_db(settings), (bk.ROOT / 'rel-data').resolve() / 'planning.db')
            with mock.patch.dict(os.environ, {'DATA_DIR': str(self.d / 'envdir')}):
                self.assertEqual(bk.default_db(settings), (self.d / 'envdir').resolve() / 'planning.db')

    def test_default_db_uses_pp_database_path_and_skips_unrelated_server_validation(self):
        settings = self.d / 'settings.json'
        settings.write_text(json.dumps({'port': 0, 'database_path': 'configured/planning.db'}))
        env = {k: v for k, v in os.environ.items() if k not in {'DATA_DIR', 'PP_DATABASE_PATH'}}
        with mock.patch.dict(os.environ, env, clear=True):
            self.assertEqual(bk.default_db(settings), (bk.ROOT / 'configured' / 'planning.db').resolve())
        with mock.patch.dict(os.environ, {'PP_DATABASE_PATH': str(self.d / 'environment.db')}, clear=True):
            self.assertEqual(bk.default_db(settings), (self.d / 'environment.db').resolve())

    def test_cli_backs_up_configured_db_not_default(self):
        configured = self.d / 'configured'; configured.mkdir()
        with closing(sqlite3.connect(configured / 'planning.db')) as con:
            con.execute('CREATE TABLE item (id INTEGER PRIMARY KEY, v TEXT)')
            con.execute("INSERT INTO item (v) VALUES ('configured')"); con.commit()
        settings = self.d / 'settings.json'
        settings.write_text(json.dumps({'data_dir': str(configured)}))
        env = {k: v for k, v in os.environ.items() if k != 'DATA_DIR'}
        made = subprocess.run([sys.executable, '-m', 'pp.backup', '--settings', str(settings), '--dest', str(self.d / 'bk')],
                              cwd=ROOT, capture_output=True, text=True, env=env)
        self.assertEqual(made.returncode, 0, made.stderr)
        self.assertEqual(rows(made.stdout.strip().splitlines()[-1]), ['configured'])

    def test_cli(self):
        run = lambda *a: subprocess.run([sys.executable, '-m', 'pp.backup', *a], cwd=ROOT, capture_output=True, text=True)
        made = run('--db', str(self.db), '--dest', str(self.d / 'bk'))
        self.assertEqual(made.returncode, 0, made.stderr)
        out = Path(made.stdout.strip().splitlines()[-1])
        back = run('--restore', str(out), '--dest', str(self.d / 'r'))
        self.assertEqual(back.returncode, 0, back.stderr)
        self.assertEqual(rows(self.d / 'r' / 'planning.db'), ['a', 'b'])
        bad = run('--restore', str(out), '--dest', str(self.d / 'r'))
        self.assertNotEqual(bad.returncode, 0)
        self.assertNotIn('Traceback', bad.stderr)


if __name__ == '__main__':
    unittest.main()
