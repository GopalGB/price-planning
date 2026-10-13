"""Operational checks: WAL-safe backups, settings, and real launcher lifecycle."""
from contextlib import closing
import json
import os
from pathlib import Path
import secrets
import signal
import socket
import sqlite3
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
import urllib.request

from app.runtime import DEFAULTS, ROOT, ServerLock, backup_database, validate_settings, lan_addresses, resolve_addresses


class Runtime(unittest.TestCase):
    def test_settings_require_matching_origin_and_secure_cookies(self):
        self.assertEqual(validate_settings(DEFAULTS)['port'], 8000)
        for update in ({'port': 80}, {'port': True}, {'public_origin': 'http://localhost:9000'},
                       {'public_origin': 'https://planning.example.com', 'cookie_secure': False},
                       {'public_origin': 'http://user:password@localhost:8000'},
                       {'public_origin': 'http://0.0.0.0:8000'}, {'data_dir': '//server/shared'}):
            with self.subTest(update=update), self.assertRaises(ValueError):
                validate_settings(DEFAULTS | update)
        https = validate_settings(DEFAULTS | dict(public_origin='https://planning.example.com', cookie_secure=True))
        self.assertTrue(https['cookie_secure'])

    def test_online_backup_includes_wal_and_can_restore_to_new_directory(self):
        with tempfile.TemporaryDirectory(prefix='planning runtime ') as temp:
            data = Path(temp) / 'data'
            data.mkdir()
            with closing(sqlite3.connect(data / 'planning.db')) as writer:
                writer.execute('PRAGMA journal_mode=WAL')
                writer.execute('PRAGMA wal_autocheckpoint=0')
                writer.execute('CREATE TABLE inputs(value TEXT)')
                writer.commit()
                writer.execute('INSERT INTO inputs VALUES(?)', ('saved in WAL',))
                writer.commit()
                self.assertGreater((data / 'planning.db-wal').stat().st_size, 0)
                saved = backup_database(data, Path(temp) / 'backups')
                self.assertTrue(saved.exists())
                # Open the standalone backup in a new location: no source WAL is present.
                restore = Path(temp) / 'restored'
                restore.mkdir()
                restored = restore / 'planning.db'
                restored.write_bytes(saved.read_bytes())
                with closing(sqlite3.connect(restored)) as db:
                    self.assertEqual(db.execute('SELECT value FROM inputs').fetchone()[0], 'saved in WAL')
                    self.assertEqual(db.execute('PRAGMA integrity_check').fetchone()[0], 'ok')
                self.assertFalse(list((Path(temp) / 'backups').glob('*.partial')))

    def test_duplicate_server_lock_and_release(self):
        with tempfile.TemporaryDirectory() as temp:
            data = Path(temp)
            with ServerLock(data):
                result = subprocess.run([sys.executable, '-c',
                    'from app.runtime import ServerLock; from pathlib import Path; import sys; '
                    'ServerLock(Path(sys.argv[1])).__enter__()', str(data)],
                    cwd=ROOT, capture_output=True, text=True, timeout=10)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn('already running', result.stderr)
            with ServerLock(data):
                pass

    def test_first_admin_restart_and_real_waitress_launcher(self):
        with tempfile.TemporaryDirectory(prefix='planning launcher ') as temp:
            data = Path(temp) / 'data'
            data.mkdir()
            with closing(socket.socket()) as sock:
                sock.bind(('127.0.0.1', 0))
                port = sock.getsockname()[1]
            settings = Path(temp) / 'settings.json'
            cfg = DEFAULTS | dict(demo_mode=False,bind='127.0.0.1',data_dir=str(data), port=port, public_origin='http://localhost:%s' % port)
            settings.write_text(json.dumps(cfg))
            env = os.environ.copy()
            env['DATA_DIR'] = str(data)
            env.pop('DATABASE_URL', None)
            # A tiny imported fixture keeps the lifecycle check independent of model size.
            subprocess.run([sys.executable, '-c',
                "from app.db import DB,init_schema; db=DB(); init_schema(db); "
                "db.execute('INSERT INTO meta VALUES(?,?)',('ready','true')); db.commit(); db.close()"],
                cwd=ROOT, env=env, check=True, capture_output=True, timeout=10)
            email = 'launcher@local.test'
            password = secrets.token_urlsafe(24)
            # Mock only terminal input; run the real initialization and password hashing.
            script = (
                'import sys; from unittest.mock import patch; from run import main; '
                'email,password,path=sys.argv[1:]; '
                'ctx1=patch("builtins.input",side_effect=[email,"Launcher admin"]); '
                'ctx2=patch("getpass.getpass",return_value=password); '
                'ctx1.start();ctx2.start();sys.exit(main(["--settings",path,"--init-only"]))')
            # Secrets are supplied through stdin, never the child command line.
            script = script.replace('email,password,path=sys.argv[1:]',
                'import json; email,password,path=json.load(sys.stdin)')
            initialized = subprocess.run([sys.executable, '-c', script],
                input=json.dumps([email, password, str(settings)]), cwd=ROOT, env=env,
                capture_output=True, text=True, timeout=20)
            self.assertEqual(initialized.returncode, 0, initialized.stderr)
            self.assertNotIn(password, settings.read_text())
            with closing(sqlite3.connect(data / 'planning.db')) as db:
                self.assertEqual(db.execute('SELECT COUNT(*) FROM users WHERE admin=1').fetchone()[0], 1)
                self.assertNotEqual(db.execute('SELECT password FROM users').fetchone()[0], password)
            # The next launch must need no terminal input or first-time setup.
            cmd = [sys.executable, '-u', str(ROOT / 'run.py'), '--settings', str(settings), '--no-browser']
            process = subprocess.Popen(cmd, cwd=temp, env=env, stdin=subprocess.DEVNULL,
                                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            try:
                base = 'http://localhost:%s' % port
                for _ in range(60):
                    if process.poll() is not None:
                        self.fail('Server exited before readiness: %s' % process.stderr.read())
                    try:
                        with urllib.request.urlopen(base + '/api/health', timeout=0.2) as response:
                            health = json.load(response)
                        break
                    except OSError:
                        time.sleep(0.05)
                else:
                    self.fail('Waitress did not become ready')
                self.assertTrue(health['imported'])
                req = urllib.request.Request(base + '/api/login',
                    data=json.dumps(dict(email=email, password=password)).encode(),
                    headers={'Content-Type': 'application/json'}, method='POST')
                with urllib.request.urlopen(req, timeout=5) as response:
                    self.assertEqual(json.load(response)['email'], email)
                    self.assertIn('HttpOnly', response.headers['Set-Cookie'])
            finally:
                process.send_signal(signal.SIGINT)
                try:
                    output, errors = process.communicate(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()  # a server that ignores Ctrl+C must not outlive the test
                    process.communicate()
                    raise
            self.assertEqual(process.returncode, 0, errors)
            self.assertIn('READY:', output)
            self.assertNotIn(password, output + errors)

    def test_lan_detection_prefers_10_address_and_lists_local_origin(self):
        addresses = [
            (socket.AF_INET, socket.SOCK_STREAM, 6, '', ('192.168.1.5', 0)),
            (socket.AF_INET, socket.SOCK_STREAM, 6, '', ('10.0.0.25', 0)),
            (socket.AF_INET, socket.SOCK_STREAM, 6, '', ('127.0.0.1', 0)),
            (socket.AF_INET, socket.SOCK_STREAM, 6, '', ('169.254.1.8', 0))]
        with patch('app.runtime.socket.getaddrinfo', return_value=addresses):
            self.assertEqual(lan_addresses(), ['10.0.0.25', '192.168.1.5'])
            origin, allowed = resolve_addresses(DEFAULTS)
        self.assertEqual(origin, 'http://10.0.0.25:8000')
        self.assertIn('http://localhost:8000', allowed)

    def test_demo_initialization_has_no_admin_or_settings_prompts(self):
        with tempfile.TemporaryDirectory(prefix='demo start ') as temp:
            data = Path(temp) / 'data'
            data.mkdir()
            env = os.environ.copy()
            env['DATA_DIR'] = str(data)
            env.pop('DATABASE_URL', None)
            subprocess.run([sys.executable, '-c',
                "from app.db import DB,init_schema; db=DB(); init_schema(db); "
                "db.execute('INSERT INTO meta VALUES(?,?)',('ready','true')); db.commit(); db.close()"],
                cwd=ROOT, env=env, check=True, capture_output=True, timeout=10)
            settings = Path(temp) / 'settings.json'
            started = subprocess.run([sys.executable, str(ROOT / 'run.py'), '--demo',
                '--settings', str(settings), '--init-only'], cwd=temp, env=env,
                stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=10)
            self.assertEqual(started.returncode, 0, started.stderr)
            self.assertTrue(json.loads(settings.read_text())['demo_mode'])
            self.assertEqual(json.loads(settings.read_text())['bind'], '0.0.0.0')
            with closing(sqlite3.connect(data / 'planning.db')) as db:
                self.assertEqual(db.execute('SELECT COUNT(*) FROM users WHERE admin=1').fetchone()[0], 0)


if __name__ == '__main__':
    unittest.main()
