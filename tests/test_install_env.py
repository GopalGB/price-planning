# tests/test_install_env.py
import sys, tempfile, unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'ops'))
import install_env as ie

class InstallStamp(unittest.TestCase):
    def test_install_stamp_changes_with_lock(self):
        with tempfile.TemporaryDirectory() as d:
            venv, lock = Path(d) / '.venv', Path(d) / 'requirements.lock'
            lock.write_text('a==1\n'); venv.mkdir()
            self.assertTrue(ie.needs_install(venv, lock))
            (venv / '.lock-sha256').write_text(ie.lock_sha256(lock))
            self.assertFalse(ie.needs_install(venv, lock))
            lock.write_text('a==2\n')
            self.assertTrue(ie.needs_install(venv, lock))
    def test_lock_has_hashes_and_pins_engine(self):
        t = (ROOT / 'requirements.lock').read_text()
        self.assertIn('formualizer==0.11.1', t)
        self.assertIn('--hash=sha256:', t)
