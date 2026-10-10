# ops/install_env.py
"""Create or refresh .venv offline from vendor/wheels; print the venv python path."""
import hashlib, os, subprocess, sys, venv
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent

def lock_sha256(lock):
    return hashlib.sha256(Path(lock).read_bytes()).hexdigest()

def needs_install(venv_dir, lock):
    stamp = Path(venv_dir) / '.lock-sha256'
    return not stamp.exists() or stamp.read_text().strip() != lock_sha256(lock)

def venv_python(venv_dir):
    return Path(venv_dir) / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')

def main():
    venv_dir, lock = ROOT / '.venv', ROOT / 'requirements.lock'
    if needs_install(venv_dir, lock):
        if not venv_python(venv_dir).exists():
            venv.EnvBuilder(with_pip=True).create(venv_dir)
        args = [str(venv_python(venv_dir)), '-m', 'pip', 'install', '--require-hashes', '-r', str(lock)]
        if os.name == 'nt':
            args[4:4] = ['--no-index', '--find-links', str(ROOT / 'vendor' / 'wheels')]
        subprocess.run(args, check=True, stdout=sys.stderr)
        (venv_dir / '.lock-sha256').write_text(lock_sha256(lock))
    print(venv_python(venv_dir))

if __name__ == '__main__':
    main()
