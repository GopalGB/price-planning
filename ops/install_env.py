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

def works(python, *args):
    try:
        subprocess.run([str(python), *args], check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, timeout=30)
        return True
    except (OSError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return False

def main():
    venv_dir, lock = ROOT / '.venv', ROOT / 'requirements.lock'
    python = venv_python(venv_dir)
    usable = python.exists() and works(python, '--version')
    pip_ready = usable and works(python, '-m', 'pip', '--version')
    if needs_install(venv_dir, lock) or not pip_ready:
        stamp = venv_dir / '.lock-sha256'
        stamp.unlink(missing_ok=True)
        if not usable:
            venv.EnvBuilder(with_pip=True, clear=True).create(venv_dir)
        if not works(python, '--version'):
            raise RuntimeError('virtual environment interpreter is unusable after recreation')
        if not works(python, '-m', 'pip', '--version'):
            subprocess.run([str(python), '-m', 'ensurepip', '--upgrade'], check=True, stdout=sys.stderr)
        if not works(python, '-m', 'pip', '--version'):
            raise RuntimeError('virtual environment pip is unavailable after bootstrap')
        args = [str(python), '-m', 'pip', 'install', '--require-hashes', '-r', str(lock)]
        if os.name == 'nt':
            args[4:4] = ['--no-index', '--find-links', str(ROOT / 'vendor' / 'wheels')]
        subprocess.run(args, check=True, stdout=sys.stderr)
        stamp.write_text(lock_sha256(lock))
    print(python)

if __name__ == '__main__':
    main()
