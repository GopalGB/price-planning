# ops/lock_deps.py
"""Regenerate hashed locks (all platforms) and download Windows wheels for offline install."""
import hashlib
import json
import subprocess
import sys
import tempfile
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
WHEELS = ROOT / 'vendor' / 'wheels'

def run(*a):
    subprocess.run(a, cwd=ROOT, check=True)

run('uv', 'pip', 'compile', 'requirements.in', '--python-version', '3.10', '--universal', '--generate-hashes', '-o', 'requirements.lock')
run('uv', 'pip', 'compile', 'requirements-dev.in', '--python-version', '3.10', '--universal', '--generate-hashes', '-o', 'requirements-dev.lock')
WHEELS.mkdir(parents=True, exist_ok=True)
for old in WHEELS.glob('*.whl'):
    old.unlink()
scratch = ROOT / 'results-private' / 'tmp'
scratch.mkdir(parents=True, exist_ok=True)
# pip evaluates markers on the host even when selecting another Python/platform.
# Vendor every universal-lock branch; retain markers in the actual install locks.
with tempfile.TemporaryDirectory(dir=scratch) as directory:
    download_lock = Path(directory) / 'download.lock'
    lines = []
    for line in (ROOT / 'requirements-dev.lock').read_text().splitlines():
        if not line.startswith((' ', '#')) and ';' in line:
            line = line.split(';', 1)[0].rstrip() + (' \\' if line.endswith('\\') else '')
        lines.append(line)
    download_lock.write_text('\n'.join(lines) + '\n')
    for pyver in ('3.10', '3.11', '3.12'):
        run(sys.executable, '-m', 'pip', 'download', '--no-deps', '--only-binary=:all:', '--platform', 'win_amd64',
            '--python-version', pyver, '--require-hashes', '-r', str(download_lock), '-d', str(WHEELS))
allow_path = ROOT / 'ops' / 'large_file_allowlist.json'
allow = {k: v for k, v in json.loads(allow_path.read_text()).items() if not k.startswith('vendor/wheels/')}
for w in sorted(WHEELS.glob('*.whl')):
    data = w.read_bytes()
    allow[w.relative_to(ROOT).as_posix()] = hashlib.sha256(data).hexdigest()
allow_path.write_text(json.dumps(allow, indent=2, sort_keys=True) + '\n')
print('locked; wheels:', len(list(WHEELS.glob('*.whl'))))
