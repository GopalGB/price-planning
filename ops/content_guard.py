"""Block company data from git: office/spreadsheet files, big files, denylisted strings."""
import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BLOCKED_EXT = {'.xlsm', '.xlsx', '.xls', '.xlsb', '.csv', '.docx', '.doc', '.pptx', '.pdf', '.db', '.bas'}
MAX_BYTES = 5_000_000


def check(paths, read, deny, allow):
    out = []
    deny = [item.lower() for item in deny if item.strip()]
    for path in paths:
        if Path(path).suffix.lower() in BLOCKED_EXT:
            out.append(f'BLOCKED {path}: office/spreadsheet file type')
            continue
        data = read(path)
        if len(data) > MAX_BYTES:
            digest = hashlib.sha256(data).hexdigest()
            if allow.get(path) != digest:
                out.append(f'BLOCKED {path}: {len(data)} bytes > {MAX_BYTES} and not allowlisted with sha256 {digest}')
                continue
        if deny and b'\0' not in data[:8192]:
            text = data.decode('utf-8', 'ignore').lower()
            for item in deny:
                if item in text:
                    out.append(f'BLOCKED {path}: contains denylisted text')
                    break
    return out


def _git(*args):
    return subprocess.run(['git', *args], cwd=ROOT, capture_output=True, check=True).stdout


def main(argv=None):
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--staged', action='store_true')
    mode.add_argument('--all', action='store_true')
    args = parser.parse_args(argv)
    if args.staged:
        paths = [path for path in _git('diff', '--cached', '--name-only', '--diff-filter=ACMR', '-z').decode().split('\0') if path]
        read = lambda path: _git('show', f':{path}')
    else:
        paths = [path for path in _git('ls-files', '-z').decode().split('\0') if path]
        read = lambda path: (ROOT / path).read_bytes()
    deny_file = ROOT / '.content-denylist'
    deny = deny_file.read_text(encoding='utf-8').splitlines() if deny_file.exists() else []
    allow = json.loads((ROOT / 'ops' / 'large_file_allowlist.json').read_text())
    bad = check(paths, read, deny, allow)
    for line in bad:
        print(line)
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
