"""Block company data from git: office/spreadsheet files, big files, denylisted strings."""
import argparse
import codecs
import fnmatch
import hashlib
import io
import json
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BLOCKED_EXT = {'.xlsm', '.xlsx', '.xls', '.xlsb', '.csv', '.docx', '.doc', '.pptx', '.pdf', '.db', '.bas'}
MAX_BYTES = 5_000_000
SIGNATURES = (b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1', b'%PDF-', b'SQLite format 3\0')
PRIVATE_DIRS = {'results-private', 'data', 'backups', 'logs'}
PRIVATE_NAMES = ('*.log', '*.db*', '*.bas', '.env*')


def private_path(path):
    parts = path.replace('\\', '/').casefold().split('/')
    return bool(PRIVATE_DIRS.intersection(parts[:-1])) or any(
        fnmatch.fnmatchcase(part, pattern) for part in parts for pattern in PRIVATE_NAMES)


def binary_kind(data):
    if data.startswith(SIGNATURES):
        return 'Office/OLE, PDF or database content'
    stream = io.BytesIO(data)
    if data.startswith((b'PK\x03\x04', b'PK\x05\x06', b'PK\x07\x08')) or zipfile.is_zipfile(stream):
        try:
            with zipfile.ZipFile(stream) as archive:
                members = {name.replace('\\', '/').lower() for name in archive.namelist()}
                if members & {'xl/workbook.xml', 'word/document.xml', 'ppt/presentation.xml'}:
                    return 'Office ZIP content'
        except zipfile.BadZipFile:
            return 'malformed ZIP content'
        return 'ZIP content'
    return None


def decode_text(data):
    # UTF-32 BOMs share their first two bytes with UTF-16 BOMs.
    if data.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
        return data.decode('utf-32')
    if data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return data.decode('utf-16')
    text = data.decode('utf-8-sig')
    if '\0' in text:
        raise ValueError('unsupported binary or text encoding')
    return text


def check(paths, read, deny, allow):
    out = []
    deny = [item.strip().casefold() for item in deny if item.strip()]
    for path in paths:
        if private_path(path):
            out.append(f'BLOCKED {path}: private/runtime path')
            continue
        if Path(path).suffix.lower() in BLOCKED_EXT:
            out.append(f'BLOCKED {path}: office/spreadsheet file type')
            continue
        data = read(path)
        vendor_wheel = path.startswith('vendor/') and path.endswith('.whl')
        digest = hashlib.sha256(data).hexdigest()
        if vendor_wheel and allow.get(path) == digest:
            continue
        if len(data) > MAX_BYTES:
            out.append(f'BLOCKED {path}: {len(data)} bytes > {MAX_BYTES} and not allowlisted with sha256 {digest}')
            continue
        kind = binary_kind(data)
        if kind:
            out.append(f'BLOCKED {path}: {kind}')
            continue
        try:
            text = decode_text(data).casefold()
        except (UnicodeError, ValueError):
            out.append(f'BLOCKED {path}: unsupported binary or text encoding')
            continue
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
        paths = [path for path in _git('diff', '--cached', '--name-only', '--diff-filter=d', '-z').decode().split('\0') if path]
        read = lambda path: _git('show', f':{path}')
    else:
        # tracked files deleted in the worktree (e.g. replaced bundle chunks) are not content
        paths = [path for path in _git('ls-files', '-z').decode().split('\0') if path and (ROOT / path).exists()]
        read = lambda path: (ROOT / path).read_bytes()
    deny_file = ROOT / '.content-denylist'
    deny = deny_file.read_text(encoding='utf-8-sig').splitlines() if deny_file.exists() else []
    allow = json.loads((ROOT / 'ops' / 'large_file_allowlist.json').read_text())
    bad = check(paths, read, deny, allow)
    for line in bad:
        print(line)
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
