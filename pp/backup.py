"""Verified SQLite backups via the online backup API (never a file copy of the live database) and restore."""
import argparse
import hashlib
import json
import os
import shutil
import sqlite3
import sys
import tempfile
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def _verify(con):
    try:
        result = con.execute('PRAGMA integrity_check').fetchall()
        if result != [('ok',)]:
            raise ValueError('integrity check failed: %s' % '; '.join(r[0] for r in result[:5]))
        if con.execute('PRAGMA foreign_key_check').fetchall():
            raise ValueError('foreign key check failed')
        return con.execute('PRAGMA user_version').fetchone()[0]
    except sqlite3.DatabaseError as e:
        raise ValueError('database is unreadable: %s' % e) from None


def backup(db_path, dest_dir):
    db_path, dest_dir = Path(db_path).resolve(), Path(dest_dir)
    if not db_path.is_file():
        raise ValueError('no database at %s' % db_path)
    dest_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    fd, tmp = tempfile.mkstemp(prefix='planning-%s-' % now.strftime('%Y%m%dT%H%M%S%fZ'), suffix='.partial', dir=dest_dir)
    os.close(fd)
    tmp = Path(tmp)
    target = tmp.with_suffix('.db')
    try:
        # Close handles before rename: Windows refuses to move an open database.
        with closing(sqlite3.connect(db_path.as_uri() + '?mode=ro', uri=True)) as src, \
                closing(sqlite3.connect(tmp)) as dst:
            src.backup(dst, pages=1024, sleep=0.05)
            dst.execute('PRAGMA journal_mode=DELETE')  # self-contained file, no -wal sidecar
            version = _verify(dst)
        digest = _sha256(tmp)
        os.replace(tmp, target)
        target.with_suffix('.json').write_text(json.dumps({
            'created_at': now.isoformat(), 'source': str(db_path), 'schema_version': version, 'sha256': digest,
        }, indent=2), encoding='utf-8')
        return target
    finally:
        tmp.unlink(missing_ok=True)


def restore(backup_path, new_dir):
    backup_path, new_dir = Path(backup_path), Path(new_dir)
    if new_dir.exists() and (not new_dir.is_dir() or any(new_dir.iterdir())):
        raise ValueError('restore target must be a new or empty folder: %s' % new_dir)
    sidecar = backup_path.with_suffix('.json')
    if not backup_path.is_file() or not sidecar.is_file():
        raise ValueError('backup or its .json record is missing: %s' % backup_path)
    expected = json.loads(sidecar.read_text(encoding='utf-8'))['sha256']
    if _sha256(backup_path) != expected:
        raise ValueError('backup checksum does not match its record; the file is damaged')
    new_dir.mkdir(parents=True, exist_ok=True)
    target = new_dir / 'planning.db'
    shutil.copyfile(backup_path, target)  # the backup is a closed file, not the live database
    try:
        if _sha256(target) != expected:
            raise ValueError('restored copy does not match the backup checksum')
        with closing(sqlite3.connect(target)) as con:
            _verify(con)
    except BaseException:
        target.unlink(missing_ok=True)
        raise
    return target


def main(argv=None):
    parser = argparse.ArgumentParser(prog='python -m pp.backup', description='Verified SQLite backup and restore.')
    default_db = Path(os.environ.get('DATA_DIR', ROOT / 'data')) / 'planning.db'
    parser.add_argument('--db', type=Path, default=default_db, help='live database (default: %(default)s)')
    parser.add_argument('--dest', type=Path, help='backup folder (default: <repo>/backups); with --restore, a new empty folder')
    parser.add_argument('--restore', type=Path, metavar='BACKUP', help='restore this backup file into --dest')
    parser.add_argument('--backup', action='store_true', help=argparse.SUPPRESS)  # run_v2 launchers pass it through
    args = parser.parse_args(argv)
    try:
        if args.restore:
            if not args.dest:
                parser.error('--restore needs --dest (a new empty folder)')
            print(restore(args.restore, args.dest))
        else:
            print(backup(args.db, args.dest or ROOT / 'backups'))
    except (ValueError, OSError) as e:
        print('Backup error: %s' % e, file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
