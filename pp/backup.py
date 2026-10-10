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
import re

from .config import resolve_database_path

ROOT = Path(__file__).resolve().parent.parent
FORMAT_VERSION = 2
TEMPLATE_ID = re.compile(r'[0-9a-f]{64}')


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


def _template_rows(database):
    with closing(sqlite3.connect(database.as_uri() + '?mode=ro', uri=True)) as connection:
        connection.row_factory = sqlite3.Row
        exists = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='pp_templates'"
        ).fetchone()
        if exists is None:
            return []
        return connection.execute(
            'SELECT id,source_path,snapshot_path FROM pp_templates ORDER BY id'
        ).fetchall()


def _snapshot_hashes(root):
    if not root.is_dir() or root.is_symlink():
        raise ValueError('registered template snapshot is missing or unsafe: %s' % root)
    hashes = {}
    for directory, names, files in os.walk(root, followlinks=False):
        directory = Path(directory)
        for name in names:
            if (directory / name).is_symlink():
                raise ValueError('registered template snapshot contains a symlink')
        for name in files:
            path = directory / name
            if path.is_symlink() or not path.is_file():
                raise ValueError('registered template snapshot contains an unsafe file')
            hashes[path.relative_to(root).as_posix()] = _sha256(path)
    return dict(sorted(hashes.items()))


def _validate_template(row):
    template_id = row['id']
    if not isinstance(template_id, str) or TEMPLATE_ID.fullmatch(template_id) is None:
        raise ValueError('registered template source id is invalid')
    source, snapshot = Path(row['source_path']), Path(row['snapshot_path'])
    if not source.is_file() or source.is_symlink() or _sha256(source) != template_id:
        raise ValueError('registered template source id does not match its file')
    hashes = _snapshot_hashes(snapshot)
    if hashes.get('source.xlsm') != template_id:
        raise ValueError('registered template snapshot source id does not match')
    try:
        metadata = json.loads((snapshot / 'meta.json').read_text(encoding='utf-8'))
    except (OSError, ValueError):
        raise ValueError('registered template snapshot metadata is unreadable') from None
    if metadata.get('hash') != template_id:
        raise ValueError('registered template snapshot metadata id does not match')
    return template_id, snapshot, hashes


def _copy_snapshot(source, destination, hashes):
    for relative, expected in hashes.items():
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source / relative, target)
        if _sha256(target) != expected:
            raise ValueError('copied template checksum does not match its source')


def _capture_templates(database, dest_dir, stem):
    rows = _template_rows(database)
    if not rows:
        return [], None
    staging = Path(tempfile.mkdtemp(prefix='.%s-templates-' % stem, dir=dest_dir))
    entries = []
    try:
        for row in rows:
            template_id, snapshot, hashes = _validate_template(row)
            _copy_snapshot(snapshot, staging / template_id, hashes)
            entries.append({'id': template_id, 'files': hashes})
        return entries, staging
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise


def _safe_relative(value):
    if not isinstance(value, str) or not value or '\\' in value or ':' in value:
        raise ValueError('template manifest contains an unsafe path')
    path = Path(value)
    if path.is_absolute() or any(part in {'', '.', '..'} for part in path.parts):
        raise ValueError('template manifest contains an unsafe path')
    if path.as_posix() != value:
        raise ValueError('template manifest contains an unsafe path')
    return path


def _manifest_templates(backup_path, record):
    version = record.get('format_version')
    if version is None:
        return None
    if version != FORMAT_VERSION or not isinstance(record.get('templates'), list):
        raise ValueError('backup template manifest version is unsupported')
    entries, identifiers = [], set()
    for raw in record['templates']:
        if not isinstance(raw, dict) or set(raw) != {'id', 'files'}:
            raise ValueError('backup template manifest is invalid')
        template_id, files = raw['id'], raw['files']
        if (not isinstance(template_id, str) or TEMPLATE_ID.fullmatch(template_id) is None
                or template_id in identifiers or not isinstance(files, dict)):
            raise ValueError('backup template manifest is invalid')
        checked = {}
        for relative, digest in files.items():
            checked[_safe_relative(relative).as_posix()] = digest
            if not isinstance(digest, str) or TEMPLATE_ID.fullmatch(digest) is None:
                raise ValueError('backup template manifest checksum is invalid')
        if checked.get('source.xlsm') != template_id or 'meta.json' not in checked:
            raise ValueError('backup template manifest source id is invalid')
        identifiers.add(template_id)
        entries.append({'id': template_id, 'files': checked})
    bundle = backup_path.with_suffix('.templates')
    if not entries:
        if bundle.exists():
            raise ValueError('backup has an unexpected template bundle')
        return []
    if not bundle.is_dir() or bundle.is_symlink():
        raise ValueError('backup template bundle is missing or unsafe')
    if {path.name for path in bundle.iterdir()} != identifiers:
        raise ValueError('backup template bundle does not match its manifest')
    for entry in entries:
        actual = _snapshot_hashes(bundle / entry['id'])
        if set(actual) != set(entry['files']):
            raise ValueError('backup template bundle is missing a template file')
        if actual != entry['files']:
            raise ValueError('backup template file checksum does not match its manifest')
        metadata = json.loads((bundle / entry['id'] / 'meta.json').read_text(encoding='utf-8'))
        if metadata.get('hash') != entry['id']:
            raise ValueError('backup template metadata id does not match')
    return entries


def _stage_templates(backup_path, staging, entries):
    if not entries:
        return
    bundle = backup_path.with_suffix('.templates')
    for entry in entries:
        source = bundle / entry['id']
        destination = staging / 'templates' / entry['id']
        _copy_snapshot(source, destination, entry['files'])


def _publish_templates(staging, new_dir):
    source, target = staging / 'templates', new_dir / 'templates'
    try:
        os.mkdir(target)
    except FileExistsError:
        raise ValueError('%s already exists; nothing was overwritten' % target) from None
    try:
        for path in sorted(source.rglob('*')):
            destination = target / path.relative_to(source)
            if path.is_dir():
                os.mkdir(destination)
            else:
                os.link(path, destination)
    except Exception:
        shutil.rmtree(target, ignore_errors=True)
        raise


def _relocate_templates(database, new_dir, entries):
    with closing(sqlite3.connect(database)) as connection:
        table = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='pp_templates'"
        ).fetchone()
        actual = ({row[0] for row in connection.execute('SELECT id FROM pp_templates')}
                  if table is not None else set())
        expected = {entry['id'] for entry in entries}
        if actual != expected:
            raise ValueError('backup template manifest does not match the database')
        root = new_dir.resolve() / 'templates'
        for template_id in sorted(expected):
            snapshot = root / template_id
            connection.execute(
                'UPDATE pp_templates SET source_path=?,snapshot_path=? WHERE id=?',
                (str(snapshot / 'source.xlsm'), str(snapshot), template_id),
            )
        connection.commit()
        _verify(connection)


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
    record_tmp = target.with_suffix('.json.partial')
    assets = target.with_suffix('.templates')
    staged_assets = None
    published = []
    try:
        # Close handles before rename: Windows refuses to move an open database.
        with closing(sqlite3.connect(db_path.as_uri() + '?mode=ro', uri=True)) as src, \
                closing(sqlite3.connect(tmp)) as dst:
            src.backup(dst, pages=1024, sleep=0.05)
            dst.execute('PRAGMA journal_mode=DELETE')  # self-contained file, no -wal sidecar
            version = _verify(dst)
        digest = _sha256(tmp)
        templates, staged_assets = _capture_templates(tmp, dest_dir, target.stem)
        record_tmp.write_text(json.dumps({
            'format_version': FORMAT_VERSION,
            'created_at': now.isoformat(),
            'source': str(db_path),
            'schema_version': version,
            'sha256': digest,
            'templates': templates,
        }, indent=2), encoding='utf-8')
        if staged_assets is not None:
            os.rename(staged_assets, assets)
            staged_assets = None
            published.append(assets)
        os.replace(tmp, target)
        published.append(target)
        os.replace(record_tmp, target.with_suffix('.json'))
        published.append(target.with_suffix('.json'))
        return target
    except Exception:
        for path in reversed(published):
            shutil.rmtree(path, ignore_errors=True) if path.is_dir() else path.unlink(missing_ok=True)
        raise
    finally:
        tmp.unlink(missing_ok=True)
        record_tmp.unlink(missing_ok=True)
        if staged_assets is not None:
            shutil.rmtree(staged_assets, ignore_errors=True)


def restore(backup_path, new_dir):
    backup_path, new_dir = Path(backup_path), Path(new_dir)
    if new_dir.is_symlink():
        raise ValueError('restore target must not be a symlink: %s' % new_dir)
    if new_dir.exists() and (not new_dir.is_dir() or any(new_dir.iterdir())):
        raise ValueError('restore target must be a new or empty folder: %s' % new_dir)
    sidecar = backup_path.with_suffix('.json')
    if not backup_path.is_file() or not sidecar.is_file():
        raise ValueError('backup or its .json record is missing: %s' % backup_path)
    record = json.loads(sidecar.read_text(encoding='utf-8'))
    expected = record['sha256']
    if _sha256(backup_path) != expected:
        raise ValueError('backup checksum does not match its record; the file is damaged')
    templates = _manifest_templates(backup_path, record)
    new_dir.mkdir(parents=True, exist_ok=True)
    target = new_dir / 'planning.db'
    staging = Path(tempfile.mkdtemp(prefix='.restore-', dir=new_dir))
    staged_database = staging / 'planning.db'
    published_templates = False
    try:
        shutil.copyfile(backup_path, staged_database)  # the backup is closed, not the live database
        if _sha256(staged_database) != expected:
            raise ValueError('restored copy does not match the backup checksum')
        _stage_templates(backup_path, staging, templates or [])
        with closing(sqlite3.connect(staged_database.as_uri() + '?mode=ro', uri=True)) as con:
            _verify(con)
        if templates is not None:
            _relocate_templates(staged_database, new_dir, templates)
        if templates:
            _publish_templates(staging, new_dir)
            published_templates = True
        try:
            os.link(staged_database, target)  # atomic publish that never overwrites (NTFS, APFS, ext4)
        except FileExistsError:
            raise ValueError('%s already exists; nothing was overwritten' % target) from None
    except Exception:
        if published_templates:
            shutil.rmtree(new_dir / 'templates', ignore_errors=True)
        raise
    finally:
        shutil.rmtree(staging, ignore_errors=True)
    return target


def default_db(settings_path=ROOT / 'settings.json'):
    """Same database as the server, after an explicit --db override is considered."""
    return resolve_database_path(ROOT, settings_path).resolve()


def main(argv=None):
    parser = argparse.ArgumentParser(prog='python -m pp.backup', description='Verified SQLite backup and restore.')
    parser.add_argument('--db', type=Path, help='live database (highest precedence; otherwise PP_DATABASE_PATH, '
                        'DATA_DIR/planning.db, settings database_path, settings data_dir/planning.db, repo/data/planning.db)')
    parser.add_argument('--settings', type=Path, default=ROOT / 'settings.json', help='settings file (default: %(default)s)')
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
            print(backup(args.db or default_db(args.settings), args.dest or ROOT / 'backups'))
    except (ValueError, OSError) as e:
        print('Backup error: %s' % e, file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
