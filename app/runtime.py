"""Direct Windows/Linux hosting, interactive setup, and safe SQLite backups."""
import argparse
from contextlib import closing
import getpass
import hashlib
import ipaddress
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import socket
import sqlite3
import sys
import tempfile
import time
import urllib.parse
import webbrowser

ROOT = Path(__file__).resolve().parents[1]
WHEEL = ROOT / 'vendor/waitress-3.0.2-py3-none-any.whl'
WHEEL_SHA256 = 'c56d67fd6e87c2ee598b76abdd4e96cfad1f24cacdea5078d382b1f9d7b5ed2e'
# The team server is meant to be reachable from the office network, so binding every
# interface is the supported deployment, not an oversight. Reaching it still requires
# a signed-in account, and `configure` asks for it explicitly. Keeping the literal in
# one named constant also stops bandit's B104 from firing on every comparison below.
ALL_INTERFACES = '0.0.0.0'  # nosec B104
DEFAULTS = dict(bind=ALL_INTERFACES, port=8000, public_origin='auto', demo_mode=True,
                cookie_secure=False, threads=8, data_dir='data')


def load_waitress():
    # No pip, virtual environment, global installation, or network is needed.
    if not WHEEL.is_file() or hashlib.sha256(WHEEL.read_bytes()).hexdigest() != WHEEL_SHA256:
        raise ValueError('The bundled web server is missing or changed. Extract a fresh application package.')
    if str(WHEEL) not in sys.path:
        sys.path.insert(0, str(WHEEL))
    from waitress.server import create_server
    return create_server


def validate_settings(settings):
    cfg = DEFAULTS | settings
    ipaddress.ip_address(cfg['bind'])
    if type(cfg['port']) is not int or not 1024 <= cfg['port'] <= 65535:
        raise ValueError('port must be a whole number from 1024 to 65535.')
    if type(cfg['threads']) is not int or not 2 <= cfg['threads'] <= 64:
        raise ValueError('threads must be a whole number from 2 to 64.')
    if type(cfg['cookie_secure']) is not bool:
        raise ValueError('cookie_secure must be true or false.')
    if type(cfg['demo_mode']) is not bool:
        raise ValueError('demo_mode must be true or false.')
    if not isinstance(cfg['public_origin'], str):
        raise ValueError('public_origin must be a URL.')
    if cfg['public_origin'] == 'auto':
        if not cfg['demo_mode'] or cfg['cookie_secure']:
            raise ValueError('Automatic LAN addresses require demo_mode true and cookie_secure false.')
    else:
        origin = urllib.parse.urlsplit(cfg['public_origin'])
        if origin.scheme not in ('http', 'https') or not origin.hostname or origin.username or origin.password or origin.path not in ('', '/') or origin.query or origin.fragment:
            raise ValueError('public_origin must be an HTTP/HTTPS address without a path, credentials, or query.')
        if origin.hostname in (ALL_INTERFACES, '::'):
            raise ValueError('public_origin must use the server name or IP address that people open in their browser.')
        if cfg['cookie_secure'] != (origin.scheme == 'https'):
            raise ValueError('cookie_secure must be true for HTTPS and false for HTTP.')
        if origin.scheme == 'http' and (origin.port or 80) != cfg['port']:
            raise ValueError('For direct HTTP access, public_origin must use the listening port.')
    if not isinstance(cfg['data_dir'], str) or not cfg['data_dir'].strip():
        raise ValueError('data_dir must be a local folder path.')
    if cfg['data_dir'].startswith(('\\\\', '//')):
        raise ValueError('Store the database on this server locally, rather than in a network share.')
    cfg['public_origin'] = cfg['public_origin'].rstrip('/')
    return cfg


def read_settings(path):
    with path.open(encoding='utf-8-sig') as f:
        settings = json.load(f)
    if not isinstance(settings, dict):
        raise ValueError('The settings file must contain a JSON object.')
    return validate_settings(settings)


def write_settings(path, settings):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=path.name + '.', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
            f.write('\n')
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def configure(path):
    current = read_settings(path) if path.exists() else DEFAULTS.copy()
    print('\nAccess settings: 1 = this computer, 2 = team LAN, 3 = existing HTTPS proxy')
    choice = input('Choose access [1]: ').strip() or '1'
    if choice not in ('1', '2', '3'):
        raise ValueError('Choose 1, 2, or 3.')
    port = int(input('Server port [%s]: ' % current['port']).strip() or current['port'])
    if choice == '1':
        bind, origin = '127.0.0.1', 'http://localhost:%s' % port
    elif choice == '2':
        bind = ALL_INTERFACES
        origin = input('Team address, for example http://10.0.0.25:%s: ' % port).strip()
        if not origin.startswith('http://'):
            raise ValueError('Use an http:// server address for direct LAN access.')
        print('Team access uses your existing company network permissions. No firewall rules are changed.')
    else:
        bind = '127.0.0.1'
        origin = input('Existing HTTPS address, for example https://planning.company.com: ').strip()
        if not origin.startswith('https://'):
            raise ValueError('Use the https:// address served by your existing proxy.')
    demo = (input('Let people create their own demo accounts? [Y/n]: ').strip().lower() or 'y') == 'y'
    cfg = validate_settings(current | dict(bind=bind, port=port, public_origin=origin,
                                            cookie_secure=choice == '3', demo_mode=demo))
    write_settings(path, cfg)
    print('Saved %s. No passwords are stored in this file.' % path.name)
    return cfg


def lan_addresses():
    """Local hostname resolution only; prefer the company's 10.x interface."""
    candidates = set()
    try:
        for result in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET, socket.SOCK_STREAM):
            addr = ipaddress.ip_address(result[4][0])
            if not (addr.is_loopback or addr.is_link_local or addr.is_multicast or addr.is_unspecified):
                candidates.add(str(addr))
    except OSError:
        pass
    # UDP connect only selects a local route; no datagram is sent.
    if not candidates:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as route:
                route.connect(('192.0.2.1', 9))
                addr = ipaddress.ip_address(route.getsockname()[0])
                if not (addr.is_loopback or addr.is_link_local or addr.is_unspecified):
                    candidates.add(str(addr))
        except OSError:
            pass
    return sorted(candidates, key=lambda ip: (not ip.startswith('10.'), not ipaddress.ip_address(ip).is_private, int(ipaddress.ip_address(ip))))


def resolve_addresses(cfg):
    """Resolve on each startup so changing server IPs are not saved forever."""
    if cfg['public_origin'] != 'auto':
        return cfg['public_origin'], [cfg['public_origin']]
    interfaces = lan_addresses() if cfg['bind'] == ALL_INTERFACES else []
    local = ['http://localhost:%s' % cfg['port'], 'http://127.0.0.1:%s' % cfg['port']]
    team = ['http://%s:%s' % (ip, cfg['port']) for ip in interfaces]
    origins = team + local
    # The server hostname is useful when clients use internal DNS instead of an IP.
    if interfaces and re_safe_hostname(socket.gethostname()):
        origins.append('http://%s:%s' % (socket.gethostname().lower(), cfg['port']))
    return (team[0] if team else local[0]), origins


def re_safe_hostname(host):
    return bool(host) and all(c.isascii() and (c.isalnum() or c in '.-') for c in host)


def resolve_data_dir(cfg):
    """DATA_DIR env beats settings data_dir; relative paths are under the repo root."""
    data = Path(os.environ.get('DATA_DIR', cfg['data_dir'])).expanduser()
    if not data.is_absolute():
        data = ROOT / data
    return data.resolve()


def apply_environment(cfg):
    data = resolve_data_dir(cfg)
    data.mkdir(parents=True, exist_ok=True)
    os.environ['DATA_DIR'] = str(data)
    origin, origins = resolve_addresses(cfg)
    os.environ['PUBLIC_ORIGIN'] = origin
    os.environ['PUBLIC_ORIGINS'] = ','.join(origins)
    os.environ['PLANNING_DEMO_MODE'] = '1' if cfg['demo_mode'] else '0'
    os.environ['COOKIE_SECURE'] = '1' if cfg['cookie_secure'] else '0'
    return data


class ServerLock:
    """OS-owned lock, released even after a crash; the file is never deleted."""
    def __init__(self, data):
        self.path = data / '.server.lock'
        self.file = None

    def __enter__(self):
        self.file = self.path.open('a+b')
        self.file.seek(0, 2)
        if not self.file.tell():
            self.file.write(b'0')
            self.file.flush()
        self.file.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.file.close()
            self.file = None
            raise ValueError('This database is already running. Use the existing server window; stop it before changing settings.') from exc
        return self

    def __exit__(self, *args):
        if self.file:
            self.file.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.file.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.file, fcntl.LOCK_UN)
            self.file.close()


def backup_database(data, destination):
    if os.environ.get('DATABASE_URL'):
        raise ValueError('This backup command is for SQLite. Use pg_dump for an externally configured PostgreSQL database.')
    source = data / 'planning.db'
    if not source.is_file():
        raise ValueError('No planning database exists yet. Start the app once before creating a backup.')
    destination.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime('%Y%m%dT%H%M%SZ', time.gmtime())
    fd, temporary = tempfile.mkstemp(prefix='planning-' + stamp + '-', suffix='.partial', dir=destination)
    os.close(fd)
    temporary = Path(temporary)
    target = temporary.with_suffix('.db')
    try:
        # Explicitly close handles before rename: Windows rejects moving an open DB.
        with closing(sqlite3.connect(source.as_uri() + '?mode=ro', uri=True)) as src, closing(sqlite3.connect(temporary)) as dst:
            src.backup(dst, pages=1024, sleep=0.05)
            result = dst.execute('PRAGMA quick_check').fetchone()[0]
            if result != 'ok':
                raise ValueError('Backup verification failed: %s' % result)
        os.replace(temporary, target)
        try:
            target.chmod(0o600)
        except OSError:
            pass
        return target
    finally:
        temporary.unlink(missing_ok=True)


def initialize():
    # Import these only after the data-folder setting has been applied.
    from .db import DB, init_schema
    from .server import create_admin
    from .workbook import import_workbook
    credentials = None
    with DB() as db:
        init_schema(db)
        if os.environ.get('PLANNING_DEMO_MODE') != '1' and not db.one('SELECT id FROM users WHERE admin=1 LIMIT 1'):
            print('\nCreate the first administrator account.')
            email = input('Administrator email: ').strip().lower()
            name = input('Administrator name: ').strip()
            password = getpass.getpass('Administrator password (12+ characters): ')
            confirm = getpass.getpass('Repeat password: ')
            if '@' not in email or len(email) > 254 or any(c.isspace() for c in email) or not name or len(name) > 100:
                raise ValueError('Enter a valid email and a name of 1 to 100 characters.')
            if len(password) < 12 or password != confirm:
                raise ValueError('Passwords must match and contain at least 12 characters.')
            if db.one('SELECT id FROM users WHERE email=?', (email,)):
                raise ValueError('That email already belongs to a non-administrator account. Use a different email.')
            credentials = (email, name, password)
        print('Checking workbook import. The first import takes several minutes; later starts reuse it.', flush=True)
        import_workbook(ROOT / 'Price Planning Template Spring 2027.xlsm', db)
    if credentials:
        create_admin(*credentials)
        print('Administrator created. The database stores a password hash, not the password.')


def configure_logging(data):
    logs = data / 'logs'
    logs.mkdir(exist_ok=True)
    handler = RotatingFileHandler(logs / 'server.log', maxBytes=2 * 1024 * 1024,
                                  backupCount=3, encoding='utf-8')
    logging.basicConfig(level=logging.INFO, handlers=[logging.StreamHandler(), handler],
                        format='%(asctime)s %(levelname)s %(name)s: %(message)s')


def main(argv=None):
    parser = argparse.ArgumentParser(description='Catalyst Brands price planning server')
    parser.add_argument('--settings', type=Path, default=ROOT / 'settings.json')
    parser.add_argument('--configure', action='store_true', help='Change access settings and exit')
    parser.add_argument('--demo', action='store_true', help='Start a LAN demo without administrator setup')
    parser.add_argument('--init-only', action='store_true', help='Initialize the workbook, then exit')
    parser.add_argument('--no-browser', action='store_true', help='Start without opening a browser')
    parser.add_argument('--backup', action='store_true', help='Create a verified online SQLite backup and exit')
    parser.add_argument('--backup-dir', type=Path, default=ROOT / 'backups')
    parser.add_argument('--source-refresh', action='store_true', help='Run the administrator source import with the remaining arguments')
    args, rest = parser.parse_known_args(argv)
    if rest and not args.source_refresh:
        parser.error('Unrecognized arguments: ' + ' '.join(rest))
    if sum((args.configure, args.init_only, args.backup, args.source_refresh)) > 1:
        parser.error('Choose only one of --configure, --init-only, --backup, and --source-refresh.')
    try:
        created = not args.settings.exists()
        if not created:
            cfg = read_settings(args.settings)
        elif args.backup or args.source_refresh:
            raise ValueError('Start the app once to create settings.json before this operation.')
        elif args.configure:
            cfg = configure(args.settings)
        else:
            cfg = validate_settings(DEFAULTS)
        if args.demo and not (args.configure or args.backup or args.source_refresh):
            # Preserve a deliberately configured LAN URL. Older loopback setups
            # become reachable when launched through the demo batch file.
            origin = cfg['public_origin'] if cfg['demo_mode'] and cfg['bind'] == ALL_INTERFACES else 'auto'
            cfg = validate_settings(cfg | dict(bind=ALL_INTERFACES, public_origin=origin,
                                               demo_mode=True, cookie_secure=False))
        data = apply_environment(cfg)
        if args.backup:
            print('Creating an online backup using the SQLite backup API...', flush=True)
            print('Verified backup: %s' % backup_database(data, args.backup_dir.resolve()))
            return 0
        if args.source_refresh:
            from .source_refresh import main as refresh
            refresh(rest)
            return 0
        with ServerLock(data):
            if args.configure:
                if not created:
                    configure(args.settings)
                return 0
            if created:
                write_settings(args.settings, cfg)
            configure_logging(data)
            server = None
            if not args.init_only:
                # Bind before the expensive import so a port collision fails immediately.
                from .server import application
                server = load_waitress()(application, host=cfg['bind'], port=cfg['port'],
                    threads=cfg['threads'], connection_limit=256, channel_timeout=65,
                    max_request_body_size=2 * 1024 * 1024, clear_untrusted_proxy_headers=True,
                    ident='Catalyst Planning', asyncore_loop_timeout=0.25)
            try:
                initialize()
                if args.init_only:
                    print('Initialization complete. Start with run.bat or python run.py.')
                    return 0
                address = os.environ['PUBLIC_ORIGIN']
                print('\nREADY: %s' % address, flush=True)
                if cfg['demo_mode']:
                    print('DEMO: no administrator account required. Users choose Create demo account in the browser.')
                    print('On this server: http://localhost:%s' % cfg['port'])
                    team = [x for x in os.environ['PUBLIC_ORIGINS'].split(',') if 'localhost' not in x and '127.0.0.1' not in x]
                    if team:
                        print('Share with your team: ' + address)
                        if len(team) > 1:
                            print('Other server addresses: ' + ', '.join(team[1:]))
                    else:
                        print('No LAN address detected. Run ipconfig and share http://YOUR_SERVER_IP:%s.' % cfg['port'])
                print('Keep this window open. Press Ctrl+C to stop.\n', flush=True)
                logging.info('Planning server started on %s:%s', cfg['bind'], cfg['port'])
                if not args.no_browser:
                    try:
                        webbrowser.open(address)
                    except Exception:
                        logging.warning('Could not open a browser. Open the printed address manually.')
                server.run()
            finally:
                if server:
                    server.task_dispatcher.shutdown()
                    from waitress import wasyncore
                    wasyncore.close_all(map=server._map)
                logging.info('Planning server stopped')
        return 0
    except KeyboardInterrupt:
        print('\nStopped. Saved inputs remain in the database.')
        return 0
    except (ValueError, OSError, EOFError) as exc:
        print('\nCould not start: %s' % exc, file=sys.stderr)
        print('Correct the settings or permissions, then run again. Existing data is preserved.', file=sys.stderr)
        return 1
    except Exception:
        logging.exception('Startup failed')
        print('Startup failed. See data/logs/server.log for details.', file=sys.stderr)
        return 1
