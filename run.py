"""Portable entry point using the bundled, unmodified Waitress wheel."""
import sys

def main(argv=None):
    if sys.version_info < (3, 10):
        print('Python 3.10 or newer is required. Python 3.12 is recommended.')
        return 1
    from app.runtime import main as start
    return start(argv)

if __name__ == '__main__':
    raise SystemExit(main())
