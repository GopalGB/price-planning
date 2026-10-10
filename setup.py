"""Optional settings wizard. Normal first-time setup is included in run.bat."""
from run import main

if __name__ == '__main__':
    raise SystemExit(main(['--configure']))
