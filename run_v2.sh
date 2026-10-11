#!/bin/sh
set -eu
cd "$(dirname "$0")"
# Same order as run_v2.bat: approved local Python, then python3, then python (3.10+).
for candidate in python/bin/python3 python3 python; do
  if command -v "$candidate" >/dev/null 2>&1 &&
     "$candidate" -c 'import sys; sys.exit(not (sys.version_info >= (3, 10)))' 2>/dev/null; then
    PY=$("$candidate" ops/install_env.py)
    case "${1:-}" in
      --backup|--restore) exec "$PY" -m pp.backup "$@" ;;
    esac
    exec "$PY" -u run_v2.py "$@"
  fi
done
echo "Python 3.10 or newer is required." >&2
exit 1
