#!/bin/sh
set -eu
cd "$(dirname "$0")"
PY=$(python3 ops/install_env.py)
exec "$PY" -u run.py --demo "$@"
