#!/bin/bash
# usage: sol.sh <name> <prompt-file>  -> runs Codex gpt-6.1-sol at HIGH effort in the repo root
set -u
cd "$(dirname "$0")/../../.."
D=.planning/excel-grid/sol
timeout 3600 codex exec --skip-git-repo-check -C . -s workspace-write \
  -c sandbox_workspace_write.network_access=true -c approval_policy=never \
  --model gpt-6.1-sol -c model_reasoning_effort=high \
  -o "$D/$1.md" "$(cat "$2")" > "$D/$1.log" 2>&1
echo "$1 exit=$?"
