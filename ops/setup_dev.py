"""One-time per clone: point git at the repo-owned hooks."""
import subprocess
from pathlib import Path

root = Path(__file__).resolve().parent.parent
subprocess.run(['git', 'config', 'core.hooksPath', '.githooks'], cwd=root, check=True)
for hook in (root / '.githooks').iterdir():
    hook.chmod(0o755)
print('hooks installed: core.hooksPath=.githooks')
