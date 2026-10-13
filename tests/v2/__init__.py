"""Track A v2 backend contract tests."""

import atexit
import os
import shutil
import tempfile

# Engine packages the tests prepare go to a throwaway folder, not the cache a running server uses
if "PP_ENGINE_CACHE_DIR" not in os.environ:
    os.environ["PP_ENGINE_CACHE_DIR"] = tempfile.mkdtemp(prefix="pp-test-engine-cache-")
    atexit.register(shutil.rmtree, os.environ["PP_ENGINE_CACHE_DIR"], True)
