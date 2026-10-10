"""Run the real ASGI app with a synthetic workbook for transport tests."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


if __name__ == "__main__":
    import uvicorn
    from tests.v2.api_support import APIFixture

    fixture = APIFixture()
    try:
        uvicorn.run(fixture.app, host="127.0.0.1", port=int(sys.argv[1]),
                    log_level="error", access_log=False, ws="websockets")
    finally:
        fixture.close()
