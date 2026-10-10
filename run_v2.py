"""Opt-in Starlette/Uvicorn entry point for the Track A backend."""

import argparse
import logging
from pathlib import Path

import uvicorn

from pp.config import load_config
from pp.runtime import create_runtime
from pp.template.importer import import_template


ROOT = Path(__file__).resolve().parent


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Price Planning v2 backend")
    parser.add_argument("--import-template", action="store_true",
                        help="Import or reuse the configured template snapshot, then exit")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    if args.import_template:
        config = load_config(ROOT)
        if config.template_path is None or not config.template_path.is_file():
            raise ValueError("template_path must name an existing .xlsm file")
        import_template(config.template_path, config.data_dir)
        return 0
    runtime = create_runtime(ROOT)
    try:
        uvicorn.run(runtime.app, host=runtime.config.bind, port=runtime.config.port,
                    loop="asyncio", ws="websockets", workers=1)
        return 0
    finally:
        runtime.close()


if __name__ == "__main__":
    raise SystemExit(main())
