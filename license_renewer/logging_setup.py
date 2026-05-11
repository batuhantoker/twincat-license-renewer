"""Logging setup: INFO to stdout, DEBUG to a rolling file."""

from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def configure(log_file: Path, verbose: bool = False) -> logging.Logger:
    log_file.parent.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    for handler in list(root.handlers):
        root.removeHandler(handler)

    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG if verbose else logging.INFO)
    console.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    root.addHandler(console)

    file_handler = RotatingFileHandler(
        log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
    )
    root.addHandler(file_handler)

    return logging.getLogger("license_renewer")
