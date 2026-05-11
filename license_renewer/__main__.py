"""CLI entrypoint: ``python -m license_renewer`` or ``renew-license``."""

from __future__ import annotations

import argparse
import ctypes
import sys


def _enable_dpi_awareness() -> None:
    """Make this process per-monitor DPI-aware so pywinauto's screen coords
    match ``PIL.ImageGrab``. Must run before any GUI or screen-capture code.
    """
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_enable_dpi_awareness()


def _load_dotenv() -> None:
    """Load a ``.env`` file from the current directory, if one exists.
    Values already set in the real environment take precedence.
    """
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    load_dotenv(override=False)


_load_dotenv()

from . import config as config_mod  # noqa: E402
from . import flow, logging_setup  # noqa: E402
from .flow import FlowError, OcrExhausted  # noqa: E402
from .xae import XaeShellNotFound  # noqa: E402


EXIT_OK = 0
EXIT_XAE_MISSING = 2
EXIT_UI_NOT_FOUND = 3
EXIT_OCR_EXHAUSTED = 4
EXIT_UNEXPECTED = 5


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="renew-license",
        description="Activate the TwinCAT 3 XAE Shell 7-day trial license.",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Print DEBUG logs to the console."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Attach to XAE Shell and verify it's reachable, but don't click anything.",
    )
    parser.add_argument(
        "--project",
        metavar="PATH",
        help=(
            "Path to a .sln or .tsproj to open in XAE Shell. If omitted and no "
            "project is loaded, the first entry in 'File > Recent Projects and "
            "Solutions' is opened."
        ),
    )
    args = parser.parse_args(argv)

    cfg = config_mod.load()
    log = logging_setup.configure(cfg.log_file, verbose=args.verbose)
    log.debug("Config: %s", cfg)

    try:
        message = flow.run(cfg, dry_run=args.dry_run, project_path=args.project)
    except XaeShellNotFound as exc:
        log.error("XAE Shell not found: %s", exc)
        return EXIT_XAE_MISSING
    except OcrExhausted as exc:
        log.error("%s", exc)
        log.error("Check captcha screenshots in %s", cfg.captcha_dir)
        return EXIT_OCR_EXHAUSTED
    except FlowError as exc:
        log.error("UI step failed: %s", exc)
        return EXIT_UI_NOT_FOUND
    except KeyboardInterrupt:
        log.warning("Interrupted by user.")
        return EXIT_UNEXPECTED
    except Exception:
        log.exception("Unexpected error")
        return EXIT_UNEXPECTED

    print(f"OK: {message}")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
