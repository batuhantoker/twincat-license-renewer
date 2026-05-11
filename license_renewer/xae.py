"""Attach to a running TwinCAT XAE Shell, or launch one if needed."""

from __future__ import annotations

import logging
from pathlib import Path

from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.application import ProcessNotFoundError

log = logging.getLogger(__name__)


class XaeShellNotFound(RuntimeError):
    """The XAE Shell executable is missing or could not be started."""


def connect_or_launch(exe_path: Path, timeout_s: float, project_path: str | None = None):
    """Return ``(Application, top_window)`` for the XAE Shell.

    Attaches to a running instance if one is found; otherwise starts a new
    instance (optionally loading ``project_path`` as a solution/project).
    Waits for the main window to be ready before returning.
    """
    if not exe_path.exists():
        raise XaeShellNotFound(f"XAE Shell not found at {exe_path}")

    app = Application(backend="uia")
    try:
        app.connect(path=str(exe_path), timeout=2)
        log.info("Attached to running XAE Shell (pid=%s).", app.process)
        if project_path:
            log.warning(
                "XAE Shell is already running; --project %s will not be "
                "loaded on an attached instance.", project_path,
            )
    except (ProcessNotFoundError, ElementNotFoundError):
        cmd = f'"{exe_path}"'
        if project_path:
            cmd += f' "{project_path}"'
        log.info("Launching XAE Shell: %s", cmd)
        app.start(cmd)

    window = app.top_window()
    window.wait("visible ready", timeout=timeout_s * 2)
    log.debug("XAE Shell main window: %r", window.window_text())
    return app, window
