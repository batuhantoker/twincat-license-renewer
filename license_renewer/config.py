"""Runtime configuration for the license renewer.

All defaults are overridable via environment variables so the script can be
dropped onto a machine with a non-standard TwinCAT or Tesseract install
without code changes.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


_DEFAULT_XAE_SHELL = r"C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe"
_DEFAULT_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

_XAE_SHELL_SEARCH_DIRS = (
    r"C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE",
    r"C:\Program Files\Beckhoff\TcXaeShell\Common7\IDE",
    r"C:\TwinCAT\3.1\Components\Base",
)


def _resolve_xae_shell() -> Path:
    override = os.environ.get("XAE_SHELL_EXE")
    if override:
        return Path(override)
    default = Path(_DEFAULT_XAE_SHELL)
    if default.exists():
        return default
    for directory in _XAE_SHELL_SEARCH_DIRS:
        base = Path(directory)
        if base.is_dir():
            for candidate in sorted(base.glob("TcXaeShell*.exe")):
                return candidate
    return default


def _appdata_dir() -> Path:
    root = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(root) / "license-renewer"


@dataclass(frozen=True)
class Config:
    xae_shell_exe: Path
    tesseract_exe: Path
    ocr_max_retries: int
    step_timeout_s: float
    captcha_dir: Path
    log_file: Path


def load() -> Config:
    app_dir = _appdata_dir()
    captcha_dir = app_dir / "captchas"
    captcha_dir.mkdir(parents=True, exist_ok=True)
    return Config(
        xae_shell_exe=_resolve_xae_shell(),
        tesseract_exe=Path(os.environ.get("TESSERACT_EXE", _DEFAULT_TESSERACT)),
        ocr_max_retries=int(os.environ.get("OCR_MAX_RETRIES", "8")),
        step_timeout_s=float(os.environ.get("STEP_TIMEOUT_S", "15")),
        captcha_dir=captcha_dir,
        log_file=app_dir / "renew.log",
    )
