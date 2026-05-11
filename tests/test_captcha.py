"""Captcha OCR regression tests.

Drop real captcha screenshots into ``tests/fixtures/captchas/`` with the
ground-truth text as the filename stem, e.g. ``AB3D9.png``. The test loads
each fixture, runs the preprocessing + OCR pipeline, and asserts the result.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from license_renewer import captcha

FIXTURE_DIR = Path(__file__).parent / "fixtures" / "captchas"


def _fixtures() -> list[Path]:
    if not FIXTURE_DIR.is_dir():
        return []
    return sorted(p for p in FIXTURE_DIR.iterdir() if p.suffix.lower() == ".png")


@pytest.mark.skipif(not _fixtures(), reason="no captcha fixtures yet")
@pytest.mark.parametrize("path", _fixtures(), ids=lambda p: p.stem)
def test_ocr_matches_ground_truth(path: Path) -> None:
    expected = path.stem
    img = Image.open(path)
    actual = captcha.read_from_image(img)
    assert actual == expected, f"expected {expected!r}, got {actual!r}"
