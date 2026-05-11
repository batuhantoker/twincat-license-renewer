"""Screenshot and OCR the TwinCAT captcha image.

The captcha is a short alphanumeric string rendered in a fixed font inside
the license activation dialog. Pipeline:

    grab region -> grayscale -> upscale 3x -> adaptive threshold ->
    slight dilation -> Tesseract (PSM 8, alnum whitelist)
"""

from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path

import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageGrab

log = logging.getLogger(__name__)

_CAPTCHA_CHARSET = (
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz"
    "0123456789"
)
_WHITELIST = f"tessedit_char_whitelist={_CAPTCHA_CHARSET}"
_PSM_CHAIN = (7, 6, 10, 11)  # ordered by empirical accuracy on TwinCAT captchas
CAPTCHA_LEN = 5

# Letters whose lowercase / uppercase glyphs are near-identical shapes —
# Tesseract can't tell them apart and always emits uppercase. We correct
# these post-hoc using per-character bounding box heights.
# Restricted to pure x-height letters (no ascender/descender in lowercase
# form) so our height-based case detection is reliable. K/P/Y are omitted
# because their lowercase forms (k/p/y) have asc/desc that make them as
# tall as the uppercase, defeating the heuristic.
_CASE_AMBIGUOUS = set("COSUVWXZ")


def configure_tesseract(exe_path: Path) -> None:
    if exe_path.exists():
        pytesseract.pytesseract.tesseract_cmd = str(exe_path)
    else:
        log.warning(
            "Tesseract not found at %s; relying on PATH", exe_path
        )


def grab_region(bbox: tuple[int, int, int, int]) -> Image.Image:
    """Screenshot a screen-coord (left, top, right, bottom) rectangle."""
    return ImageGrab.grab(bbox=bbox, all_screens=True)


def preprocess(img: Image.Image) -> np.ndarray:
    """Isolate the red captcha text and binarize for Tesseract.

    Pipeline:
        1. Upscale 4x first (preserves anti-aliasing detail for masking).
        2. Compute "redness" = R - max(G, B). Pure red => ~255, gray/black
           text => ~0, white => 0. This kills the dark instruction text
           above the captcha and the input-field border below it.
        3. Threshold: anything with redness > 15 becomes text (black on
           white), rest becomes white.
        4. Small morphological close to fill anti-alias gaps, then open to
           drop specks.
        5. Pad the image so Tesseract has quiet borders (it needs whitespace
           around glyphs to commit).
    """
    arr = np.array(img.convert("RGB"))
    upscaled = cv2.resize(arr, None, fx=4.0, fy=4.0, interpolation=cv2.INTER_CUBIC)
    r = upscaled[:, :, 0].astype(np.int16)
    g = upscaled[:, :, 1].astype(np.int16)
    b = upscaled[:, :, 2].astype(np.int16)
    redness = r - np.maximum(g, b)
    mask = np.where(redness > 15, 0, 255).astype(np.uint8)

    close_kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, close_kernel)
    open_kernel = np.ones((2, 2), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, open_kernel)

    padded = cv2.copyMakeBorder(
        mask, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255
    )
    return padded


def _ocr_single(img_array: np.ndarray, psm: int) -> str:
    config = f"--psm {psm} -c {_WHITELIST}"
    raw = pytesseract.image_to_string(img_array, config=config)
    return "".join(raw.split())


def _correct_case(img_array: np.ndarray, text: str, psm: int) -> str:
    """Use per-character bounding box heights to lowercase 'short' glyphs.

    Tesseract cannot visually distinguish ``Z`` from ``z`` (same glyph in
    most fonts), so it always returns uppercase. We rerun OCR in
    ``image_to_boxes`` mode to get each glyph's rectangle, then lowercase
    any ambiguous letter whose height is meaningfully less than the tallest
    glyph on the line.
    """
    try:
        config = f"--psm {psm} -c {_WHITELIST}"
        boxes_str = pytesseract.image_to_boxes(img_array, config=config)
    except Exception as exc:
        log.debug("image_to_boxes failed (%s); skipping case correction.", exc)
        return text

    boxes: list[tuple[str, int]] = []
    for line in boxes_str.strip().splitlines():
        parts = line.split()
        if len(parts) >= 5:
            try:
                height = int(parts[4]) - int(parts[2])
            except ValueError:
                continue
            boxes.append((parts[0], height))

    if not boxes or len(boxes) != len(text):
        # Character count mismatch — e.g. if boxes includes digits we already
        # stripped. Bail out rather than misalign.
        return text

    heights = sorted(h for _, h in boxes)
    median_h = heights[len(heights) // 2]
    if median_h <= 0:
        return text
    # Anything meaningfully shorter than the median character is an
    # x-height (i.e. lowercase no-ascender/no-descender) glyph. Using the
    # median is robust against a single descender letter (e.g. Q, q)
    # inflating the max and against a single lowercase letter pulling the
    # average down.
    threshold = median_h * 0.80

    corrected: list[str] = []
    for (_, h), ch in zip(boxes, text):
        if ch in _CASE_AMBIGUOUS and h < threshold:
            corrected.append(ch.lower())
        else:
            corrected.append(ch)
    return "".join(corrected)


def ocr(img_array: np.ndarray) -> str:
    """Run Tesseract with a PSM fallback chain and height-based case fix.

    Tries each PSM in ``_PSM_CHAIN``; the first result whose length matches
    ``CAPTCHA_LEN`` wins. Case correction is applied using the same PSM so
    the per-character boxes align with the returned text.
    """
    last_result = ""
    last_psm = _PSM_CHAIN[0]
    for psm in _PSM_CHAIN:
        result = _ocr_single(img_array, psm)
        last_result, last_psm = result, psm
        if len(result) == CAPTCHA_LEN:
            return _correct_case(img_array, result, psm)
    return _correct_case(img_array, last_result, last_psm)


def save_debug(img: Image.Image, captcha_dir: Path, suffix: str = "") -> Path:
    captcha_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    path = captcha_dir / f"{stamp}{suffix}.png"
    img.save(path)
    return path


def read_from_image(img: Image.Image) -> str:
    return ocr(preprocess(img))


def save_preprocessed(img: Image.Image, captcha_dir: Path, stem: str) -> Path:
    """Save the post-preprocess binary image so we can see what Tesseract
    actually receives (for debugging bad OCR)."""
    captcha_dir.mkdir(parents=True, exist_ok=True)
    arr = preprocess(img)
    out = captcha_dir / f"{stem}_pp.png"
    Image.fromarray(arr).save(out)
    return out
