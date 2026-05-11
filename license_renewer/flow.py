"""End-to-end click sequence for the TwinCAT 7-day trial license activation.

High-level steps:
    1. Solution Explorer -> <project> -> SYSTEM -> License (double-click)
    2. License editor -> Order Information tab
    3. Click "7 Days Trial License..."
    4. Solve captcha (OCR loop, OK button stays disabled until text matches)
    5. Confirm success dialog
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from typing import Iterable

from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.timings import TimeoutError as PWATimeoutError
from pywinauto.timings import Timings

from . import captcha, xae
from .config import Config

log = logging.getLogger(__name__)


def _apply_fast_timings() -> None:
    """pywinauto's UIA backend defaults are conservative. Crank them down —
    none of the TwinCAT dialogs we drive need long polling intervals."""
    Timings.fast()
    Timings.window_find_timeout = 2
    Timings.window_find_retry = 0.05
    Timings.after_click_wait = 0.02
    Timings.after_clickinput_wait = 0.02
    Timings.after_setcursorpos_wait = 0.0
    Timings.after_menu_wait = 0.02
    Timings.exists_timeout = 0.2
    Timings.exists_retry = 0.05


class FlowError(RuntimeError):
    """Raised when a UI step cannot be completed."""


class OcrExhausted(RuntimeError):
    """Raised when we have used all OCR retries without matching the captcha."""


# ---- low-level helpers -----------------------------------------------------


_PROBE_TIMEOUT_S = 0.3


def _spec_lookup(parent, spec: dict):
    """child_window for wrappers, window for Applications (top-level)."""
    if isinstance(parent, Application):
        return parent.window(**spec)
    return parent.child_window(**spec)


def _find_first(parent, candidates: Iterable[dict], timeout: float):
    """Try several spec dicts; return the first that resolves.

    Strategy: list candidates in order of specificity. We do a quick probe
    pass over every candidate (so a later spec can win if it's ready right
    away), and if nothing matched we wait patiently on the FIRST (most
    specific) spec for the remainder of the timeout.
    """
    specs = list(candidates)
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None

    for spec in specs:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        try:
            ctrl = _spec_lookup(parent, spec)
            ctrl.wait("exists visible", timeout=min(_PROBE_TIMEOUT_S, remaining))
            log.debug("Matched control with spec=%s", spec)
            return ctrl
        except (PWATimeoutError, ElementNotFoundError, ElementAmbiguousError) as exc:
            last_error = exc
            continue

    remaining = deadline - time.monotonic()
    if specs and remaining > 0.1:
        try:
            ctrl = _spec_lookup(parent, specs[0])
            ctrl.wait("exists visible", timeout=remaining)
            log.debug("Matched control (patient) with spec=%s", specs[0])
            return ctrl
        except (PWATimeoutError, ElementNotFoundError, ElementAmbiguousError) as exc:
            last_error = exc

    raise FlowError(
        f"None of {specs} matched within {timeout}s; last error: {last_error!r}"
    )


_NO_PROJECT_TITLES = {"Start Page", "TcXaeShell", ""}


def _project_name_from_title(title: str) -> str | None:
    """Extract the project name from the main window title.

    Titles look like 'PT2 - TcXaeShell' or
    'PT2 (Running) - TcXaeShell'. Returns ``None`` for the empty
    ('TcXaeShell' or 'Start Page - TcXaeShell') state.
    """
    match = re.match(r"^(?P<name>.+?)\s*(\([^)]*\))?\s*-\s*TcXaeShell", title)
    if not match:
        # No " - TcXaeShell" suffix => either no project or unrecognized.
        stripped = title.strip()
        return None if stripped in _NO_PROJECT_TITLES else None
    name = match.group("name").strip()
    if name in _NO_PROJECT_TITLES:
        return None
    return name or None


def open_most_recent_project(main_window, timeout: float = 6.0) -> None:
    """Open File -> Recent Projects and Solutions -> first entry.

    Used when we attached to an already-running XAE Shell that has no
    project loaded.
    """
    log.info("No project loaded; opening most recent from File menu.")

    file_menu = _find_first(
        main_window,
        [{"title": "File", "control_type": "MenuItem"}],
        timeout=2.0,
    )
    file_menu.click_input()
    time.sleep(0.15)

    recent = _find_first(
        main_window,
        [
            {"title": "Recent Projects and Solutions", "control_type": "MenuItem"},
            {"title_re": r".*Recent Projects.*", "control_type": "MenuItem"},
        ],
        timeout=2.0,
    )
    recent.click_input()
    time.sleep(0.2)

    # Each recent entry has an accelerator "1", "2", ... — pressing "1"
    # opens the most recent project. This is more reliable than trying to
    # locate the menu item by title (paths vary per machine).
    from pywinauto.keyboard import send_keys
    send_keys("1")

    deadline = time.monotonic() + timeout + 20
    while time.monotonic() < deadline:
        try:
            title = main_window.window_text() or ""
        except Exception:
            title = ""
        name = _project_name_from_title(title)
        if name:
            log.info("Project loaded: %s", name)
            return
        time.sleep(0.3)

    raise FlowError("Recent project did not finish loading in time.")


def _expand_tree_item(item) -> None:
    """Expand a TreeItem whether its ExpandCollapse pattern is available or
    not — some VS tree items only respond to a keyboard Right-arrow."""
    try:
        if item.is_expanded():
            return
    except Exception:
        pass
    try:
        item.expand()
        time.sleep(0.05)
        return
    except Exception as exc:
        log.debug("expand() failed (%s); falling back to keyboard.", exc)
    try:
        item.select()
        item.type_keys("{RIGHT}", set_foreground=False)
        time.sleep(0.05)
    except Exception as exc:
        log.debug("keyboard expand failed (%s).", exc)


def _tree_child_by_name(parent, name: str, max_depth: int = 2):
    """Walk a TreeItem's descendants up to ``max_depth`` levels looking for a
    TreeItem whose text equals ``name`` or starts with ``name`` + ' '/'(' .

    Uses ``.children(control_type='TreeItem')`` which is an order of
    magnitude faster than pywinauto's descendant search used by
    ``child_window``.
    """
    if max_depth < 0:
        return None
    try:
        for child in parent.children(control_type="TreeItem"):
            try:
                text = child.window_text() or ""
            except Exception:
                text = ""
            if text == name or text.startswith(name + " ") or text.startswith(name + "("):
                return child
            # Some VS trees wrap the project under "Solution 'X' (...)" -
            # recurse into the root-level solution wrapper so we can find
            # the project node underneath it.
            deeper = _tree_child_by_name(child, name, max_depth - 1)
            if deeper is not None:
                return deeper
    except Exception as exc:
        log.debug("children() walk failed on %r: %s", parent, exc)
    return None


def _find_tree_child(parent, name: str, timeout: float):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        found = _tree_child_by_name(parent, name)
        if found is not None:
            return found
        time.sleep(0.1)
    raise FlowError(
        f"Tree item {name!r} not found under {parent.window_text()!r} "
        f"within {timeout}s"
    )


# ---- individual steps ------------------------------------------------------


def open_license_manager(main_window, timeout: float):
    """Navigate Solution Explorer: <project> -> SYSTEM -> License.

    Double-clicks the License node, which opens the License editor as a
    document tab in the main window.
    """
    project = _project_name_from_title(main_window.window_text() or "")
    if not project:
        raise FlowError(
            f"Couldn't parse project name from window title "
            f"{main_window.window_text()!r}. Open a TwinCAT solution first."
        )
    log.info("Project name detected: %s", project)

    # Find the Solution Explorer tree. The name is shared by several panes
    # (toolbar, search box, main view), so target the Tree control directly.
    tree = _find_first(
        main_window,
        [
            {"title": "Solution Explorer", "control_type": "Tree"},
            {"title_re": r".*Solution Explorer.*", "control_type": "Tree"},
            {"control_type": "Tree", "found_index": 0},
        ],
        timeout=3.0,
    )

    # Fast path: walk tree children directly rather than using pywinauto's
    # (slow) descendant search. The project may sit under a
    # "Solution 'PT2' (1 of 1 project)" wrapper, which _tree_child_by_name
    # descends into automatically.
    project_item = _find_tree_child(tree, project, timeout=3.0)
    _expand_tree_item(project_item)

    system_item = _find_tree_child(project_item, "SYSTEM", timeout=3.0)
    _expand_tree_item(system_item)

    license_item = _find_tree_child(system_item, "License", timeout=3.0)
    try:
        license_item.select()
    except Exception:
        pass
    license_item.double_click_input()
    log.info("Opened License editor.")


def select_order_information_tab(license_mgr, timeout: float):
    """The License editor opens on Order Information by default, so this is
    best-effort with a short timeout. If the tab isn't found we assume
    we're already on it."""
    try:
        tab = _find_first(
            license_mgr,
            [
                {"title": "Order Information", "control_type": "TabItem"},
                {"title_re": ".*Order Information.*", "control_type": "TabItem"},
            ],
            timeout=1.0,
        )
        tab.click_input()
        log.info("Selected Order Information tab.")
    except FlowError:
        log.info("Order Information tab not found (likely already active).")


def click_trial_license_button(license_mgr, timeout: float):
    btn = _find_first(
        license_mgr,
        [
            {"title": "7 Days Trial License...", "control_type": "Button"},
            {"title_re": r"7\s*Days?\s*Trial\s*License.*", "control_type": "Button"},
            {"title_re": r".*Trial\s*License.*", "control_type": "Button"},
        ],
        timeout=5.0,
    )
    btn.click_input()
    log.info("Clicked '7 Days Trial License...' button.")


_CAPTCHA_TITLE = "Enter Security Code"


def find_captcha_dialog(app, main_window, timeout: float):
    """Locate the 'Enter Security Code' modal.

    Modal dialogs in TwinCAT can be spawned outside the XAE Shell's pywinauto
    process tree, so ``app.window(...)`` often misses them. We poll the
    full desktop across both backends instead.
    """
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None

    while time.monotonic() < deadline:
        # Win32 backend: fast and reliable for standard Windows dialogs.
        try:
            dlg = Desktop(backend="win32").window(title=_CAPTCHA_TITLE)
            if dlg.exists(timeout=0.2):
                log.debug("Found captcha dialog via Desktop(win32).")
                # Re-wrap through UIA so we keep consistent control-type
                # semantics for the captcha image / OK button lookups.
                try:
                    handle = dlg.handle
                    uia_dlg = Desktop(backend="uia").window(handle=handle)
                    if uia_dlg.exists(timeout=0.3):
                        return uia_dlg
                except Exception as exc:
                    log.debug("UIA re-wrap failed (%s); using win32 wrapper.", exc)
                return dlg
        except Exception as exc:
            last_error = exc

        # UIA backend at desktop level.
        try:
            dlg = Desktop(backend="uia").window(title=_CAPTCHA_TITLE)
            if dlg.exists(timeout=0.2):
                log.debug("Found captcha dialog via Desktop(uia).")
                return dlg
        except Exception as exc:
            last_error = exc

        time.sleep(0.1)

    raise FlowError(
        f"Could not find '{_CAPTCHA_TITLE}' dialog within {timeout}s "
        f"(last error: {last_error!r})"
    )


def grab_captcha_bbox(dialog) -> tuple[int, int, int, int]:
    """Return screen-coord bbox of the captcha image inside the 'Enter
    Security Code' dialog.

    The dialog layout is stable: red 5-char text sits under the instruction
    label in the upper-left half, above the input field. We screenshot a
    generous region — the red-channel mask in ``captcha.preprocess`` handles
    cropping away non-red pixels.
    """
    rect = dialog.rectangle()
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    # We crop generously in the left/middle of the dialog: the red-channel
    # filter in ``captcha.preprocess`` wipes out everything that isn't the
    # red captcha glyphs (instruction label, input field border, etc.), so
    # a wider region is safer than a tight one.
    left = rect.left + int(width * 0.05)
    right = rect.left + int(width * 0.65)
    top = rect.top + int(height * 0.25)
    bottom = rect.top + int(height * 0.70)
    log.debug(
        "Dialog rect=%s; captcha region=(%d,%d,%d,%d)",
        rect, left, top, right, bottom,
    )
    return (left, top, right, bottom)


def find_ok_button(dialog, timeout: float):
    return _find_first(
        dialog,
        [
            {"title": "OK", "control_type": "Button"},
            {"title": "OK", "control_type": "Button", "found_index": 0},
            {"title_re": r"^OK$"},
        ],
        timeout,
    )


def _refresh_captcha(dialog, license_mgr, timeout: float):
    """Cancel the current captcha dialog and reopen it to get a fresh
    (different) captcha image. Returns the new dialog wrapper."""
    from pywinauto.keyboard import send_keys
    log.debug("Refreshing captcha (cancel + reopen 7-Day Trial button).")
    # Cancel via Escape — works whether the Cancel button is visible or not.
    try:
        dialog.set_focus()
    except Exception:
        pass
    send_keys("{ESC}")
    # Wait a moment for the dialog to close before reopening.
    time.sleep(0.25)
    click_trial_license_button(license_mgr, timeout)
    return find_captcha_dialog(None, license_mgr, timeout)


def solve_captcha(dialog, license_mgr, cfg: Config):
    """Screenshot captcha, OCR, type the guess, click OK.

    If the OK button stays disabled (wrong captcha), cancel the dialog and
    reopen it to get a fresh image — Tesseract is deterministic, so retrying
    OCR on the *same* image would produce the same wrong guess forever.

    The Edit field in the 'Enter Security Code' modal has focus when the
    dialog opens, so we type directly via keyboard instead of locating the
    Edit control (there are two Edits, making a control-based lookup
    ambiguous).

    Returns the (possibly refreshed) dialog wrapper on success.
    """
    from pywinauto.keyboard import send_keys

    last_saved: Path | None = None
    for attempt in range(1, cfg.ocr_max_retries + 1):
        bbox = grab_captcha_bbox(dialog)
        img = captcha.grab_region(bbox)
        guess = captcha.read_from_image(img)
        last_saved = captcha.save_debug(
            img, cfg.captcha_dir, suffix=f"_try{attempt}_{guess or 'empty'}"
        )
        captcha.save_preprocessed(img, cfg.captcha_dir, last_saved.stem)
        log.info("Attempt %d: OCR guess = %r (saved %s)", attempt, guess, last_saved)

        if len(guess) == captcha.CAPTCHA_LEN:
            try:
                dialog.set_focus()
            except Exception as exc:
                log.debug("dialog.set_focus() failed (%s).", exc)
            send_keys("^a")
            send_keys(guess, pause=0.01, with_spaces=False)

            # Re-find OK button each loop because the dialog handle may have
            # changed after a refresh.
            try:
                ok_btn = find_ok_button(dialog, timeout=2.0)
            except FlowError:
                ok_btn = None

            if ok_btn and _wait_for_enabled(ok_btn, timeout_s=1.5):
                try:
                    ok_btn.click_input()
                except Exception:
                    send_keys("{ENTER}")
                log.info("Captcha accepted on attempt %d: %r", attempt, guess)
                return dialog

            log.info("OK still disabled after %r; refreshing captcha.", guess)
        else:
            log.info(
                "Rejecting guess %r (wrong length, expected %d); refreshing captcha.",
                guess, captcha.CAPTCHA_LEN,
            )

        # Failed this captcha — cancel and reopen to get a new image.
        if attempt < cfg.ocr_max_retries:
            try:
                dialog = _refresh_captcha(dialog, license_mgr, cfg.step_timeout_s)
            except Exception as exc:
                log.warning("Captcha refresh failed (%s); retrying on same image.", exc)
                time.sleep(0.3)

    raise OcrExhausted(
        f"Captcha unsolved after {cfg.ocr_max_retries} attempts "
        f"(last screenshot: {last_saved})"
    )


def _wait_for_enabled(ctrl, timeout_s: float, poll_s: float = 0.2) -> bool:
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        try:
            if ctrl.is_enabled():
                return True
        except Exception:
            pass
        time.sleep(poll_s)
    return False


_SUCCESS_TITLE_RE = re.compile(
    r"(Success|Activated|License|Information|TcXaeShell|Microsoft Visual Studio)",
    re.IGNORECASE,
)


def _enumerate_top_windows() -> list[tuple[str, str]]:
    """Return [(title, class_name), ...] for currently visible top-level
    windows, for debugging what success dialog actually shows up as."""
    out: list[tuple[str, str]] = []
    try:
        for w in Desktop(backend="win32").windows():
            try:
                title = w.window_text() or ""
                cls = w.class_name() or ""
                if title.strip() and w.is_visible():
                    out.append((title, cls))
            except Exception:
                continue
    except Exception:
        pass
    return out


_STD_DIALOG_CLASS = "#32770"  # Windows' classic dialog class.


def wait_for_success(app, main_window, timeout: float):
    """Find and dismiss the success confirmation modal.

    The modal is a standard Windows dialog (class ``#32770``), same as the
    captcha. We watch for any top-level dialog of that class that isn't the
    main XAE Shell window or a lingering captcha dialog, then click its OK
    button (or send Enter as a fallback).
    """
    from pywinauto.keyboard import send_keys

    try:
        main_title = main_window.window_text() or ""
    except Exception:
        main_title = ""

    deadline = time.monotonic() + timeout
    dialog_handle: int | None = None
    dialog_title = ""

    while time.monotonic() < deadline and dialog_handle is None:
        try:
            for w in Desktop(backend="win32").windows():
                try:
                    cls = w.class_name() or ""
                    title = w.window_text() or ""
                except Exception:
                    continue
                if cls != _STD_DIALOG_CLASS:
                    continue
                if title == _CAPTCHA_TITLE or title == main_title or not title.strip():
                    continue
                dialog_handle = w.handle
                dialog_title = title
                log.info("Found success dialog: %r (handle=%s)", title, dialog_handle)
                break
        except Exception as exc:
            log.debug("Desktop(win32).windows() failed: %s", exc)
        if dialog_handle is None:
            time.sleep(0.1)

    if dialog_handle is None:
        log.warning(
            "Success dialog not found; visible windows were: %s",
            _enumerate_top_windows(),
        )
        send_keys("{ENTER}")
        return "(no success dialog detected; Enter sent as fallback)"

    # Re-wrap via UIA using the handle so we get child_window / control_type
    # support for finding the OK button.
    try:
        dlg = Desktop(backend="uia").window(handle=dialog_handle)
        ok = dlg.child_window(title="OK", control_type="Button")
        ok.wait("enabled visible", timeout=1.5)
        ok.click_input()
        log.info("Clicked OK on success dialog.")
    except Exception as exc:
        log.debug("OK-click path failed (%s); falling back to Enter.", exc)
        try:
            dlg = Desktop(backend="win32").window(handle=dialog_handle)
            dlg.set_focus()
        except Exception:
            pass
        send_keys("{ENTER}")

    return dialog_title


# ---- top-level orchestration ----------------------------------------------


def _timed(label: str):
    """Context-manager-ish timer via decorator pattern."""
    class _T:
        def __enter__(self_inner):
            self_inner.t0 = time.monotonic()
            log.debug("[step] %s ...", label)
            return self_inner
        def __exit__(self_inner, *exc):
            dt = time.monotonic() - self_inner.t0
            log.info("[step] %s done in %.2fs", label, dt)
    return _T()


def run(cfg: Config, dry_run: bool = False, project_path: str | None = None) -> str:
    _apply_fast_timings()
    captcha.configure_tesseract(cfg.tesseract_exe)

    with _timed("attach XAE Shell"):
        app, main_window = xae.connect_or_launch(
            cfg.xae_shell_exe, cfg.step_timeout_s, project_path=project_path,
        )

    if dry_run:
        log.info("Dry run: attached to XAE Shell, skipping click sequence.")
        return "dry-run-ok"

    # If the window title shows no project (e.g. we attached to an empty
    # XAE Shell), open the most recent solution from the File menu.
    current_project = _project_name_from_title(main_window.window_text() or "")
    if not current_project:
        with _timed("open most recent project"):
            open_most_recent_project(main_window, timeout=cfg.step_timeout_s)

    with _timed("open License editor"):
        open_license_manager(main_window, cfg.step_timeout_s)

    license_mgr = main_window
    with _timed("select Order Information tab"):
        select_order_information_tab(license_mgr, cfg.step_timeout_s)
    with _timed("click 7-Day Trial button"):
        click_trial_license_button(license_mgr, cfg.step_timeout_s)

    with _timed("find captcha dialog"):
        dialog = find_captcha_dialog(app, main_window, cfg.step_timeout_s)
    with _timed("solve captcha"):
        solve_captcha(dialog, license_mgr, cfg)

    with _timed("wait for success"):
        message = wait_for_success(app, main_window, cfg.step_timeout_s)
    log.info("License activation successful: %s", message or "(no message captured)")
    return message or "activated"
