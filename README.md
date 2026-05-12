# TwinCAT License Renewer

One-shot Python tool that activates the TwinCAT 3 XAE Shell 7-day trial
license — including OCR'ing the captcha — so you don't have to click
through the dialog every week.

## What it does

1. Attaches to a running **TwinCAT XAE Shell (standalone)** or launches one.
2. If no project is loaded, opens one (via `--project` or
   **File → Recent Projects and Solutions → 1**).
3. Navigates the Solution Explorer tree to `<project> → SYSTEM → License`
   and double-clicks it to open the License editor.
4. Clicks **"7 Days Trial License..."** on the Order Information tab.
5. Screenshots the captcha image from the modal dialog, OCRs it with
   Tesseract, types the guess, and polls the **OK** button — a wrong
   captcha leaves it disabled (TwinCAT doesn't throw an error), so on
   failure the script **cancels the dialog and reopens it to get a fresh
   captcha image** before retrying.
6. Finds and dismisses the success confirmation dialog.

## Prerequisites

- Windows 10/11
- Python 3.11+
- **TwinCAT XAE Shell** installed at
  `C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe`
  (override via `XAE_SHELL_EXE` env var). The script also scans
  `C:\Program Files\Beckhoff\TcXaeShell\Common7\IDE\` and
  `C:\TwinCAT\3.1\Components\Base\` as fallbacks.
- **Tesseract OCR for Windows** — the [UB Mannheim build](https://github.com/UB-Mannheim/tesseract/wiki)
  is recommended. Default path: `C:\Program Files\Tesseract-OCR\tesseract.exe`
  (override via `TESSERACT_EXE`).

### 32-bit vs 64-bit XAE Shell

Beckhoff ships TcXaeShell in two bitnesses, and they can be installed side by side:

| Bitness | Install path |
|---------|--------------|
| 32-bit  | `C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe` |
| 64-bit  | `C:\Program Files\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe` |

Two things to know:

- **`XAE_SHELL_EXE` must match the shell you actually use.** The script
  attaches to a *running* shell by exact executable path, so the default
  (the 32-bit path above) will only ever attach to the 32-bit instance —
  even if a 64-bit one is also running. If you work in the 64-bit shell,
  set `XAE_SHELL_EXE` to the `Program Files\` path (env var or `.env`).
- **Run Python with the same bitness as the shell.** TcXaeShell's Solution
  Explorer is a WPF control; a *64-bit* Python + pywinauto often can't
  enumerate the WPF tree of the *32-bit* shell (and vice versa), which
  shows up as `Exit 3 — Could not find 'Solution Explorer' tree`. Matching
  bitness avoids it. Easiest path if you're unsure: use the **32-bit**
  TcXaeShell with a **32-bit** Python venv.

Check what's running and which Python you have:

```powershell
Get-CimInstance Win32_Process -Filter "Name='TcXaeShell.exe'" | Select-Object ProcessId, ExecutablePath
python -c "import struct; print(struct.calcsize('P') * 8)"   # 32 or 64
```

Install Tesseract via winget:

```bash
winget install --exact --id UB-Mannheim.TesseractOCR
```

## Install

```bash
pip install -e .
```

## Usage

```bash
python renew.py                              # use open project, or auto-open most recent
python renew.py --project C:\path\to\PT2.sln # launch XAE Shell with this solution
python renew.py --verbose                    # DEBUG logs + step timings in console
python renew.py --dry-run                    # attach to XAE Shell, verify path, do nothing
```

**Project handling:**

- If XAE Shell is **not running**, the script launches it. When `--project`
  is supplied, that solution is loaded at launch time via the command line.
- If XAE Shell **is already running** with a project loaded, the script
  uses that project. `--project` is ignored in this case (you can't load
  a project into an attached instance from the outside — close XAE Shell
  first if you need a specific one).
- If XAE Shell **is already running** with no project loaded (title is
  `TcXaeShell` or `Start Page - TcXaeShell`), the script opens the first
  entry under **File → Recent Projects and Solutions** via UIA menu
  navigation + pressing `1`.

Exit codes:

| Code | Meaning |
|------|---------|
| 0 | Success |
| 2 | XAE Shell executable not found |
| 3 | A UI element couldn't be located (TwinCAT version may differ) |
| 4 | OCR exhausted retries — captcha couldn't be read |
| 5 | Unexpected error |

## Configuration (env vars)

| Variable | Default | Notes |
|----------|---------|-------|
| `XAE_SHELL_EXE` | `C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe` | The **32-bit** shell. Set to the `Program Files\` path for the 64-bit shell — see [32-bit vs 64-bit XAE Shell](#32-bit-vs-64-bit-xae-shell). Falls back to scanning the Beckhoff and legacy TwinCAT install folders |
| `TESSERACT_EXE` | `C:\Program Files\Tesseract-OCR\tesseract.exe` | |
| `OCR_MAX_RETRIES` | `8` | Max number of fresh captchas to try (each refresh = cancel + reopen) |
| `STEP_TIMEOUT_S` | `15` | Per-step wait ceiling for a UI control to appear |

Copy `.env.example` to `.env` and edit in place if you'd rather not set
shell env vars. The script loads `.env` from the current working directory
at startup; real shell env vars always take precedence over values in
`.env`.

## How the OCR works

The captcha is 5 mixed-case alphanumeric characters in red on white.
Pipeline in `license_renewer/captcha.py`:

1. **Crop** a region in the upper-left of the captcha dialog (relative to
   its `.rectangle()`, which pywinauto reports in DPI-aware screen coords
   — the script enables per-monitor DPI awareness at startup so
   `pywinauto` and `PIL.ImageGrab` agree on coordinates).
2. **Upscale 4x** (INTER_CUBIC) to preserve anti-aliasing detail.
3. **Red-channel isolation**: `redness = R - max(G, B)`. Anything with
   redness > 15 becomes text (black); everything else (the instruction
   label above the captcha, the input field border, etc.) becomes white.
4. Small **morphological close + open** to fill anti-alias gaps and drop
   specks; white border padding so Tesseract has quiet margins.
5. **Tesseract PSM fallback chain**: tries PSM 7 (single text line),
   then 6, 10, 11. The first result matching the expected 5-char length
   wins.
6. **Height-based case correction**: Tesseract can't visually distinguish
   `z` from `Z`, `c` from `C`, `o` from `O`, etc. (same glyph shape). The
   script calls `pytesseract.image_to_boxes()`, measures each glyph's
   bounding box height, and lowercases letters from the pure-x-height set
   (`COSUVWXZ`) whose height is under 80% of the median letter height.

## Fresh-captcha refresh on failure

TwinCAT's captcha dialog doesn't auto-regenerate when the input is wrong —
the OK button just stays disabled. That means re-OCRing the same image
gives the same wrong answer forever. On each OCR miss the script:

1. Sends Escape to close the current captcha dialog.
2. Re-clicks **7 Days Trial License...** to reopen it with a different
   captcha image.
3. Runs OCR on the new image and tries again.

`OCR_MAX_RETRIES` caps the number of fresh captchas attempted (default 8).
In practice most captchas are readable first try; the occasional
hard-to-read glyph combo (e.g. `j` vs `y`) gets skipped by refreshing.

## Where it writes things

`%LOCALAPPDATA%\license-renewer\`

- `renew.log` — rolling log file (DEBUG level).
- `captchas\<ts>_tryN_<guess>.png` — raw screenshot of each captcha attempt.
- `captchas\<ts>_tryN_<guess>_pp.png` — post-preprocess binary image (what
  Tesseract actually sees). Useful for diagnosing bad OCR.

## Improving OCR accuracy

If you want tighter OCR on a specific build/font:

1. Pick a handful of real captchas from `%LOCALAPPDATA%\license-renewer\captchas\`
   (use the `_pp.png` versions), rename them to their ground-truth text
   (e.g. `CUzHQ.png`), and move them to `tests\fixtures\captchas\`.
2. Run `pytest tests\test_captcha.py` — failing cases show exactly where
   the pipeline breaks and point to tweaks in `license_renewer\captcha.py`
   (threshold, morphology kernel sizes, PSM chain, case-ambiguous set).

## Troubleshooting

- **Exit 3, "Could not find ... within Ns"**: the UI layout depends on the
  XAE Shell version / installed extensions. Run with `--verbose` — the
  `[step]` timings pinpoint which step failed. Common fixes:
  - Project tree lookups: confirm the project name is exactly what's shown
    in the window title (before `- TcXaeShell`). Case sensitive.
  - "Recent Projects and Solutions": if the menu item has different
    wording in your VS shell version, adjust `open_most_recent_project`
    in `license_renewer/flow.py`.
- **Exit 3, License editor / "Solution Explorer" tree not found**: this
  step walks the Solution Explorer tree. If the Solution Explorer pane is
  closed, open it (Ctrl+Alt+L) before running. If the tree is still empty,
  make sure the project has loaded fully (look for `SYSTEM` under the
  project node). If the tree control itself can't be found at all (not just
  a missing node), it's almost always a **bitness mismatch** between Python
  and TcXaeShell — see [32-bit vs 64-bit XAE Shell](#32-bit-vs-64-bit-xae-shell).
- **Exit 4, captcha unsolved**: open the latest `_pp.png` in
  `%LOCALAPPDATA%\license-renewer\captchas\`.
  - All white / blank → crop region is off; adjust the percentages in
    `grab_captcha_bbox` (`license_renewer/flow.py`).
  - Text visible but OCR result is wrong letters → adjust `preprocess()`
    in `captcha.py` (threshold value, kernel sizes, or add another PSM
    to `_PSM_CHAIN`).
  - Specific letter pair swapped (e.g. `j`/`y`, `I`/`1`) → these are
    font-level Tesseract limitations. The refresh-on-failure loop skips
    these by generating a new captcha.
- **Slow runtime**: the script applies `pywinauto.Timings.fast()` and
  hand-tunes several intervals. If steps still take >5s in `--verbose`
  output, check that XAE Shell isn't busy (building, debugging, etc.) —
  UIA queries block behind VS's main thread.

## Project layout

```
license-renewer/
├── pyproject.toml
├── README.md
├── renew.py                    # thin CLI wrapper
├── license_renewer/
│   ├── __init__.py
│   ├── __main__.py             # argparse + DPI awareness + exit codes
│   ├── config.py               # env-var-overridable paths & timeouts
│   ├── logging_setup.py        # stdout + rolling file
│   ├── xae.py                  # attach/launch TcXaeShell
│   ├── flow.py                 # click sequence + captcha refresh loop
│   └── captcha.py              # preprocess + OCR + case correction
└── tests/
    ├── test_captcha.py         # OCR regression against fixture PNGs
    └── fixtures/captchas/      # drop ground-truth-named PNGs here
```
