# Running the Qt version

## Files
- `pdf_checker_core.py` — the scanning engine (no GUI dependency; also
  usable from a CLI or test suite).
- `pdf_checker_qt.py` — the PySide6 front end. **This is the file you run.**
- `requirements.txt` — dependencies for this version.

The two files must stay in the same folder — `pdf_checker_qt.py` imports
directly from `pdf_checker_core.py`.

## Setup

```bash
pip install -r requirements.txt
python pdf_checker_qt.py
```

That's the whole setup. `PySide6` and `requests` are the only hard
requirements; everything else in `requirements.txt` is optional and the
app tells you in the UI (a banner, and the About tab) if something's
missing and what you lose without it.

## What each optional dependency gets you
- **No PyMuPDF**: the "Not Clear (Low Resolution)" check is disabled
  entirely, and "Missing Information" falls back to a slower reader
  (PyPDF2 or pdfminer.six, whichever is installed).
- **No PyPDF2 or pdfminer.six** (and no PyMuPDF either): PDF text/page
  inspection can't run at all — most files will get flagged as missing
  information, since the app has no way to read them.
- **No python-docx**: everything works except "Export to Word", which
  shows an error explaining what to install.

## First run
- On first launch it creates `pdf_checker_settings.json`,
  `pdf_checker_cache.json`, and `pdf_checker_history.json` next to
  `pdf_checker_core.py` — same behavior as the Tkinter version, just now
  centralized in the engine module rather than the UI file.
- A `pdf_checker.log` file (rotating, capped at ~1MB) is created the same
  way, for diagnosing anything that goes wrong without a console window.

## App icon (optional)
Drop `icon.ico` and/or `icon.png` next to `pdf_checker_qt.py` — the window
and system tray icon pick it up automatically at startup. No code changes
needed.

## Auto-updates
On startup, the app checks `Brandon-Morision/pdf-error-checker-pro` for a
newer tagged release. This only finds anything once you've actually
published a release on GitHub — until then it silently finds nothing,
which is expected. You can turn this check off under Settings → General.

If a release includes a matching installable asset (a `.py` file for a
from-source install like this one, or a `.exe` for a frozen/PyInstaller
build on Windows), the app offers to download and install it automatically
and then relaunches itself. Otherwise it offers to open the release page
for a manual download.

## Scan profiles
Under Settings → Scan Settings → Scan Profiles, you can save the current
DPI threshold, page-check limit, text-length threshold, empty-page ratio,
and target-subfolder settings as a named profile (e.g. one per client), and
reload any saved profile later. Profiles save immediately when you click
"Save As..." — they aren't tied to the dialog's own Save/Cancel, so
cancelling an unrelated settings edit won't lose a profile you just saved.

## The old Tkinter version
`pdf checker pro.py` (the Tkinter app, Version 1) still exists separately and is
unaffected by any of this — the two don't share code or settings files
unless you point them at the same folder. They're independent apps for now.
