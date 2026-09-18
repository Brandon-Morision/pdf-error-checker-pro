# PDF Error Checker Pro

A professional desktop application for scanning and validating PDF files across multiple project folders.

**This is Version 2 — the Qt Edition.** The app has been rewritten on PySide6/Qt in place of the original Tkinter UI. The scanning logic itself (`pdf_checker_core.py`) is unchanged in behavior and carries forward every fix and feature from Version 1; only the interface layer and its threading model changed.

## What's new in Version 2 (Qt Edition)

- 🎨 **Modern, flat, rounded UI** — a clean, cohesive stylesheet replaces the old hand-rolled Tkinter custom-drawn canvas controls with native Qt widgets, rich tooltips across every setting, and adaptive empty-state messages for initial, zero-error, and filtered states.
- ⚡ **Virtualized results table with live duplicate updates** — the results list is a high-performance `QTableView` backed by a virtualized data model with instant multi-column sorting and filtering. When duplicate detection flags a file that was already in the results list, the table dynamically updates on screen via model signals.
- 🧵 **Resilient threading architecture & crash guards** — scans run on a background `QThread` using Qt's native signal/slot mechanism. Worker exceptions are caught and safely surfaced so the UI never locks or freezes during unexpected scanning errors, and thread termination is bounded and explicit on window close.
- 🛡️ **Accurate corruption detection & no redundant disk I/O** — `is_pdf_corrupt()` now cleanly resolves on PyMuPDF inspection rather than falling through to PyPDF2, preventing redundant disk reads and stopping valid modern PDFs from being falsely flagged as "Cannot Open".
- 🔒 **Guaranteed PyMuPDF handle disposal** — all PDF inspection routines now use guarded `finally` blocks to guarantee file handles close immediately, preventing Windows file-locking conflicts (`WinError 32`).
- 📐 **Persistent window geometry & report preferences** — window width and height are remembered across launches, and Word report export offers a "Don't ask again" prompt that persists your auto-open preference.
- ⚡ **Thread-safe & case-normalized scan cache** — `ScanCache` operations are protected by thread locks, and cache pruning handles Windows case-insensitivity cleanly without accidental cache purges.
- 🔔 **Native desktop notifications** — via Qt's built-in `QSystemTrayIcon`, removing the third-party `plyer` dependency entirely.
- 🐛 **Fixed persistent data directory in built `.exe`** — settings, cache, history, and logs are anchored next to the running executable rather than PyInstaller's temporary per-launch folder, ensuring data is never lost between sessions.
- 📦 **Hardened standalone executable builds** — `build_exe.bat` packages dynamic hidden imports for PyMuPDF, pdfminer, and python-docx, and dependency specifications are unified into a single `requirements.txt`.
- 🧾 **Modular dual-file architecture** — the scanning engine (`pdf_checker_core.py`) is completely independent of the GUI (`pdf_checker_qt.py`), ready for automated testing, CLI, or future web interfaces.

Two files ship together now instead of one: `pdf_checker_qt.py` (the application you run) and `pdf_checker_core.py` (the scanning engine it imports — no GUI dependency at all, so it could in principle back a different UI or a CLI later). They must stay in the same folder.

## Feature history (carried forward from Version 1)

Everything below describes scanning behavior and capabilities that are identical in Version 2 — only the UI framework changed, not what the app actually checks or how it checks it.

- 🛡️ **Scope-aware scan caching** — cache pruning is strictly scoped to the folder(s) actively scanned in the current session, so scanning one folder doesn't purge cached entries for other folders checked in earlier sessions.
- 🧬 **Timeout-protected parallel duplicate hashing** — duplicate detection hashes unhashed files concurrently using a worker pool protected by the configured per-file timeout, so one oversized or slow-to-read PDF can't stall the duplicate check.
- 🔒 **UI state protection during scans** — Browse, Settings, and History are disabled during an active scan alongside Scan and Export, preventing conflicting threshold changes or in-memory cache resets while background threads are running.
- ⏱️ **Auto-update installer safety timeout** — the Windows executable self-replacement script caps its wait at 30 seconds; if an external lock (e.g. antivirus) prevents replacement, it fails with a clear message and leaves the downloaded update in place instead of looping indefinitely.
- 🔍 **Strict structural inspection for "Missing Information"** — an unexpected exception during content inspection flags the file for review instead of silently passing it as clean.
- ⚡ **Parallel scanning** — files are checked concurrently with a configurable worker pool (Settings → Scan Settings → Performance).
- ⏱️ **Per-file timeout guard** — a single huge or malformed PDF can't stall the whole scan; if one file exceeds the configured timeout, it's flagged `Scan Timeout` and the scan moves on.
- 💾 **Skip-unchanged-files caching** — re-scanning the same tree after fixing a handful of files reuses cached results for anything whose size and modified-time haven't changed.
- 🔑 **Password-protected PDF detection** — encrypted PDFs are flagged as their own "Password Protected" error, separate from "Cannot Open".
- 🧬 **Duplicate detection** — file contents are hashed to flag the same PDF appearing in more than one scanned location.
- 🗒️ **Scan history log** — the History button records date, folder, mode, files checked, errors found, and duration for past runs.
- 🧾 **Scan profiles** — save and reload named sets of thresholds and target subfolders per client or project type, under Settings → Scan Settings.
- 📖 **Auto-open Word report** — after exporting, optionally open the report immediately, with a "Don't ask again" choice that sticks.
- 🕳️ **Empty-state polish** — the results panel shows a friendly message before any scan, after a clean scan, or when a filter excludes everything.
- 🪵 **File-based logging** — failures are logged to `pdf_checker.log` (with rotation) instead of only a console window nobody's watching.
- 🚀 **In-app auto-updater** — download and install updates directly from GitHub releases with progress tracking and automatic restart.
- 🎯 **Accurate content detection for "Missing Information"** — pages with embedded raster images, vector drawings, or interactive form widgets aren't falsely treated as empty.
- 🔬 **Smart DPI calculation for "Not Clear"** — resolution is computed against each image's actual rendered bounding box, filtering out thin decorative rules/separators and protecting standard-resolution graphics with a tolerance margin.
- 📁 **Flexible folder matching** — subfolder names aren't hardcoded to `open`/`confidential`; matching is case-insensitive, only one target subfolder is required (not all), and common naming variations are recognized automatically.
- 🗂️ **Scan Mode toggle** — **Project Structure** (folder-matching above) or **All PDFs (Recursive)**, which scans every PDF under the selected folder regardless of naming.
- 🔀 **Sortable, filterable results table** — click any column header to sort; type in the filter box to narrow results live.
- 🖱️ **Open results directly** — double-click a row to open the PDF, or right-click for Open File / Show in Folder / Copy Path.
- ⌨️ **Keyboard shortcuts** — `Enter` starts a scan, `Esc` cancels, `Ctrl+E` exports to Word.
- ⏱️ **Live elapsed time + ETA** during a scan.
- ⏹️ **Faster cancel** — cancel takes effect within the file currently being checked, not only between files.
- 🎨 **App icon support** — drop an `icon.ico`/`icon.png` next to the script (or bundle via PyInstaller) and it's picked up automatically for the window and About dialog.

## Features

- 🔍 **Multi-folder scanning** — recursively finds project folders containing your configured target subfolders (default: 'open' and 'confidential'), matched case-insensitively and tolerant of naming variations
- 🗂️ **Two scan modes** — strict Project Structure matching, or All PDFs (Recursive) to scan everything under a folder regardless of naming
- ⚡ **Parallel, cached processing** — a configurable worker pool checks files concurrently, with per-file timeouts and a scope-aware, skip-unchanged-files cache for fast re-scans across multiple sessions
- 🧬 **Password & duplicate detection** — encrypted PDFs and files that appear more than once (by content hash, with parallel timeout-guarded processing) are flagged as their own error types
- 🔀 **Sortable, filterable, virtualized results table** — click-to-sort columns and a live filter box that stay responsive even with large result sets
- 🖱️ **One-click file access** — double-click or right-click any result to open it or reveal it on disk
- 📊 **Detailed reports** — export a professional Word document with error summaries, with optional auto-open
- 🗒️ **Scan history** — a local log of past runs (date, folder, mode, files, errors, duration)
- 🧾 **Scan profiles** — save/load named threshold sets per client or project type
- 🔒 **UI state protection** — controls are automatically locked during scans to safeguard settings and cached state
- 🔔 **Native desktop notifications** when a scan finishes
- ⚙️ **Configurable** — settings dialog with persistent configuration (DPI threshold, text-length threshold, empty-page ratio, auto-update toggle, parallel workers, per-file timeout, scan cache)
- 🎨 **Modern UI** — a real Qt widget theme with rounded controls, tooltips, and friendly empty states throughout
- 🔄 **In-app auto-updates** — automatic startup check against GitHub releases with background downloading, progress display, and self-restarting installation with timeout guards

## Installation

### Option 1: Run from Source

1. Install Python 3.9 or higher (PySide6 requires 3.9+)
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python pdf_checker_qt.py
   ```

`pdf_checker_qt.py` and `pdf_checker_core.py` must be in the same folder — the first imports directly from the second.

### Option 2: Build EXE (Windows)

1. Install Python 3.9 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the build script:
   ```bash
   build_exe.bat
   ```
4. Find your EXE in the `dist` folder

`build_exe.bat` invokes PyInstaller directly (no `.spec` file needed) and bundles `pdf_checker_core.py` automatically via PyInstaller's normal import analysis — the same way it already handled every module in the single-file Version 1 build.

To give the built EXE a custom icon, place `icon.ico` (and optionally `icon.png`) next to `pdf_checker_qt.py` before running `build_exe.bat` — it's picked up for both the `.exe`'s own icon and the in-app window/About-dialog icon automatically.

**Persistent files and the built EXE:** settings, scan cache, scan history, and the log file are created next to the running `.exe` itself the first time it launches — not inside PyInstaller's `dist`/`build` folders, and (as of this version) not lost between runs. If you move the `.exe`, its data moves with whatever folder you place it in.

## Usage

1. Click **Browse** to select a parent folder
2. Choose a **Scan Mode**:
   - **Project Structure** — the app finds all subfolders containing at least one of your configured target names (default: 'open', 'confidential'), matched case-insensitively and tolerant of common naming variations
   - **All PDFs (Recursive)** — every PDF under the selected folder is scanned, regardless of subfolder naming
3. Configure scan options (which error types to check, resolution threshold)
4. Click **Start Scan** (or press `Enter`)
5. Use the filter box or click a column header to sort/narrow results as they come in
6. Double-click a result to open the file, or right-click it to reveal it in your file manager
7. Export to Word if needed (`Ctrl+E`)

## Keyboard Shortcuts

| Key      | Action           |
|----------|------------------|
| `Enter`  | Start Scan       |
| `Esc`    | Cancel Scan      |
| `Ctrl+E` | Export to Word   |

Shortcuts are disabled while typing in a text field (e.g. the results filter box) so they don't interrupt normal typing.

## Settings

Access settings via the **Settings** button in the header:

- **General**: the auto-update-check toggle, and scan-cache controls (enable/disable, Clear Cache)
- **Scan Settings**: scan profiles (save/load/delete named threshold sets), target subfolder names and whether to match common naming variations, DPI threshold, pages-to-check for the resolution check, minimum text length, empty-page ratio threshold, parallel worker count, and per-file timeout
- **About**: version info and live dependency status

## Notifications

When a scan finishes, the app sends a native desktop notification via Qt's `QSystemTrayIcon` — no extra dependency needed. If a system tray isn't available on your platform/session, it falls back to a standard pop-up dialog.

## Auto-Updates

On startup, the app checks this repository's latest GitHub release. If a newer version is found, it opens an **Update Available** dialog with the release notes and one of two actions:

- If the release has a compatible downloadable asset (a `.exe` when running as a built app, or a `.py` when running from source), you'll see **Download & Install**: the app downloads it in the background with a progress dialog, then — after you confirm — closes and restarts itself with the new version in place. On Windows, this hands off to a small helper script that waits for the app to fully exit, replaces the executable, and relaunches it (capped at 30 seconds to prevent an indefinite background hang if locked by antivirus). Running from source, the script file is replaced directly and the app relaunches; a `.bak` copy of the previous version is kept alongside it.
- If no compatible asset is found, you'll instead see **Open Release Page**, which opens the release in your browser for a manual download.

You can disable the startup check entirely under Settings → General.

**Note:** Ensure `GITHUB_REPO` (near the top of `pdf_checker_core.py`) points at your GitHub repository (currently `"Brandon-Morision/pdf-error-checker-pro"`) and publish tagged releases there with the appropriate asset(s) attached. If you fork or rename the repository, update `GITHUB_REPO` accordingly.

## Requirements

- Python 3.9+
- `PySide6` — **required**; the app's UI cannot start without it
- `requests` — **required**; the auto-updater and version check cannot run without it
- `PyMuPDF` (fitz) — recommended; without it, the resolution check is disabled and the missing-information check falls back to a slower reader
- `PyPDF2` — recommended; used as a fallback for the corrupt-file check and the missing-information check
- `pdfminer.six` — last-resort text extraction fallback if both PyMuPDF and PyPDF2 are unavailable
- `python-docx` — required only to use "Export to Word"

See `requirements.txt` for exact version pins and the same core/recommended/build-only grouping.

**No longer required as of Version 2:** `Pillow` (Qt's `QIcon` loads images natively) and `plyer` (replaced by Qt's built-in `QSystemTrayIcon`).

## License

MIT License
