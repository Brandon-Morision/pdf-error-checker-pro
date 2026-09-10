# PDF Error Checker Pro

A professional desktop application for scanning and validating PDF files across multiple project folders.

## What's new in this update

- ⚡ **Parallel scanning** — files are now checked concurrently with a configurable worker pool (Settings → Scan Settings → Performance), instead of one at a time, so large folders scan noticeably faster on multi-core machines.
- ⏱️ **Per-file timeout guard** — a single huge or malformed PDF can no longer stall the whole scan; if one file takes longer than the configured timeout, it's flagged `Scan Timeout` and the scan moves on.
- 💾 **Skip-unchanged-files caching** — re-scanning the same tree after fixing a handful of files reuses cached results for anything whose size and modified-time haven't changed, instead of re-checking everything. Toggle or clear it under Settings → General.
- 🔑 **Password-protected PDF detection** — encrypted PDFs are now flagged as their own "Password Protected" error, separated out from the generic "Cannot Open" bucket.
- 🧬 **Duplicate detection** — file contents are hashed to flag the same PDF appearing in more than one scanned location (e.g. both `open` and `confidential`) — its own kind of compliance issue.
- 🗒️ **Scan history log** — a new History button records date, folder, mode, files checked, errors found, and duration for past runs, so you can compare over time without re-scanning.
- 🧾 **Scan profiles** — save and reload named sets of thresholds and target subfolders per client or project type, under Settings → Scan Settings.
- 📖 **Auto-open Word report** — after exporting, optionally open the report immediately, with a "Don't ask again" choice that sticks.
- 🕳️ **Empty-state polish** — the results panel now shows a friendly message before any scan, after a clean scan, or when a filter excludes everything, instead of an unexplained blank box.
- 🪵 **File-based logging** — failures are now logged to `pdf_checker.log` (with rotation) instead of only printing to a console window nobody's watching, so issues are diagnosable after the fact.
- 🚀 **In-app auto-updater** — download and install updates directly from GitHub releases with live progress tracking and automatic restart (via a self-deleting helper script for `.exe` on Windows or script replacement with `.bak` backup when running from source).
- 📱 **Scrollable sidebar & small-screen support** — the left sidebar is now scrollable via mousewheel or vertical scrollbar, and window startup dimensions auto-clamp to fit smaller displays (such as 1366x768 or 1536x864 at 125% scaling) so action buttons are never pushed off-screen.
- 🎯 **Accurate content detection for "Missing Information"** — pages containing embedded raster images, vector graphics/drawings, or interactive form widgets are no longer falsely treated as empty, preventing legitimate scanned documents, certificates, and diagrams from triggering false positives.
- 🔬 **Smart DPI calculation for "Not Clear"** — resolution is now computed against each image's actual rendered bounding box on the page rather than entire page dimensions. Thin decorative rules, gradient lines, and separator strips (< 0.5" or aspect ratio > 15:1) are filtered out, standard nominal screen-resolution graphics (72 DPI) are protected with a tolerance margin, and pages with rich native digital text distinguish standard on-screen graphics (≥ 70 DPI) from low-resolution scans (which strictly enforce the configured DPI threshold).

## Previous update

- 📁 **Flexible folder matching** — subfolder names are no longer hardcoded to `open`/`confidential`. Matching is now case-insensitive, a project folder only needs *one* of the target subfolders (not both), and common naming variations (`opened`, `public`, `conf`, `restricted`, `private`, etc.) are recognized automatically. Target names are configurable in Settings → Scan Settings.
- 🗂️ **Scan Mode toggle** — choose **Project Structure** (the folder-matching behavior above) or **All PDFs (Recursive)**, which ignores subfolder naming entirely and scans every PDF under the selected folder.
- 🔀 **Sortable, filterable results table** — results now show in a proper table (file, parent folder, subfolder, errors) instead of a plain text log. Click any column header to sort; type in the filter box to narrow results live.
- 🖱️ **Open results directly** — double-click a row to open the PDF, or right-click for Open File / Show in Folder / Copy Path.
- ⌨️ **Keyboard shortcuts** — `Enter` starts a scan, `Esc` cancels, `Ctrl+E` exports to Word.
- ⏱️ **Live elapsed time + ETA** during a scan, not just a raw file count.
- 🔔 **Desktop notifications** on scan completion, so you don't have to babysit the window (falls back to a normal popup if the optional `plyer` package isn't installed).
- ⏹️ **Faster cancel** — Cancel now takes effect within the file currently being checked instead of only between files, so it no longer stalls on one large PDF.
- 🔄 **Auto-update checker now points at this repository** and correctly handles releases with no description instead of erroring.
- 🎨 **App icon support** — drop an `icon.ico` (window/taskbar) and/or `icon.png` (About tab) next to the script, or bundle them via PyInstaller, and they'll be picked up automatically.
- 🛠️ Assorted correctness fixes: the Word report's summary table used to silently drop half its rows; the progress bar could cap out early on folders with 100+ PDFs; the missing-information check could flag every file as an error on machines without PyMuPDF installed. All fixed.

## Features

- 🔍 **Multi-folder scanning** — recursively finds project folders containing your configured target subfolders (default: 'open' and 'confidential'), matched case-insensitively and tolerant of naming variations
- 🗂️ **Two scan modes** — strict Project Structure matching, or All PDFs (Recursive) to scan everything under a folder regardless of naming
- ⚡ **Parallel, cached processing** — a configurable worker pool checks files concurrently, with per-file timeouts and a skip-unchanged-files cache for fast re-scans
- 🔑 **Password & duplicate detection** — encrypted PDFs and files that appear more than once (by content hash) are flagged as their own error types
- 🔀 **Sortable, filterable results table** — click-to-sort columns and a live filter box
- 🖱️ **One-click file access** — double-click or right-click any result to open it or reveal it on disk
- 📊 **Detailed reports** — export a professional Word document with error summaries, with optional auto-open
- 🗒️ **Scan history** — a local log of past runs (date, folder, mode, files, errors, duration)
- 🧾 **Scan profiles** — save/load named threshold sets per client or project type
- 🔔 **Desktop notifications** when a scan finishes
- ⚙️ **Configurable** — settings dialog with persistent configuration (DPI threshold, text-length threshold, empty-page ratio, window size, auto-update toggle, parallel workers, per-file timeout, scan cache)
- 🎨 **Modern UI** — clean interface with custom rounded-corner controls, tooltips, and friendly empty states throughout
- 🔄 **In-app auto-updates** — automatic startup check against GitHub releases with background downloading, progress display, and seamless self-restarting installation


## Installation

### Option 1: Run from Source

1. Install Python 3.8 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python "pdf checker pro.py"
   ```

### Option 2: Build EXE (Windows)

1. Install Python 3.8 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the build script:
   ```bash
   build_exe.bat
   ```
4. Find your EXE in the `dist` folder

To give the built EXE a custom icon, place `icon.ico` next to the script (or point PyInstaller's `--icon` flag at it) — the app also looks for `icon.ico`/`icon.png` inside the bundle automatically at runtime for the taskbar icon and the About tab.

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

- **General**: window size (applies immediately on save) and the auto-update-check toggle
- **Scan Settings**: target subfolder names (comma-separated) and whether to match common naming variations, DPI threshold, pages-to-check for the resolution check, minimum text length, and empty-page ratio threshold for the missing-information check
- **About**: version info and live dependency status (scrolls if the content grows)

## Notifications

When a scan finishes, the app sends a non-blocking desktop notification via the optional `plyer` package. If `plyer` isn't installed, it falls back to the original pop-up dialog — no functionality is lost either way, but the notification is less disruptive if you're working in another window while a scan runs.

## Auto-Updates

On startup, the app checks this repository's latest GitHub release. If a newer version is found, it opens an **Update Available** dialog with the release notes and one of two actions:

- If the release has a compatible downloadable asset (a `.exe` when running as a built app, or a `.py` when running from source), you'll see **Download & Install**: the app downloads it in the background with a progress dialog, then — after you confirm — closes and restarts itself with the new version already in place. On Windows, this works by handing off to a small helper script that waits for the app to fully exit, replaces the executable, and relaunches it (the standard pattern most self-updating desktop apps use, since Windows won't let a running `.exe` overwrite itself directly). Running from source, the script file is replaced directly and the app relaunches via a fresh Python process; a `.bak` copy of the previous version is kept alongside it.
- If no compatible asset is found for your platform/build, you'll instead see **Open Release Page**, which opens the release in your browser so you can download it manually.

You can disable the startup check entirely under Settings → General.

**Note:** Ensure `GITHUB_REPO` (near the top of `pdf checker pro.py`) points at your GitHub repository (currently `"Brandon-Morision/pdf-error-checker-pro"`) and publish tagged releases there with the appropriate asset(s) attached (a Windows `.exe` for built distributions, and/or a `.py` for source distributions). If you fork or rename the repository, update `GITHUB_REPO` accordingly.

## Requirements

- Python 3.8+
- `requests` — **required**; the app cannot start without it
- `PyMuPDF` (fitz) — recommended; without it, the resolution check is disabled and the missing-information check falls back to a slower reader
- `PyPDF2` — recommended; used as a fallback for the corrupt-file check and the missing-information check
- `pdfminer.six` — last-resort text extraction fallback if both PyMuPDF and PyPDF2 are unavailable
- `python-docx` — required only to use "Export to Word"
- `Pillow` — optional; improves About-tab icon rendering
- `plyer` — optional; enables non-blocking desktop notifications

See `requirements.txt` for exact version pins and the same core/recommended/optional grouping.

## License

MIT License
