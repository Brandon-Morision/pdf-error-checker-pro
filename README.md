# PDF Error Checker Pro

A professional desktop application for scanning and validating PDF files across multiple project folders.

## What's new in this update

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

- 🔍 **Multi-folder scanning** — recursively finds all folders with 'open' and 'confidential' subfolders
- ⚡ **Fast processing** — background threading keeps the UI responsive
- 🔀 **Sortable, filterable results table** — click-to-sort columns and a live filter box
- 🖱️ **One-click file access** — double-click or right-click any result to open it or reveal it on disk
- 📊 **Detailed reports** — export a professional Word document with error summaries
- 🔔 **Desktop notifications** when a scan finishes
- ⚙️ **Configurable** — settings dialog with persistent configuration (DPI threshold, text-length threshold, empty-page ratio, window size, auto-update toggle)
- 🎨 **Modern UI** — clean interface with custom rounded-corner controls and tooltips throughout

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
2. The app automatically finds all subfolders containing both 'open' and 'confidential' folders
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
- **Scan Settings**: DPI threshold, pages-to-check for the resolution check, minimum text length, and empty-page ratio threshold for the missing-information check
- **About**: version info and live dependency status (scrolls if the content grows)

## Notifications

When a scan finishes, the app sends a non-blocking desktop notification via the optional `plyer` package. If `plyer` isn't installed, it falls back to the original pop-up dialog — no functionality is lost either way, but the notification is less disruptive if you're working in another window while a scan runs.

## Auto-Updates

On startup, the app checks this repository's latest GitHub release and offers to open the release page if a newer version is available. This only works once at least one tagged release has been published — until then the check simply finds nothing and the app behaves as up to date. You can disable this check under Settings → General.

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
