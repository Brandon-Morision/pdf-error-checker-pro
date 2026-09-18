"""
pdf_checker_core.py — UI-agnostic scanning engine for PDF Error Checker Pro.

This module knows nothing about Tkinter or Qt. A front end drives it by:
  - constructing Settings / ScanCache / ScanHistory
  - constructing a PDFScanEngine with a `should_continue` callable (polled
    frequently, including mid-file, so a cancel request from any UI takes
    effect quickly without the engine needing to know how that flag is
    stored)
  - calling engine.run_scan(...) from a worker thread, passing callbacks
    for progress/result/status updates

Any UI (the legacy Tkinter app, a PySide6 app, a CLI, or a test suite) can
share this exact engine, so scan behavior and bug fixes only need to exist
in one place.
"""
import os
import sys
import json
import hashlib
import logging
import logging.handlers
import threading
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import requests

try:
    import pymupdf as fitz
    FITZ_AVAILABLE = True
except ImportError:
    try:
        import fitz
        FITZ_AVAILABLE = True
    except ImportError:
        FITZ_AVAILABLE = False

try:
    from pdfminer.high_level import extract_text as pdfminer_extract_text
    PDFMINER_AVAILABLE = True
except ImportError:
    PDFMINER_AVAILABLE = False

try:
    from PyPDF2 import PdfReader
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

try:
    from docx import Document
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

GITHUB_REPO = "Brandon-Morision/pdf-error-checker-pro"
CURRENT_VERSION = "0.2.0"  # bumped: this is the Qt-engine rewrite

FOLDER_NAME_GROUPS = [
    {"open", "opened", "openned", "public"},
    {"confidential", "conf", "restricted", "private"},
]


def _persistent_data_dir():
    """Directory for settings/cache/history/log files.

    When PyInstaller builds a --onefile executable, __file__ for a bundled
    module resolves inside sys._MEIPASS — a temp folder that's extracted
    fresh and deleted on every launch. Anchoring persistent state there
    would silently lose all settings, cache, and history each time the
    app closes. Anchoring to the running .exe's own directory instead
    keeps it next to wherever the user actually placed the built app,
    matching how the source-run version already behaves (files next to
    the script)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _setup_logger():
    logger = logging.getLogger("pdf_checker_pro")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        try:
            log_path = _persistent_data_dir() / "pdf_checker.log"
            handler = logging.handlers.RotatingFileHandler(
                log_path, maxBytes=1_000_000, backupCount=2, encoding="utf-8"
            )
            handler.setFormatter(logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
            ))
            logger.addHandler(handler)
        except Exception:
            logger.addHandler(logging.NullHandler())
    return logger


logger = _setup_logger()


class Settings:
    """Application settings management. Unchanged from the Tkinter version —
    this class was already UI-agnostic."""

    DEFAULT_SETTINGS = {
        "resolution_threshold": 150,
        "check_cannot_open": True,
        "check_not_clear": True,
        "check_missing_info": True,
        "empty_page_threshold": 0.8,
        "min_text_length": 50,
        "max_pages_check_resolution": 5,
        "auto_check_updates": True,
        "target_subfolders": "open, confidential",
        "use_folder_aliases": True,
        "scan_mode": "project",
        "check_password_protected": True,
        "check_duplicates": True,
        "enable_scan_cache": True,
        "max_workers": 4,
        "per_file_timeout_seconds": 30,
        "auto_open_word_report": "ask",
        "profiles": {},
        "last_folder": "",
        "window_width": 1200,
        "window_height": 850,
    }

    def __init__(self, settings_dir=None):
        base = Path(settings_dir) if settings_dir else _persistent_data_dir()
        self.settings_file = base / "pdf_checker_settings.json"
        self.settings = self.load()

    def load(self):
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    return {**self.DEFAULT_SETTINGS, **saved}
            except Exception:
                return self.DEFAULT_SETTINGS.copy()
        return self.DEFAULT_SETTINGS.copy()

    def save(self, settings_dict):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings_dict, f, indent=2)
            return True
        except Exception as e:
            logger.error(f"Error saving settings: {e}")
            return False

    def get(self, key, default=None):
        return self.settings.get(key, default)

    def set(self, key, value):
        self.settings[key] = value


class ScanCache:
    """Per-file result cache keyed by (size, mtime, settings signature).
    Unchanged from the Tkinter version, including the scoped prune_missing
    fix (pruning is restricted to entries under `scanned_roots`)."""

    def __init__(self, cache_dir=None):
        base = Path(cache_dir) if cache_dir else _persistent_data_dir()
        self.cache_file = base / "pdf_checker_cache.json"
        self.data = self._load()
        # Up to 16 worker threads can call get()/set() concurrently during
        # a scan. CPython's GIL already makes individual dict operations
        # atomic, so this isn't fixing a corruption bug — but it removes
        # any doubt about interleaved get-then-set races across threads.
        self._lock = threading.Lock()

    def _load(self):
        if self.cache_file.exists():
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.warning(f"Could not load scan cache, starting fresh: {e}")
                return {}
        return {}

    def save(self):
        try:
            with open(self.cache_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f)
        except Exception as e:
            logger.warning(f"Failed to save scan cache: {e}")

    def get(self, path, size, mtime, signature):
        with self._lock:
            entry = self.data.get(path)
            if not entry:
                return None
            if entry.get("size") != size or entry.get("mtime") != mtime or entry.get("signature") != signature:
                return None
            return entry

    def set(self, path, size, mtime, signature, errors, file_hash=None):
        with self._lock:
            self.data[path] = {
                "size": size, "mtime": mtime, "signature": signature,
                "errors": errors, "hash": file_hash,
            }

    def prune_missing(self, valid_paths, scanned_roots=None):
        if not scanned_roots:
            return
        normalized_roots = [os.path.normcase(os.path.normpath(r)) for r in scanned_roots]
        # Fixed: valid_paths used to be compared as raw strings while
        # _under_scanned_roots below normalizes case — on a case-insensitive
        # filesystem (Windows, and macOS by default), a cached key that
        # differs only in casing from how this run's path was constructed
        # would fail the raw-string comparison and get wrongly pruned even
        # though it's the same, still-actively-scanned file.
        normalized_valid = {os.path.normcase(os.path.normpath(p)) for p in valid_paths}

        def _under_scanned_roots(path):
            norm_path = os.path.normcase(os.path.normpath(path))
            return any(
                norm_path == root or norm_path.startswith(root + os.sep)
                for root in normalized_roots
            )

        with self._lock:
            stale = [
                p for p in self.data
                if os.path.normcase(os.path.normpath(p)) not in normalized_valid and _under_scanned_roots(p)
            ]
            for p in stale:
                del self.data[p]

    def clear(self):
        self.data = {}
        try:
            if self.cache_file.exists():
                self.cache_file.unlink()
        except Exception as e:
            logger.warning(f"Could not delete cache file: {e}")


class ScanHistory:
    """Local log of past scan runs. Unchanged from the Tkinter version."""

    MAX_ENTRIES = 50

    def __init__(self, history_dir=None):
        base = Path(history_dir) if history_dir else _persistent_data_dir()
        self.history_file = base / "pdf_checker_history.json"
        self.entries = self._load()

    def _load(self):
        if self.history_file.exists():
            try:
                with open(self.history_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.warning(f"Could not load scan history, starting fresh: {e}")
                return []
        return []

    def add(self, entry):
        self.entries.insert(0, entry)
        self.entries = self.entries[: self.MAX_ENTRIES]
        self._save()

    def clear(self):
        self.entries = []
        self._save()

    def _save(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.entries, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save scan history: {e}")


class UpdateChecker:
    """Check for updates from GitHub. Unchanged from the Tkinter version —
    already had no GUI dependency."""

    def __init__(self, repo, current_version):
        self.repo = repo
        self.current_version = current_version
        self.api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    def check_for_updates(self):
        try:
            headers = {
                "User-Agent": f"PDF-Error-Checker-Pro/{self.current_version}",
                "Accept": "application/vnd.github.v3+json",
            }
            response = requests.get(self.api_url, headers=headers, timeout=5)
            if response.status_code == 200:
                data = response.json()
                latest_version = (data.get('tag_name') or '').lstrip('v')
                if latest_version:
                    return self.compare_versions(latest_version, self.current_version), latest_version, data
            elif response.status_code != 404:
                logger.info(f"Update check returned HTTP status {response.status_code}")
            return False, None, None
        except Exception as e:
            logger.warning(f"Update check failed: {e}")
            return False, None, None

    def compare_versions(self, latest, current):
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            for l, c in zip(latest_parts, current_parts):
                if l > c:
                    return True
                elif l < c:
                    return False
            return len(latest_parts) > len(current_parts)
        except Exception:
            return False

    @staticmethod
    def find_installable_asset(release_data):
        """Picks a release asset this build can actually install
        automatically, or None if there isn't one — in which case the
        caller should fall back to opening the release page for the user
        to download manually.

        - Running as a frozen PyInstaller .exe on Windows: look for a
          '.exe' asset (self-replace via a relaunching helper script).
          PyInstaller bundles the whole app (including pdf_checker_core.py)
          into that one file, so replacing it is a complete, consistent
          update.
        - Running from source: no automatic path is supported. The app is
          now two files (pdf_checker_qt.py + pdf_checker_core.py) rather
          than the single-file Tkinter original this mechanism was
          designed for. A release asset can only be one file, so
          auto-installing just the entry-point script while a release also
          changed pdf_checker_core.py would silently leave the engine file
          out of sync — a real version-mismatch risk, not a hypothetical
          one. Source installs always fall back to "Open Release Page" for
          a manual update until source releases ship a full-directory
          bundle (e.g. a .zip) with dedicated extraction support.
        - Anything else (a frozen build on macOS/Linux, or a release with
          no matching asset): no automatic path either.
        """
        assets = release_data.get("assets") or []
        is_frozen = getattr(sys, "frozen", False)

        if not (is_frozen and sys.platform.startswith("win")):
            return None
        wanted_ext = ".exe"

        candidates = [a for a in assets if a.get("name", "").lower().endswith(wanted_ext)]
        if not candidates:
            return None
        for a in candidates:
            if "win" in a.get("name", "").lower():
                return a
        return candidates[0]


def download_asset(download_url, dest_path, on_progress: Callable[[float], None] = lambda p: None,
                   user_agent: str = f"PDF-Error-Checker-Pro/{CURRENT_VERSION}", timeout=30,
                   size_hint: int = 0):
    """Streams a release asset to disk, calling on_progress(fraction) as
    bytes arrive.

    `size_hint` lets the caller pass a size it already knows (e.g. the
    "size" field from the GitHub release's asset metadata, available
    before the download even starts) to use if the live HTTP response
    itself doesn't send a Content-Length header. If neither is available,
    on_progress is never called — the caller should show an indeterminate
    progress indicator in that case rather than a bar frozen at 0%.
    Raises on failure — the caller decides how to surface that (a dialog,
    a log line, a CLI message)."""
    dest_path = Path(dest_path)
    with requests.get(download_url, headers={"User-Agent": user_agent}, stream=True, timeout=timeout) as resp:
        resp.raise_for_status()
        total_size = int(resp.headers.get("Content-Length") or 0) or int(size_hint or 0)
        downloaded = 0
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size:
                        on_progress(downloaded / total_size)
    if downloaded == 0:
        raise IOError("Downloaded file is empty.")
    logger.info(f"Downloaded update to {dest_path} ({downloaded} bytes)")
    return dest_path


def apply_windows_exe_update(new_exe_path, wait_attempts=30):
    """Replaces the running .exe with the downloaded one. Windows won't let
    a running executable overwrite itself directly, so this writes a tiny
    helper batch script that: waits (capped at wait_attempts, not
    forever) for the current process to release its file lock, moves the
    new exe into place, relaunches it, then deletes itself. Hands off to
    that script and returns — the CALLER is responsible for exiting the
    current process (e.g. QApplication.quit() / sys.exit()) immediately
    after this returns, since the old exe can't be replaced while still
    running."""
    import subprocess
    import tempfile

    current_exe = Path(sys.executable)
    new_exe_path = Path(new_exe_path)
    batch_path = Path(tempfile.gettempdir()) / "pdf_checker_update.bat"

    batch_script = f"""@echo off
chcp 65001 >NUL
setlocal enabledelayedexpansion
set ATTEMPTS=0
:wait_loop
del "{current_exe}" >NUL 2>&1
if exist "{current_exe}" (
    set /a ATTEMPTS+=1
    REM Delayed expansion (!VAR! not %VAR%) is required here: inside a
    REM parenthesized block, %ATTEMPTS% is substituted when the block is
    REM PARSED - i.e. before the "set /a" line above it executes - so the
    REM guard would compare a stale value and the cap would be off by one.
    REM !ATTEMPTS! reads the value at execution time instead.
    if !ATTEMPTS! GEQ {wait_attempts} (
        echo Could not replace "{current_exe}" - it may still be running or locked.
        echo The downloaded update is still available at "{new_exe_path}".
        pause
        exit /b 1
    )
    timeout /t 1 /nobreak >NUL
    goto wait_loop
)
copy /Y "{new_exe_path}" "{current_exe}" >NUL
del "{new_exe_path}" >NUL 2>&1
start "" "{current_exe}"
(goto) 2>nul & del "%~f0"
"""
    batch_path.write_text(batch_script, encoding="utf-8")
    logger.info(f"Launching update helper script: {batch_path}")
    subprocess.Popen(["cmd", "/c", str(batch_path)],
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def apply_source_update(new_script_path, current_script_path):
    """Running from source: overwrite the running script's file with the
    downloaded one and relaunch. No file-lock workaround is needed —
    Python reads the script into memory at startup and doesn't keep the
    file open. The CALLER is responsible for exiting the current process
    immediately after this returns.

    current_script_path must be the front end's own __file__ (e.g. the
    Qt app's pdf_checker_qt.py) — it is NOT inferred from this module's
    __file__, which would wrongly point at pdf_checker_core.py itself."""
    import subprocess
    import shutil

    current_script = Path(current_script_path).resolve()
    new_script_path = Path(new_script_path)
    backup_path = current_script.with_suffix(current_script.suffix + ".bak")
    shutil.copy2(current_script, backup_path)
    shutil.copy2(new_script_path, current_script)
    logger.info(f"Replaced {current_script} with downloaded update (backup at {backup_path})")
    subprocess.Popen([sys.executable, str(current_script)])


@dataclass
class ScanResult:
    path: str
    folder: str
    parent: str
    filename: str
    errors: list = field(default_factory=list)


@dataclass
class ScanProgress:
    index: int
    total: int
    filename: str
    parent: str
    elapsed_seconds: float
    eta_seconds: Optional[float]


@dataclass
class ScanSummary:
    total_checked: int
    total_errors: int
    unique_parents: int
    duration_seconds: float
    cache_hits: int
    timeouts: int
    cancelled: bool


class PDFScanEngine:
    """The actual PDF-checking logic, ported from the Tkinter app with
    every `self.running` check replaced by `self.should_continue()`. Every
    fix from the Tkinter version's review passes is preserved: exceptions
    in has_missing_information flag the file rather than pass it; cache
    pruning is scope-limited; per-file timeouts apply to the duplicate-hash
    pass too, not just the main check pass.
    """

    def __init__(self, settings: Settings, cache: ScanCache, should_continue: Callable[[], bool]):
        self.settings = settings
        self.cache = cache
        self.should_continue = should_continue

    # ---- folder / target-subfolder matching -----------------------------

    def get_target_subfolder_names(self):
        raw = self.settings.get("target_subfolders", "open, confidential")
        names = [n.strip() for n in raw.split(",") if n.strip()]
        return names or ["open", "confidential"]

    @staticmethod
    def expand_folder_aliases(name_lower):
        for group in FOLDER_NAME_GROUPS:
            if name_lower in group:
                return group
        return {name_lower}

    def match_subfolders(self, dirs):
        targets = self.get_target_subfolder_names()
        use_aliases = self.settings.get("use_folder_aliases", True)
        dirs_lower_map = {}
        for d in dirs:
            dirs_lower_map.setdefault(d.lower(), d)
        matches = []
        seen_actual = set()
        for target in targets:
            target_lower = target.lower()
            candidates = self.expand_folder_aliases(target_lower) if use_aliases else {target_lower}
            for cand in candidates:
                if cand in dirs_lower_map:
                    actual = dirs_lower_map[cand]
                    if actual not in seen_actual:
                        matches.append((target, actual))
                        seen_actual.add(actual)
                    break
        return matches

    def find_valid_folders(self, root_folder, scan_mode):
        """Returns (valid_folders, folder_matches) for project mode, or
        ([root_folder], {}) for 'all' mode."""
        if scan_mode == "all":
            return [root_folder], {}
        valid_folders, folder_matches = [], {}
        for root, dirs, _files in os.walk(root_folder):
            matches = self.match_subfolders(dirs)
            if matches:
                valid_folders.append(root)
                folder_matches[root] = matches
        return valid_folders, folder_matches

    # ---- per-file checks --------------------------------------------------

    def is_pdf_password_protected(self, pdf_path):
        if FITZ_AVAILABLE:
            doc = None
            try:
                doc = fitz.open(pdf_path)
                return doc.needs_pass
            except Exception:
                pass
            finally:
                if doc is not None:
                    doc.close()
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    return bool(PdfReader(f).is_encrypted)
            except Exception:
                pass
        return False

    def is_pdf_corrupt(self, pdf_path):
        if FITZ_AVAILABLE:
            doc = None
            try:
                doc = fitz.open(pdf_path)
                # Fixed: this used to fall through to the PyPDF2 check below
                # even after PyMuPDF successfully opened the file, so a
                # valid PDF that merely made PyPDF2 raise a warning/exception
                # (common on modern PDFs with newer xref-stream compression)
                # got wrongly flagged as "Cannot Open". PyMuPDF succeeding is
                # authoritative on its own; PyPDF2 is a fallback for when
                # PyMuPDF isn't available, not a second opinion to override it.
                return len(doc) == 0
            except Exception:
                return True
            finally:
                # try/finally (not doc.close() inline in the try body):
                # if len(doc) itself raised, an inline close would be
                # skipped, leaking the file handle — a real risk on Windows,
                # where an open MuPDF handle can block a later rename/delete
                # of that same file until the process exits.
                if doc is not None:
                    doc.close()
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    if len(PdfReader(f).pages) == 0:
                        return True
            except Exception:
                return True
        return False

    def is_pdf_not_clear(self, pdf_path, resolution_threshold):
        if not FITZ_AVAILABLE:
            return False
        doc = None
        try:
            doc = fitz.open(pdf_path)
            max_pages = min(self.settings.get("max_pages_check_resolution", 5), len(doc))
            min_text = self.settings.get("min_text_length", 50)
            for page_num in range(max_pages):
                if not self.should_continue():
                    return False
                page = doc[page_num]
                page_text = page.get_text() or ""
                has_rich_text = len(page_text.strip()) >= min_text
                page_rect = page.rect

                for img in page.get_images(full=True):
                    if not self.should_continue():
                        return False
                    xref = img[0]
                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            continue
                        width = base_image.get("width", 0)
                        height = base_image.get("height", 0)
                        if width <= 0 or height <= 0 or width < 10 or height < 10:
                            continue

                        rects = page.get_image_rects(xref)
                        if rects:
                            for rect in rects:
                                if rect.width < 36 or rect.height < 36 or rect.width <= 0 or rect.height <= 0:
                                    continue
                                aspect = max(rect.width, rect.height) / max(min(rect.width, rect.height), 0.1)
                                if aspect > 15:
                                    continue
                                dpi_x = width / (rect.width / 72.0)
                                dpi_y = height / (rect.height / 72.0)
                                dpi = min(dpi_x, dpi_y)
                                effective_thresh = (min(resolution_threshold, 70) if has_rich_text
                                                     else resolution_threshold * 0.95)
                                if dpi < effective_thresh:
                                    return True
                        elif width >= 400 or height >= 400:
                            page_w_inch = page_rect.width / 72.0
                            page_h_inch = page_rect.height / 72.0
                            if page_w_inch > 0 and page_h_inch > 0:
                                dpi = min(width / page_w_inch, height / page_h_inch)
                                eff = (min(resolution_threshold, 70) if has_rich_text
                                       else resolution_threshold * 0.95)
                                if dpi < eff:
                                    return True
                    except Exception:
                        continue
            return False
        except Exception as e:
            logger.warning(f"Could not check image resolution for {pdf_path}: {e}")
            return False
        finally:
            # try/finally instead of a doc.close() at each return point:
            # code outside the inner per-image try/except above (e.g.
            # page.get_text(), page.rect, doc[page_num] itself) could raise
            # and jump straight past every one of those inline closes,
            # leaking the handle. finally always runs regardless of which
            # return path was taken.
            if doc is not None:
                doc.close()

    def has_missing_information(self, pdf_path):
        threshold = self.settings.get("empty_page_threshold", 0.8)
        min_text = self.settings.get("min_text_length", 50)

        if FITZ_AVAILABLE:
            text_content, page_count, empty_pages = "", 0, 0
            total_images = total_drawings = total_widgets = 0
            doc = None
            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)
                if page_count == 0:
                    return True
                for page in doc:
                    if not self.should_continue():
                        return False
                    page_text = page.get_text() or ""
                    clean_text = page_text.strip()
                    if clean_text:
                        text_content += page_text
                    num_images = len(page.get_images() or [])
                    total_images += num_images
                    num_drawings = len(page.get_drawings() or [])
                    total_drawings += num_drawings
                    try:
                        num_widgets = len(list(page.widgets() or []))
                    except Exception:
                        num_widgets = 0
                    total_widgets += num_widgets
                    if not (clean_text or num_images or num_drawings or num_widgets):
                        empty_pages += 1
            except Exception as e:
                # Flag on inspection failure — consistent with is_pdf_corrupt.
                logger.warning(f"Could not inspect content of {pdf_path}: {e}")
                return True
            finally:
                # Guarantees the handle is closed on every exit path above
                # (early "return True"/"return False", the exception branch,
                # or falling through normally) — a bare doc.close() placed
                # inline at each return point would be skipped by any
                # exception raised outside the inner widgets() try/except,
                # e.g. page.get_text() or page.get_images() itself.
                if doc is not None:
                    doc.close()
        else:
            text_content, page_count, empty_pages, total_images = self._extract_text_fallback(pdf_path)
            total_drawings = total_widgets = 0

        if page_count == 0:
            return True
        if empty_pages > page_count * threshold:
            return True
        has_visual_data = (total_images > 0) or (total_drawings > 0) or (total_widgets > 0)
        if not has_visual_data and len(text_content.strip()) < min_text:
            return True
        return False

    def _extract_text_fallback(self, pdf_path):
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    reader = PdfReader(f)
                    page_count = len(reader.pages)
                    if page_count == 0:
                        return "", 0, 0, 0
                    text_content, empty_pages, total_images = "", 0, 0
                    for page in reader.pages:
                        if not self.should_continue():
                            return text_content, page_count, empty_pages, total_images
                        try:
                            page_text = page.extract_text() or ""
                        except Exception:
                            page_text = ""
                        try:
                            page_images = len(page.images)
                        except Exception:
                            page_images = 0
                        total_images += page_images
                        if page_text.strip() or page_images:
                            text_content += page_text
                        else:
                            empty_pages += 1
                    return text_content, page_count, empty_pages, total_images
            except Exception:
                pass
        if PDFMINER_AVAILABLE:
            try:
                text = pdfminer_extract_text(pdf_path) or ""
                page_count = 1 if text.strip() else 0
                empty_pages = 0 if text.strip() else 1
                return text, page_count, empty_pages, 0
            except Exception:
                pass
        return "", 0, 0, 0

    def _hash_file(self, path, chunk_size=1024 * 1024):
        try:
            hasher = hashlib.sha256()
            with open(path, "rb") as f:
                while True:
                    if not self.should_continue():
                        return None
                    chunk = f.read(chunk_size)
                    if not chunk:
                        break
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception as e:
            logger.warning(f"Could not hash file {path}: {e}")
            return None

    def check_pdf(self, pdf_path, check_cannot_open, check_not_clear, check_missing_info,
                  check_password_protected, resolution_threshold):
        errors = []
        if not self.should_continue():
            return errors

        is_protected = False
        if check_password_protected:
            is_protected = self.is_pdf_password_protected(pdf_path)
            if is_protected:
                errors.append("Password Protected")
        if not self.should_continue() or is_protected:
            return errors

        if check_cannot_open and self.is_pdf_corrupt(pdf_path):
            errors.append("Cannot Open")
            return errors
        if not self.should_continue():
            return errors
        if check_not_clear and self.is_pdf_not_clear(pdf_path, resolution_threshold):
            errors.append("Not Clear")
        if not self.should_continue():
            return errors
        if check_missing_info and self.has_missing_information(pdf_path):
            errors.append("Missing Information")
        return errors

    def _check_pdf_cached(self, pdf_path, check_cannot_open, check_not_clear, check_missing_info,
                          check_password_protected, resolution_threshold, cache_enabled,
                          signature, need_hash=False):
        size, mtime = None, None
        try:
            stat = os.stat(pdf_path)
            size, mtime = stat.st_size, stat.st_mtime
        except OSError as e:
            logger.warning(f"Could not stat {pdf_path}: {e}")

        if cache_enabled and size is not None:
            cached = self.cache.get(pdf_path, size, mtime, signature)
            if cached is not None:
                file_hash = cached.get("hash")
                if need_hash and file_hash is None:
                    file_hash = self._hash_file(pdf_path)
                return cached.get("errors", []), file_hash, True

        errors = self.check_pdf(pdf_path, check_cannot_open, check_not_clear,
                                check_missing_info, check_password_protected, resolution_threshold)
        file_hash = self._hash_file(pdf_path) if need_hash else None
        if cache_enabled and size is not None and self.should_continue():
            self.cache.set(pdf_path, size, mtime, signature, errors, file_hash)
        return errors, file_hash, False

    # ---- duplicate detection ----------------------------------------------

    def _detect_duplicates(self, all_pdfs, per_file_timeout, on_duplicate, on_status=lambda s: None):
        needs_hash = [p for p in all_pdfs if p.get("_hash") is None]
        if needs_hash:
            total = len(needs_hash)
            done = 0
            with ThreadPoolExecutor(max_workers=4, thread_name_prefix="pdfhash") as executor:
                future_map = {executor.submit(self._hash_file, p["path"]): p for p in needs_hash}
                for future, pdf_info in future_map.items():
                    if not self.should_continue():
                        break
                    try:
                        pdf_info["_hash"] = future.result(timeout=per_file_timeout)
                    except FuturesTimeoutError:
                        pdf_info["_hash"] = None
                    except Exception:
                        pdf_info["_hash"] = None
                    done += 1
                    on_status(f"Checking for duplicates... ({done}/{total})")

        hash_map = {}
        for pdf_info in all_pdfs:
            if not self.should_continue():
                return
            file_hash = pdf_info.get("_hash")
            if file_hash is None:
                continue
            hash_map.setdefault(file_hash, []).append(pdf_info)

        for group in hash_map.values():
            if len(group) < 2:
                continue
            for pdf_info in group:
                others = [p for p in group if p is not pdf_info]
                other_desc = ", ".join(f"{p['parent']}/{p['folder']}/{p['filename']}" for p in others[:3])
                if len(others) > 3:
                    other_desc += f", +{len(others) - 3} more"
                on_duplicate(pdf_info, f"Duplicate (also in {other_desc})")

    # ---- the main entry point ----------------------------------------------

    def run_scan(self, root_folder, scan_mode, check_flags, resolution_threshold,
                 on_progress: Callable[[ScanProgress], None],
                 on_result: Callable[[ScanResult], None],
                 on_status: Callable[[str], None] = lambda s: None,
                 on_result_updated: Callable[[ScanResult], None] = lambda r: None):
        """Runs synchronously on the calling thread — the caller is
        responsible for running this in a worker thread (a QThread, a
        Tkinter background thread, etc.) so the UI stays responsive.
        Returns a ScanSummary when done (including when cancelled)."""
        check_cannot_open = check_flags.get("cannot_open", True)
        check_not_clear = check_flags.get("not_clear", True)
        check_missing_info = check_flags.get("missing_info", True)
        check_password_protected = check_flags.get("password_protected", True)
        check_duplicates = check_flags.get("duplicates", True)

        cache_enabled = self.settings.get("enable_scan_cache", True)
        max_workers = max(1, int(self.settings.get("max_workers", 4)))
        per_file_timeout = self.settings.get("per_file_timeout_seconds", 30)

        SCAN_ENGINE_VERSION = 4  # bumped for the Qt-engine port
        check_signature = [
            SCAN_ENGINE_VERSION, check_cannot_open, check_not_clear, check_missing_info,
            check_password_protected, resolution_threshold,
            self.settings.get("max_pages_check_resolution", 5),
            self.settings.get("min_text_length", 50), self.settings.get("empty_page_threshold", 0.8),
        ]

        valid_folders, folder_matches = self.find_valid_folders(root_folder, scan_mode)

        all_pdfs = []
        if scan_mode == "all":
            base_folder = valid_folders[0]
            base_name = os.path.basename(os.path.normpath(base_folder)) or base_folder
            for root, _, files in os.walk(base_folder):
                rel_dir = os.path.relpath(root, base_folder)
                subfolder_label = "(root)" if rel_dir == "." else rel_dir.replace("\\", "/")
                for file in files:
                    if file.lower().endswith(".pdf"):
                        all_pdfs.append({
                            "path": os.path.join(root, file), "folder": subfolder_label,
                            "parent": base_name, "filename": file,
                        })
        else:
            for folder in valid_folders:
                parent_name = os.path.basename(folder)
                for _, subfolder in folder_matches.get(folder, []):
                    subfolder_path = os.path.join(folder, subfolder)
                    if not os.path.exists(subfolder_path):
                        continue
                    for root, _, files in os.walk(subfolder_path):
                        for file in files:
                            if file.lower().endswith(".pdf"):
                                all_pdfs.append({
                                    "path": os.path.join(root, file), "folder": subfolder,
                                    "parent": parent_name, "filename": file,
                                })

        if not all_pdfs:
            return ScanSummary(0, 0, 0, 0.0, 0, 0, cancelled=not self.should_continue()), []

        total_pdfs = len(all_pdfs)
        start_time = datetime.now()
        logger.info(f"Scan started: mode={scan_mode}, files={total_pdfs}, workers={max_workers}")

        results = []
        cache_hits = timeouts = 0

        executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="pdfcheck")
        futures = []
        try:
            for pdf_info in all_pdfs:
                if not self.should_continue():
                    break
                future = executor.submit(
                    self._check_pdf_cached, pdf_info["path"], check_cannot_open, check_not_clear,
                    check_missing_info, check_password_protected, resolution_threshold,
                    cache_enabled, check_signature, check_duplicates,
                )
                futures.append((future, pdf_info))

            for i, (future, pdf_info) in enumerate(futures):
                if not self.should_continue():
                    break
                elapsed = (datetime.now() - start_time).total_seconds()
                avg_per_file = elapsed / max(i, 1) if i > 0 else 0
                eta = avg_per_file * (total_pdfs - i) if i > 0 else None
                on_progress(ScanProgress(i + 1, total_pdfs, pdf_info["filename"], pdf_info["parent"],
                                         elapsed, eta))
                try:
                    errors, file_hash, used_cache = future.result(timeout=per_file_timeout)
                except FuturesTimeoutError:
                    errors, file_hash, used_cache = ["Scan Timeout"], None, False
                    timeouts += 1
                    logger.warning(f"Per-file timeout exceeded: {pdf_info['path']}")
                except Exception as e:
                    errors, file_hash, used_cache = [], None, False
                    logger.error(f"Error checking {pdf_info['path']}: {e}")

                pdf_info["_hash"] = file_hash
                if used_cache:
                    cache_hits += 1
                if errors and self.should_continue():
                    result = ScanResult(pdf_info["path"], pdf_info["folder"], pdf_info["parent"],
                                        pdf_info["filename"], errors)
                    results.append(result)
                    on_result(result)
        finally:
            for future, _ in futures:
                future.cancel()
            # Deterministic, bounded teardown: wait for the pool's worker
            # threads to actually finish here, in our own code, rather than
            # shutdown(wait=False) — which defers real cleanup to Python's
            # atexit thread-join machinery, whose timing (and interaction
            # with a Qt background thread's own event loop) is out of our
            # control. This runs on the calling thread (a background
            # QThread, not the GUI thread), so blocking briefly here does
            # not freeze the UI.
            try:
                executor.shutdown(wait=True, cancel_futures=True)
            except TypeError:
                # cancel_futures was added in Python 3.9; fall back cleanly
                # on older interpreters.
                executor.shutdown(wait=True)

        if check_duplicates and self.should_continue():
            on_status("Checking for duplicates...")

            def _on_duplicate(pdf_info, label):
                existing = next((r for r in results if r.path == pdf_info["path"]), None)
                if existing:
                    if label not in existing.errors:
                        existing.errors.append(label)
                        # Fixed: this mutated the result in place but never
                        # told the caller anything changed, so a UI backed
                        # by a data model (Qt's QAbstractTableModel, e.g.)
                        # had no way to know that row needed re-rendering —
                        # the duplicate error was silently invisible on
                        # screen even though it was present in the object
                        # and correctly showed up in the exported Word report.
                        on_result_updated(existing)
                    return
                result = ScanResult(pdf_info["path"], pdf_info["folder"], pdf_info["parent"],
                                    pdf_info["filename"], [label])
                results.append(result)
                on_result(result)

            self._detect_duplicates(all_pdfs, per_file_timeout, _on_duplicate, on_status=on_status)

        try:
            self.cache.prune_missing({p["path"] for p in all_pdfs}, scanned_roots=valid_folders)
            self.cache.save()
        except Exception as e:
            logger.warning(f"Failed to persist scan cache: {e}")

        duration = (datetime.now() - start_time).total_seconds()
        cancelled = not self.should_continue()
        summary = ScanSummary(
            total_checked=total_pdfs, total_errors=len(results),
            unique_parents=len({r.parent for r in results}), duration_seconds=duration,
            cache_hits=cache_hits, timeouts=timeouts, cancelled=cancelled,
        )
        logger.info(f"Scan {'cancelled' if cancelled else 'completed'}: {summary}")
        return summary, results


def export_results_to_word(doc_path, results, folder_path, resolution_threshold, check_flags):
    """Builds the Word report. Unchanged logic from the Tkinter version,
    including the fixed summary table (sized to its actual row count)."""
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is required. Install with: pip install python-docx")

    doc = Document()
    title = doc.add_heading("PDF Error Checker Report", level=1)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"Scanned Parent Folder: {folder_path}")

    doc.add_heading("Scan Settings", level=2)
    settings_para = doc.add_paragraph()
    settings_para.add_run("Error Types: ").bold = True
    enabled = [label for key, label in [
        ("cannot_open", "Cannot Open"), ("password_protected", "Password Protected"),
        ("not_clear", "Not Clear"), ("missing_info", "Missing Information"),
        ("duplicates", "Duplicate Detection"),
    ] if check_flags.get(key)]
    settings_para.add_run(", ".join(enabled) if enabled else "None")
    doc.add_paragraph(f"Resolution Threshold: {resolution_threshold} DPI")

    doc.add_heading("Summary", level=2)
    total_checked = len(results)
    counts = {
        "Cannot Open": sum(1 for r in results if "Cannot Open" in r.errors),
        "Password Protected": sum(1 for r in results if "Password Protected" in r.errors),
        "Not Clear (Low Resolution)": sum(1 for r in results if "Not Clear" in r.errors),
        "Missing Information": sum(1 for r in results if "Missing Information" in r.errors),
        "Duplicate PDFs": sum(1 for r in results if any(e.startswith("Duplicate") for e in r.errors)),
    }
    unique_parents = {r.parent for r in results}
    summary_rows = [["Total PDFs With Errors", str(total_checked)],
                    ["Parent Folders Affected", str(len(unique_parents))]]
    summary_rows += [[label, str(count)] for label, count in counts.items()]

    summary_table = doc.add_table(rows=len(summary_rows), cols=2)
    summary_table.style = "Table Grid"
    for row_idx, (label, value) in enumerate(summary_rows):
        summary_table.cell(row_idx, 0).text = label
        summary_table.cell(row_idx, 1).text = value
        summary_table.cell(row_idx, 0).paragraphs[0].runs[0].bold = True

    doc.add_heading("Detailed Results", level=2)
    for result in results:
        doc.add_heading(result.filename, level=3)
        info = doc.add_paragraph()
        info.add_run("Path: ").bold = True
        info.add_run(f"{result.path}\n")
        info.add_run("Parent: ").bold = True
        info.add_run(f"{result.parent}\n")
        info.add_run("Subfolder: ").bold = True
        info.add_run(f"{result.folder}\n")
        info.add_run("Errors: ").bold = True
        info.add_run(", ".join(result.errors))

    doc.save(doc_path)
    logger.info(f"Word report saved: {doc_path}")
