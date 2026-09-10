import os
import sys
import subprocess
import threading
import json
import webbrowser
import hashlib
import logging
import logging.handlers
import tempfile
import shutil
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import requests

# Use pymupdf instead of deprecated fitz
try:
    import pymupdf as fitz
    FITZ_AVAILABLE = True
except ImportError:
    try:
        import fitz
        FITZ_AVAILABLE = True
    except ImportError:
        FITZ_AVAILABLE = False

# Optional imports with fallbacks
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

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from plyer import notification as desktop_notification
    PLYER_AVAILABLE = True
except ImportError:
    PLYER_AVAILABLE = False

# GitHub repo for updates
GITHUB_REPO = "Brandon-Morision/pdf-error-checker-pro"  # Format: username/repo
CURRENT_VERSION = "0.1.8"

# Known naming variations for the two target subfolders. Matching is always
# case-insensitive; when "use_folder_aliases" is enabled in Settings, typing
# any name in a group (in either the target-subfolders setting or on disk)
# matches every other name in that same group. This keeps the app working
# when a department renames "confidential" to "restricted", etc., without
# requiring a code change.
FOLDER_NAME_GROUPS = [
    {"open", "opened", "openned", "public"},
    {"confidential", "conf", "restricted", "private"},
]


def _setup_logger():
    """File-based logging so failures are diagnosable after the fact —
    especially once the app is running unattended for someone else and
    nobody is watching a console window for stray print() output."""
    logger = logging.getLogger("pdf_checker_pro")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        try:
            log_path = Path(__file__).parent / "pdf_checker.log"
            handler = logging.handlers.RotatingFileHandler(
                log_path, maxBytes=1_000_000, backupCount=2, encoding="utf-8"
            )
            handler.setFormatter(logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
            ))
            logger.addHandler(handler)
        except Exception:
            # If the log file can't be created (e.g. read-only install dir),
            # fall back to a null handler rather than raising on startup.
            logger.addHandler(logging.NullHandler())
    return logger


logger = _setup_logger()


class RoundedButton(tk.Canvas):
    """Custom button with rounded corners and hover effects."""

    def __init__(self, parent, text, command=None, bg="#3498db", fg="white",
                 font=("Segoe UI", 10, "bold"), width=200, height=40,
                 state=tk.NORMAL, hover_factor=20, **kwargs):
        super().__init__(parent, width=width, height=height, highlightthickness=0,
                        bg=parent.cget('bg'), **kwargs)

        self.text = text
        self.command = command
        self.normal_bg = bg
        self.hover_bg = self.lighten_color(bg, hover_factor)
        self.click_bg = self.darken_color(bg, hover_factor)
        # Disabled state is a pale tint of the button's own color rather than
        # one flat gray for every button — the cancel button still reads as
        # "red family" and export still reads as "blue family" while dimmed,
        # which is clearer than a single neutral gray for every action.
        self.disabled_bg = self.tint_color(bg, 0.82)
        self.disabled_fg = "#95a5a6"
        self.fg = fg
        self.font = font
        self.width = width
        self.height = height
        self.radius = 12
        self.enabled = (state != tk.DISABLED)

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)
        self.bind("<ButtonRelease-1>", self.on_release)

        self.draw_button()
        self._update_cursor()

    def _update_cursor(self):
        # Give the button a pointer cursor when it's actually clickable.
        self.config_cursor("hand2" if self.enabled else "arrow")

    def config_cursor(self, cursor):
        try:
            self.configure(cursor=cursor)
        except tk.TclError:
            pass

    def config(self, state=None, **kwargs):
        if state == tk.DISABLED:
            self.enabled = False
            self.draw_button(self.disabled_bg)
            self._update_cursor()
        elif state == tk.NORMAL:
            self.enabled = True
            self.draw_button(self.normal_bg)
            self._update_cursor()

    def configure(self, state=None, **kwargs):
        if state is not None:
            self.config(state=state)
        else:
            tk.Canvas.configure(self, **kwargs)

    def lighten_color(self, color, amount):
        try:
            color = color.lstrip('#')
            lv = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            r, g, b = [min(255, c + amount) for c in lv]
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return color

    def darken_color(self, color, amount):
        try:
            color = color.lstrip('#')
            lv = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            r, g, b = [max(0, c - amount) for c in lv]
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return color

    def tint_color(self, color, toward_white_ratio):
        """Blend a color toward white to get a pale, desaturated version —
        used for the disabled state so it still hints at the button's hue."""
        try:
            color = color.lstrip('#')
            lv = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            r, g, b = [round(c + (255 - c) * toward_white_ratio) for c in lv]
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return color

    def draw_button(self, bg_color=None):
        self.delete("all")
        if bg_color is None:
            bg_color = self.normal_bg if self.enabled else self.disabled_bg
        text_color = self.fg if self.enabled else self.disabled_fg

        self.create_rounded_rect(2, 2, self.width-2, self.height-2, self.radius,
                                fill=bg_color, outline="")
        self.create_text(self.width//2, self.height//2, text=self.text,
                        fill=text_color, font=self.font)

    def create_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1+radius, y1, x2-radius, y1, x2, y1, x2, y1+radius,
            x2, y2-radius, x2, y2, x2-radius, y2, x1+radius, y2,
            x1, y2, x1, y2-radius, x1, y1+radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def on_enter(self, event):
        if self.enabled:
            self.draw_button(self.hover_bg)

    def on_leave(self, event):
        if self.enabled:
            self.draw_button(self.normal_bg)
        else:
            self.draw_button(self.disabled_bg)

    def on_click(self, event):
        if self.enabled:
            self.draw_button(self.click_bg)

    def on_release(self, event):
        if self.enabled:
            self.draw_button(self.hover_bg)
            if self.command:
                self.command()


class Tooltip:
    """Small hover tooltip for any widget — used to explain scan options
    without cluttering the main layout with extra text."""

    def __init__(self, widget, text, delay_ms=450):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self.tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT, background="#2c3e50",
                          foreground="white", relief=tk.SOLID, borderwidth=0,
                          font=("Segoe UI", 8), padx=8, pady=5, wraplength=260)
        label.pack()

    def _hide(self, _event=None):
        self._cancel()
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class Settings:
    """Application settings management."""

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

    def __init__(self):
        self.settings_file = Path(__file__).parent / "pdf_checker_settings.json"
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
    """Persists per-file scan results keyed by (size, mtime), so re-scanning
    the same tree after only fixing a handful of files doesn't have to
    re-run every check on files that haven't changed.

    A cache hit also requires the current scan's settings ("signature") to
    match what was in effect when the entry was recorded — if thresholds or
    which checks are enabled have changed since, a stale hit could report
    the wrong result, so we simply treat that as a miss and re-check.
    """

    def __init__(self):
        self.cache_file = Path(__file__).parent / "pdf_checker_cache.json"
        self.data = self._load()

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
        entry = self.data.get(path)
        if not entry:
            return None
        if entry.get("size") != size or entry.get("mtime") != mtime or entry.get("signature") != signature:
            return None
        return entry

    def set(self, path, size, mtime, signature, errors, file_hash=None):
        self.data[path] = {
            "size": size, "mtime": mtime, "signature": signature,
            "errors": errors, "hash": file_hash,
        }

    def prune_missing(self, valid_paths):
        """Drop entries for files no longer in scope, so the cache file
        doesn't grow forever as folders are renamed or files move/delete."""
        stale = [p for p in self.data if p not in valid_paths]
        for p in stale:
            del self.data[p]


class ScanHistory:
    """A small local log of past scan runs (date, folder, mode, files
    checked, errors found, duration) so results can be compared over time
    without re-scanning."""

    MAX_ENTRIES = 50

    def __init__(self):
        self.history_file = Path(__file__).parent / "pdf_checker_history.json"
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

    def _save(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.entries, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save scan history: {e}")


class ScanHistoryDialog(tk.Toplevel):
    """Read-only view of past scan runs, with an option to clear the log."""

    def __init__(self, parent, history):
        super().__init__(parent)
        self.title("Scan History")
        self.configure(bg="#f5f6fa")
        self.transient(parent)
        self.grab_set()
        self.history = history

        self._setup_ui()

        self.update_idletasks()
        w = max(760, self.winfo_reqwidth())
        h = min(max(400, self.winfo_reqheight()), int(self.winfo_screenheight() * 0.8))
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(620, 350)

    def _setup_ui(self):
        main_frame = tk.Frame(self, bg="#f5f6fa", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        tk.Label(main_frame, text="Scan History", font=("Segoe UI", 16, "bold"),
                 bg="#f5f6fa", fg="#2c3e50").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        columns = ("timestamp", "folder", "mode", "files", "errors", "duration")
        headings = {"timestamp": "Date", "folder": "Folder", "mode": "Mode",
                    "files": "Files Checked", "errors": "Errors Found", "duration": "Duration (s)"}
        widths = {"timestamp": 150, "folder": 260, "mode": 110, "files": 90, "errors": 90, "duration": 90}

        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor=tk.W)
        vsb = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=1, column=0, sticky=tk.NSEW)
        vsb.grid(row=1, column=1, sticky=tk.NS)

        self._populate()

        button_frame = tk.Frame(main_frame, bg="#f5f6fa")
        button_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(15, 0))
        RoundedButton(button_frame, text="Clear History", command=self._clear_history,
                      bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"),
                      width=140, height=36).pack(side=tk.LEFT)
        RoundedButton(button_frame, text="Close", command=self.destroy,
                      bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                      width=110, height=36).pack(side=tk.RIGHT)

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        if not self.history.entries:
            self.tree.insert("", tk.END, values=("No scans recorded yet.", "", "", "", "", ""))
            return
        for entry in self.history.entries:
            mode_label = "All PDFs" if entry.get("scan_mode") == "all" else "Project"
            if entry.get("cancelled"):
                mode_label += " (cancelled)"
            self.tree.insert("", tk.END, values=(
                entry.get("timestamp", ""), entry.get("folder", ""), mode_label,
                entry.get("files_checked", 0), entry.get("errors_found", 0),
                entry.get("duration_seconds", 0),
            ))

    def _clear_history(self):
        if messagebox.askyesno("Clear History", "Remove all recorded scan history?", parent=self):
            self.history.entries = []
            self.history._save()
            self._populate()


class UpdateChecker:
    """Check for updates from GitHub."""

    def __init__(self, repo, current_version):
        self.repo = repo
        self.current_version = current_version
        self.api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    def check_for_updates(self):
        """Check if new version is available."""
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
        """Compare version strings."""
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
        - Running from source (.py) on any OS: look for a '.py' asset
          (direct overwrite + relaunch — no file-lock issues to work
          around, since Python doesn't hold the script file open).
        - Anything else (e.g. a frozen build on macOS/Linux, or a release
          with no matching asset): no automatic path is supported yet.
        """
        assets = release_data.get("assets") or []
        is_frozen = getattr(sys, "frozen", False)

        if is_frozen and sys.platform.startswith("win"):
            wanted_ext = ".exe"
        elif not is_frozen:
            wanted_ext = ".py"
        else:
            return None

        candidates = [a for a in assets if a.get("name", "").lower().endswith(wanted_ext)]
        if not candidates:
            return None
        # Prefer a name that mentions "windows"/"win" when there are several
        # exe assets (e.g. a repo also publishing a mac/linux build), else
        # just take the first match.
        for a in candidates:
            name_lower = a.get("name", "").lower()
            if "win" in name_lower:
                return a
        return candidates[0]


class SettingsDialog(tk.Toplevel):
    """Professional settings dialog with rounded corners."""

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title("Settings")
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#f5f6fa")

        self.settings = settings
        self.modified_settings = settings.settings.copy()
        self._scrollable_canvases = []

        self.setup_ui()
        self._finalize_scrollable_widths()
        # Size the window to fit what's actually inside it (the Scan Settings
        # tab grows over time as options are added, so a hardcoded geometry
        # eventually clips the action buttons below the visible area and the
        # user has to manually resize to see Save/Cancel/Reset). Measuring
        # the real required size after layout, capped to a sane min/max,
        # keeps the buttons visible immediately no matter how tall the
        # tallest tab gets.
        self._apply_auto_size()
        self.center_window()

    def _apply_auto_size(self):
        self.update_idletasks()
        min_w, min_h = 560, 420
        req_w = max(min_w, self.winfo_reqwidth())
        req_h = max(min_h, self.winfo_reqheight())
        # Never spawn taller/wider than the screen — fall back to scrolling
        # within tabs (the About tab already supports this) rather than an
        # oversized window on small displays.
        max_h = int(self.winfo_screenheight() * 0.85)
        max_w = int(self.winfo_screenwidth() * 0.9)
        width = min(req_w, max_w)
        height = min(req_h, max_h)
        self.geometry(f"{width}x{height}")
        self.minsize(min_w, min_h)

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 2
        self.geometry(f"+{x}+{y}")

    def setup_ui(self):
        main_frame = tk.Frame(self, bg="#f5f6fa", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Grid (not pack) for the top-level layout: the button row is a
        # fixed-height grid row that always renders in full, while the
        # notebook is the only row allowed to grow or shrink. This means
        # Save/Cancel/Reset can never end up pushed below the visible
        # window, even as tabs grow (e.g. new Scan Settings options) or the
        # user manually resizes the dialog smaller.
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        tk.Label(main_frame, text="Application Settings", font=("Segoe UI", 18, "bold"),
                bg="#f5f6fa", fg="#2c3e50").grid(row=0, column=0, sticky=tk.W, pady=(0, 15))

        self.notebook = notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, sticky=tk.NSEW, pady=(0, 15))

        general_frame = tk.Frame(notebook, bg="#f5f6fa", padx=20, pady=20)
        notebook.add(general_frame, text="  General  ")
        self.setup_general_tab(general_frame)

        scan_frame = tk.Frame(notebook, bg="#f5f6fa")
        notebook.add(scan_frame, text="  Scan Settings  ")
        self.setup_scan_tab(scan_frame)

        about_frame = tk.Frame(notebook, bg="#f5f6fa")
        notebook.add(about_frame, text="  About  ")
        self.setup_about_tab(about_frame)

        # Action buttons live in their own fixed row directly under the
        # tabs, and are hidden while the About tab is showing — About is
        # read-only reference material, not something to Save/Cancel.
        self.button_frame = tk.Frame(main_frame, bg="#f5f6fa")
        self.button_frame.grid(row=2, column=0, sticky=tk.EW)

        RoundedButton(self.button_frame, text="Reset to Defaults", command=self.reset_defaults,
                      bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                      width=150, height=38).pack(side=tk.LEFT)

        RoundedButton(self.button_frame, text="Save", command=self.save,
                      bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"),
                      width=110, height=38).pack(side=tk.RIGHT)

        RoundedButton(self.button_frame, text="Cancel", command=self.cancel,
                      bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                      width=110, height=38).pack(side=tk.RIGHT, padx=(0, 10))

        notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _on_tab_changed(self, _event=None):
        current_tab_text = self.notebook.tab(self.notebook.select(), "text").strip()
        if current_tab_text == "About":
            # grid_remove (not grid_forget) preserves this widget's grid
            # options so restoring it below doesn't need to respecify them.
            self.button_frame.grid_remove()
        else:
            self.button_frame.grid()

    def _make_scrollable(self, parent, padx=20, pady=20, bg="#f5f6fa"):
        """Wrap a notebook tab's content in a scrolling canvas so it stays
        usable as more sections get added over time, instead of the dialog
        having to keep growing (or clipping content) forever.

        A plain tk.Canvas does NOT propagate its embedded window's natural
        size the way a Frame would, so left alone, the outer dialog's
        auto-sizing (which measures widget request sizes) would undersize
        itself and squeeze this tab's content narrower than it needs. The
        canvas's own width is set explicitly from the inner frame's real
        required width once all of its children exist — see
        _finalize_scrollable_widths(), called after setup_ui() builds every
        tab and before the dialog computes its final geometry.
        """
        canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        inner = tk.Frame(canvas, bg=bg, padx=padx, pady=pady)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _update_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event):
            # Only ever grow the inner frame to fill extra space — never
            # shrink it below its own natural width, which would squeeze
            # or clip its children with no way to reach the rest.
            target_width = max(event.width, inner.winfo_reqwidth())
            canvas.itemconfigure(inner_id, width=target_width)

        inner.bind("<Configure>", _update_scrollregion)
        canvas.bind("<Configure>", _resize_inner)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self._scrollable_canvases.append(canvas)
        return inner

    def _finalize_scrollable_widths(self):
        """Called once, after every tab's content has been built: sizes
        each scrollable tab's canvas to its content's real required width,
        so the dialog's own auto-sizing (measured right after this) accounts
        for the widest tab correctly instead of an arbitrary default."""
        for canvas in self._scrollable_canvases:
            canvas.update_idletasks()
            for item in canvas.find_all():
                inner_name = canvas.itemcget(item, "window")
                if inner_name:
                    inner = canvas.nametowidget(inner_name)
                    canvas.configure(width=inner.winfo_reqwidth())
                    break

    def setup_general_tab(self, parent):
        size_frame = tk.LabelFrame(parent, text="Window Size", font=("Segoe UI", 11, "bold"),
                                   bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        size_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(size_frame, text="Width:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.width_var = tk.IntVar(value=self.settings.get("window_width", 1200))
        tk.Spinbox(size_frame, from_=800, to=2560, textvariable=self.width_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        tk.Label(size_frame, text="Height:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=0, column=2, sticky=tk.W, padx=(20, 0), pady=5)
        self.height_var = tk.IntVar(value=self.settings.get("window_height", 850))
        tk.Spinbox(size_frame, from_=600, to=1440, textvariable=self.height_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=0, column=3, sticky=tk.W, padx=10, pady=5)

        tk.Label(size_frame, text="Applies immediately on save.", font=("Segoe UI", 8, "italic"),
                 bg="#ffffff", fg="#95a5a6").grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))

        update_frame = tk.LabelFrame(parent, text="Updates", font=("Segoe UI", 11, "bold"),
                                    bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        update_frame.pack(fill=tk.X, pady=(0, 15))

        self.auto_update_var = tk.BooleanVar(value=self.settings.get("auto_check_updates", True))
        tk.Checkbutton(update_frame, text="Automatically check for updates on startup",
                      variable=self.auto_update_var, font=("Segoe UI", 10),
                      bg="#ffffff", selectcolor="#ffffff").pack(anchor=tk.W, pady=5)

        cache_frame = tk.LabelFrame(parent, text="Scan Cache", font=("Segoe UI", 11, "bold"),
                                    bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        cache_frame.pack(fill=tk.X, pady=(0, 15))

        self.enable_cache_var = tk.BooleanVar(value=self.settings.get("enable_scan_cache", True))
        cache_cb = tk.Checkbutton(cache_frame,
                      text="Skip re-checking unchanged files (cache by file size + modified time)",
                      variable=self.enable_cache_var, font=("Segoe UI", 10),
                      bg="#ffffff", selectcolor="#ffffff", wraplength=480, justify=tk.LEFT)
        cache_cb.pack(anchor=tk.W, pady=(0, 10))
        Tooltip(cache_cb, "If a file's size and modified-date haven't changed since the last "
                          "scan (with the same settings), its previous result is reused instead "
                          "of re-checking it.")

        RoundedButton(cache_frame, text="Clear Cache", command=self._clear_cache,
                      bg="#95a5a6", fg="white", font=("Segoe UI", 9, "bold"),
                      width=130, height=32).pack(anchor=tk.W)

    def _clear_cache(self):
        cache_path = Path(__file__).parent / "pdf_checker_cache.json"
        try:
            if cache_path.exists():
                cache_path.unlink()
            messagebox.showinfo("Cache Cleared", "The scan cache has been cleared. "
                                "The next scan will re-check every file.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Could not clear cache:\n{e}", parent=self)

    def setup_scan_tab(self, parent):
        parent = self._make_scrollable(parent, padx=20, pady=20)

        if not FITZ_AVAILABLE:
            warn = tk.Label(
                parent,
                text=("PyMuPDF is not installed. Resolution checks will be skipped and "
                      "text-based checks will fall back to a slower reader. "
                      "Install with: pip install pymupdf"),
                font=("Segoe UI", 9), bg="#fdecea", fg="#c0392b", padx=10, pady=8,
                wraplength=520, justify=tk.LEFT,
            )
            warn.pack(fill=tk.X, pady=(0, 15))

        profile_frame = tk.LabelFrame(parent, text="Scan Profiles", font=("Segoe UI", 11, "bold"),
                                      bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        profile_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(profile_frame,
                 text="Save and reload named sets of thresholds and target subfolders for "
                      "different clients or project types.",
                 font=("Segoe UI", 9), bg="#ffffff", fg="#7f8c8d", wraplength=520,
                 justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 8))

        profile_row = tk.Frame(profile_frame, bg="#ffffff")
        profile_row.pack(fill=tk.X)

        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(profile_row, textvariable=self.profile_var, state="readonly",
                                          values=list(self.settings.get("profiles", {}).keys()), width=20)
        self.profile_combo.pack(side=tk.LEFT, padx=(0, 8))

        RoundedButton(profile_row, text="Load", command=self._load_profile,
                      bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"),
                      width=80, height=32).pack(side=tk.LEFT, padx=(0, 6))
        RoundedButton(profile_row, text="Save As...", command=self._save_profile_as,
                      bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold"),
                      width=100, height=32).pack(side=tk.LEFT, padx=(0, 6))
        RoundedButton(profile_row, text="Delete", command=self._delete_profile,
                      bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold"),
                      width=80, height=32).pack(side=tk.LEFT)

        folder_names_frame = tk.LabelFrame(parent, text="Target Subfolders", font=("Segoe UI", 11, "bold"),
                                           bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        folder_names_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(folder_names_frame,
                 text="Comma-separated subfolder names to scan for (used in Project Structure mode):",
                 font=("Segoe UI", 9), bg="#ffffff", fg="#7f8c8d", wraplength=520,
                 justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 8))

        self.target_subfolders_var = tk.StringVar(
            value=self.settings.get("target_subfolders", "open, confidential"))
        target_entry = tk.Entry(folder_names_frame, textvariable=self.target_subfolders_var,
                                font=("Segoe UI", 10))
        target_entry.pack(fill=tk.X, pady=(0, 8))
        Tooltip(target_entry, "e.g. 'open, confidential' or 'public, internal'. "
                              "A project folder is scanned if it contains at least one of these.")

        self.use_aliases_var = tk.BooleanVar(value=self.settings.get("use_folder_aliases", True))
        tk.Checkbutton(folder_names_frame,
                      text="Also match common naming variations (opened, public, conf, restricted, private, etc.)",
                      variable=self.use_aliases_var, font=("Segoe UI", 9), bg="#ffffff", fg="#34495e",
                      selectcolor="#ffffff", wraplength=520, justify=tk.LEFT).pack(anchor=tk.W)

        tk.Label(folder_names_frame,
                 text="Matching is always case-insensitive, and a project folder no longer needs "
                      "every target subfolder to be scanned — just one.",
                 font=("Segoe UI", 8, "italic"), bg="#ffffff", fg="#95a5a6", wraplength=520,
                 justify=tk.LEFT).pack(anchor=tk.W, pady=(8, 0))

        res_frame = tk.LabelFrame(parent, text="Resolution Check", font=("Segoe UI", 11, "bold"),
                                  bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        res_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(res_frame, text="Minimum DPI:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.resolution_var = tk.IntVar(value=self.settings.get("resolution_threshold", 150))
        tk.Spinbox(res_frame, from_=72, to=600, textvariable=self.resolution_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        tk.Label(res_frame, text="Pages to check:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.max_pages_var = tk.IntVar(value=self.settings.get("max_pages_check_resolution", 5))
        tk.Spinbox(res_frame, from_=1, to=20, textvariable=self.max_pages_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)

        text_frame = tk.LabelFrame(parent, text="Missing Information Check", font=("Segoe UI", 11, "bold"),
                                    bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        text_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(text_frame, text="Minimum text length (chars):", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.min_text_var = tk.IntVar(value=self.settings.get("min_text_length", 50))
        tk.Spinbox(text_frame, from_=0, to=1000, textvariable=self.min_text_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)

        tk.Label(text_frame, text="Empty page ratio threshold:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.empty_ratio_var = tk.DoubleVar(value=self.settings.get("empty_page_threshold", 0.8))
        tk.Spinbox(text_frame, from_=0.1, to=1.0, increment=0.05, textvariable=self.empty_ratio_var,
                  font=("Segoe UI", 10), width=10, state="readonly").grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)

        perf_frame = tk.LabelFrame(parent, text="Performance", font=("Segoe UI", 11, "bold"),
                                   bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        perf_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(perf_frame, text="Parallel workers:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.max_workers_var = tk.IntVar(value=self.settings.get("max_workers", 4))
        workers_spin = tk.Spinbox(perf_frame, from_=1, to=16, textvariable=self.max_workers_var,
                  font=("Segoe UI", 10), width=10, state="readonly")
        workers_spin.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        Tooltip(workers_spin, "How many PDFs to check at once. Higher values speed up large "
                              "folders on multi-core machines, but too many can thrash a slow disk.")

        tk.Label(perf_frame, text="Per-file timeout (sec):", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.per_file_timeout_var = tk.IntVar(value=self.settings.get("per_file_timeout_seconds", 30))
        timeout_spin = tk.Spinbox(perf_frame, from_=5, to=300, textvariable=self.per_file_timeout_var,
                  font=("Segoe UI", 10), width=10, state="readonly")
        timeout_spin.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        Tooltip(timeout_spin, "If a single PDF takes longer than this to check, it's marked "
                              "'Scan Timeout' and the scan moves on instead of stalling on it.")

    def _profile_values_from_ui(self):
        return {
            "resolution_threshold": self.resolution_var.get(),
            "max_pages_check_resolution": self.max_pages_var.get(),
            "min_text_length": self.min_text_var.get(),
            "empty_page_threshold": self.empty_ratio_var.get(),
            "target_subfolders": self.target_subfolders_var.get(),
            "use_folder_aliases": self.use_aliases_var.get(),
        }

    def _apply_profile_values(self, values):
        self.resolution_var.set(values.get("resolution_threshold", self.resolution_var.get()))
        self.max_pages_var.set(values.get("max_pages_check_resolution", self.max_pages_var.get()))
        self.min_text_var.set(values.get("min_text_length", self.min_text_var.get()))
        self.empty_ratio_var.set(values.get("empty_page_threshold", self.empty_ratio_var.get()))
        self.target_subfolders_var.set(values.get("target_subfolders", self.target_subfolders_var.get()))
        self.use_aliases_var.set(values.get("use_folder_aliases", self.use_aliases_var.get()))

    def _persist_profiles(self, profiles):
        # Profiles are saved immediately, independent of the dialog's normal
        # Save/Cancel flow — a saved profile shouldn't vanish just because
        # the user later cancels an unrelated threshold edit.
        self.modified_settings["profiles"] = profiles
        self.settings.settings["profiles"] = profiles
        self.settings.save(self.settings.settings)
        self.profile_combo.configure(values=list(profiles.keys()))

    def _save_profile_as(self):
        name = simpledialog.askstring("Save Profile", "Profile name:", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        profiles = dict(self.settings.get("profiles", {}))
        profiles[name] = self._profile_values_from_ui()
        self._persist_profiles(profiles)
        self.profile_var.set(name)
        messagebox.showinfo("Profile Saved", f"Saved profile '{name}'.", parent=self)

    def _load_profile(self):
        name = self.profile_var.get()
        if not name:
            messagebox.showwarning("No Profile Selected", "Choose a profile to load first.", parent=self)
            return
        values = self.settings.get("profiles", {}).get(name)
        if values is None:
            messagebox.showerror("Not Found", f"Profile '{name}' no longer exists.", parent=self)
            return
        self._apply_profile_values(values)

    def _delete_profile(self):
        name = self.profile_var.get()
        if not name:
            return
        if not messagebox.askyesno("Delete Profile", f"Delete profile '{name}'?", parent=self):
            return
        profiles = dict(self.settings.get("profiles", {}))
        profiles.pop(name, None)
        self._persist_profiles(profiles)
        self.profile_var.set("")

    def setup_about_tab(self, parent):
        # Wrapped in a canvas + scrollbar (via the shared helper) so the
        # About content can scroll if it ever grows past the available
        # height, instead of being clipped or forcing the whole dialog taller.
        info_frame = self._make_scrollable(parent, padx=20, pady=20, bg="#ffffff")

        tk.Label(info_frame, text="PDF Error Checker Pro", font=("Segoe UI", 16, "bold"),
                 bg="#ffffff", fg="#2c3e50").pack(pady=(0, 5))

        # Application icon below the title
        self.about_icon = None
        base_dir = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
        icon_path = base_dir / "icon.png"
        if not icon_path.exists():
            icon_path = Path(__file__).parent / "icon.png"
        if not icon_path.exists():
            icon_path = base_dir / "icon.ico"
            if not icon_path.exists():
                icon_path = Path(__file__).parent / "icon.ico"

        if icon_path.exists():
            try:
                if PIL_AVAILABLE:
                    resample = getattr(Image, 'Resampling', Image).LANCZOS
                    img = Image.open(icon_path).convert("RGBA").resize((64, 64), resample)
                    self.about_icon = ImageTk.PhotoImage(img)
                elif str(icon_path).lower().endswith('.png'):
                    raw_img = tk.PhotoImage(file=str(icon_path))
                    sub = max(1, raw_img.width() // 64)
                    self.about_icon = raw_img.subsample(sub, sub)

                if self.about_icon:
                    icon_label = tk.Label(info_frame, image=self.about_icon, bg="#ffffff")
                    icon_label.image = self.about_icon
                    icon_label.pack(pady=(2, 6))
            except Exception as e:
                logger.warning(f"Error loading about icon: {e}")

        tk.Label(info_frame, text=f"Version {CURRENT_VERSION}", font=("Segoe UI", 10),
                 bg="#ffffff", fg="#7f8c8d").pack(pady=(0, 10))

        tk.Label(info_frame, text="Developed by", font=("Segoe UI", 9), bg="#ffffff", fg="#95a5a6").pack()
        tk.Label(info_frame, text="Brandon & Ian", font=("Segoe UI", 13, "bold"), bg="#ffffff",
                 fg="#3498db").pack(pady=(0, 15))

        desc_text = ("A professional tool for scanning and validating PDF files\n"
                    "across multiple project folders.\n\n"
                    "Features:\n"
                    "- Multi-folder recursive scanning\n"
                    "- Thread-safe background processing\n"
                    "- Professional Word report export\n"
                    "- Automatic update checking\n"
                    "- Modern UI with rounded corners")
        tk.Label(info_frame, text=desc_text, font=("Segoe UI", 10), bg="#ffffff", fg="#34495e",
                 justify=tk.CENTER).pack(pady=(0, 20))

        deps_text = "Dependencies:\n"
        deps_text += f"- PyMuPDF: {'OK' if FITZ_AVAILABLE else 'MISSING'}\n"
        deps_text += f"- PyPDF2: {'OK' if PYPDF2_AVAILABLE else 'MISSING'}\n"
        deps_text += f"- pdfminer: {'OK' if PDFMINER_AVAILABLE else 'MISSING'}\n"
        deps_text += f"- python-docx: {'OK' if DOCX_AVAILABLE else 'MISSING'}\n"
        deps_text += f"- plyer (desktop notifications, optional): {'OK' if PLYER_AVAILABLE else 'not installed'}"
        tk.Label(info_frame, text=deps_text, font=("Segoe UI", 9), bg="#ecf0f1", fg="#2c3e50",
                 padx=15, pady=10, justify=tk.LEFT).pack(fill=tk.X, pady=(0, 20))
        tk.Label(info_frame, text="Built with Python & Tkinter", font=("Segoe UI", 9, "italic"),
                 bg="#ffffff", fg="#95a5a6").pack()

    def reset_defaults(self):
        if messagebox.askyesno("Reset Settings", "Are you sure?", parent=self):
            existing_profiles = self.modified_settings.get("profiles", {})
            self.modified_settings = Settings.DEFAULT_SETTINGS.copy()
            self.modified_settings["profiles"] = existing_profiles
            self.load_current_values()
            messagebox.showinfo("Reset Complete", "Settings reset to defaults.", parent=self)

    def load_current_values(self):
        self.width_var.set(self.modified_settings.get("window_width", 1200))
        self.height_var.set(self.modified_settings.get("window_height", 850))
        self.resolution_var.set(self.modified_settings.get("resolution_threshold", 150))
        self.max_pages_var.set(self.modified_settings.get("max_pages_check_resolution", 5))
        self.min_text_var.set(self.modified_settings.get("min_text_length", 50))
        self.empty_ratio_var.set(self.modified_settings.get("empty_page_threshold", 0.8))
        self.auto_update_var.set(self.modified_settings.get("auto_check_updates", True))
        self.target_subfolders_var.set(self.modified_settings.get("target_subfolders", "open, confidential"))
        self.use_aliases_var.set(self.modified_settings.get("use_folder_aliases", True))
        self.enable_cache_var.set(self.modified_settings.get("enable_scan_cache", True))
        self.max_workers_var.set(self.modified_settings.get("max_workers", 4))
        self.per_file_timeout_var.set(self.modified_settings.get("per_file_timeout_seconds", 30))

    def save(self):
        self.modified_settings["window_width"] = self.width_var.get()
        self.modified_settings["window_height"] = self.height_var.get()
        self.modified_settings["resolution_threshold"] = self.resolution_var.get()
        self.modified_settings["max_pages_check_resolution"] = self.max_pages_var.get()
        self.modified_settings["min_text_length"] = self.min_text_var.get()
        self.modified_settings["empty_page_threshold"] = self.empty_ratio_var.get()
        self.modified_settings["auto_check_updates"] = self.auto_update_var.get()
        self.modified_settings["target_subfolders"] = self.target_subfolders_var.get()
        self.modified_settings["use_folder_aliases"] = self.use_aliases_var.get()
        self.modified_settings["enable_scan_cache"] = self.enable_cache_var.get()
        self.modified_settings["max_workers"] = self.max_workers_var.get()
        self.modified_settings["per_file_timeout_seconds"] = self.per_file_timeout_var.get()

        if self.settings.save(self.modified_settings):
            self.settings.settings = self.modified_settings.copy()
            self.destroy()
        else:
            messagebox.showerror("Error", "Failed to save settings.", parent=self)

    def cancel(self):
        self.destroy()


class PDFErrorChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Error Checker Pro")

        # Set application icon if available
        base_dir = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
        icon_path = base_dir / "icon.ico"
        if not icon_path.exists():
            icon_path = Path(__file__).parent / "icon.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.settings = Settings()
        width = self.settings.get("window_width", 1200)
        height = self.settings.get("window_height", 850)
        # Ensure initial window dimensions do not exceed the physical screen
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        target_w = min(width, max(850, screen_w - 60))
        target_h = min(height, max(500, screen_h - 90))
        self.root.geometry(f"{target_w}x{target_h}")
        self.root.minsize(800, 500)
        self.root.configure(bg="#f5f6fa")
        self.root.resizable(True, True)

        self.folder_path = tk.StringVar(value=self.settings.get("last_folder", ""))
        self.results = []
        self.running = False
        self.has_scanned = False
        self.scan_start_time = None
        self.scan_thread = None
        self.scan_cache = ScanCache()
        self.scan_history = ScanHistory()

        self.export_button = None
        self.scan_button = None
        self.cancel_button = None
        self.settings_button = None
        self.history_button = None

        self.setup_ui()
        self.center_window()
        self.root.after(1000, self.check_updates_async)
        self._setup_shortcuts()

    def _setup_shortcuts(self):
        """Enter = Start Scan, Escape = Cancel Scan, Ctrl+E = Export to Word.
        Bound with bind_all so they work regardless of which widget has
        focus, but guarded to only fire for events belonging to this main
        window — otherwise they'd also fire while the Settings dialog (a
        separate Toplevel) is open and has its own grab."""
        self.root.bind_all("<Return>", self._on_return_key)
        self.root.bind_all("<Escape>", self._on_escape_key)
        self.root.bind_all("<Control-e>", self._on_ctrl_e)
        self.root.bind_all("<Control-E>", self._on_ctrl_e)

    def _event_belongs_to_main_window(self, event):
        try:
            return event.widget.winfo_toplevel() is self.root
        except Exception:
            return False

    def _on_return_key(self, event):
        if not self._event_belongs_to_main_window(event):
            return
        # Don't hijack Enter while the user is typing in a text field (e.g.
        # the results filter box) — only treat it as "Start Scan" when focus
        # isn't in an editable field.
        if isinstance(event.widget, (tk.Entry, tk.Text, tk.Spinbox)):
            return
        if self.scan_button and self.scan_button.enabled:
            self.start_scan()

    def _on_escape_key(self, event):
        if not self._event_belongs_to_main_window(event):
            return
        if self.cancel_button and self.cancel_button.enabled:
            self.cancel_scan()

    def _on_ctrl_e(self, event):
        if not self._event_belongs_to_main_window(event):
            return
        if self.export_button and self.export_button.enabled:
            self.export_to_word()
        return "break"

    def _notify(self, title, message):
        """Non-blocking desktop notification when plyer is available;
        otherwise falls back to the original blocking message box so
        nothing is lost on a machine without it installed."""
        if PLYER_AVAILABLE:
            try:
                desktop_notification.notify(
                    title=title, message=message,
                    app_name="PDF Error Checker Pro", timeout=6,
                )
                return
            except Exception:
                pass
        messagebox.showinfo(title, message)

    def center_window(self):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_reqwidth()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_reqheight()) // 2
        self.root.geometry(f"+{x}+{y}")

    def check_updates_async(self):
        """Check for updates in background. All UI work is marshalled back
        onto the main thread via root.after — Tkinter calls are not
        thread-safe and must never run directly inside the worker thread."""
        if not self.settings.get("auto_check_updates", True):
            return

        def worker():
            checker = UpdateChecker(GITHUB_REPO, CURRENT_VERSION)
            result = checker.check_for_updates()
            self.root.after(0, lambda: self._handle_update_result(result))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _handle_update_result(self, result):
        # Runs on the main thread — safe to touch Tk widgets here.
        has_update, latest, data = result
        if not (has_update and data):
            return

        release_url = data.get('html_url', '')
        release_notes = data.get('body') or 'No release notes available.'
        asset = UpdateChecker.find_installable_asset(data)

        dialog = tk.Toplevel(self.root)
        dialog.title("Update Available")
        dialog.configure(bg="#f5f6fa")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = tk.Frame(dialog, bg="#f5f6fa", padx=25, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text=f"A new version ({latest}) is available.",
                 font=("Segoe UI", 12, "bold"), bg="#f5f6fa", fg="#2c3e50").pack(anchor=tk.W)
        tk.Label(frame, text=f"You have {CURRENT_VERSION}.", font=("Segoe UI", 9),
                 bg="#f5f6fa", fg="#7f8c8d").pack(anchor=tk.W, pady=(0, 10))

        notes_frame = tk.Frame(frame, bg="#ffffff", padx=10, pady=8)
        notes_frame.pack(fill=tk.X, pady=(0, 15))
        tk.Label(notes_frame, text=release_notes[:500], font=("Segoe UI", 9), bg="#ffffff",
                 fg="#34495e", wraplength=400, justify=tk.LEFT).pack(anchor=tk.W)

        if not asset:
            tk.Label(frame, text="No auto-installable update is available for this build — "
                                 "you'll need to download it yourself.",
                     font=("Segoe UI", 8, "italic"), bg="#f5f6fa", fg="#95a5a6",
                     wraplength=400, justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 10))

        btn_row = tk.Frame(frame, bg="#f5f6fa")
        btn_row.pack(fill=tk.X)

        def close_and(fn=None):
            dialog.destroy()
            if fn:
                fn()

        if asset:
            RoundedButton(btn_row, text="Download && Install", bg="#27ae60", fg="white",
                          font=("Segoe UI", 10, "bold"), width=170, height=36,
                          command=lambda: close_and(lambda: self._start_update_download(latest, asset))
                          ).pack(side=tk.RIGHT)
        else:
            RoundedButton(btn_row, text="Open Release Page", bg="#3498db", fg="white",
                          font=("Segoe UI", 10, "bold"), width=170, height=36,
                          command=lambda: close_and(lambda: webbrowser.open(release_url) if release_url else None)
                          ).pack(side=tk.RIGHT)

        RoundedButton(btn_row, text="Later", bg="#95a5a6", fg="white",
                      font=("Segoe UI", 10, "bold"), width=90, height=36,
                      command=lambda: close_and()).pack(side=tk.RIGHT, padx=(0, 10))

        if asset and release_url:
            RoundedButton(btn_row, text="View Release Page", bg="#f5f6fa", fg="#3498db",
                          font=("Segoe UI", 9), width=150, height=36,
                          command=lambda: webbrowser.open(release_url)).pack(side=tk.LEFT)

        dialog.update_idletasks()
        w, h = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
        x = (dialog.winfo_screenwidth() - w) // 2
        y = (dialog.winfo_screenheight() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")

    def _start_update_download(self, latest_version, asset):
        """Downloads the release asset in the background with a progress
        dialog, then hands off to the platform-appropriate installer."""
        download_url = asset.get("browser_download_url")
        total_size = asset.get("size") or 0
        if not download_url:
            messagebox.showerror("Update Failed", "The release asset has no download URL.")
            return

        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("Downloading Update")
        progress_dialog.configure(bg="#f5f6fa")
        progress_dialog.transient(self.root)
        progress_dialog.resizable(False, False)
        progress_dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # no closing mid-download

        frame = tk.Frame(progress_dialog, bg="#f5f6fa", padx=25, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        status_label = tk.Label(frame, text=f"Downloading version {latest_version}...",
                                font=("Segoe UI", 10), bg="#f5f6fa", fg="#2c3e50")
        status_label.pack(anchor=tk.W, pady=(0, 10))
        bar = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=320,
                              mode="determinate" if total_size else "indeterminate")
        bar.pack(fill=tk.X)
        if not total_size:
            bar.start(15)

        progress_dialog.update_idletasks()
        w, h = progress_dialog.winfo_reqwidth(), progress_dialog.winfo_reqheight()
        x = (progress_dialog.winfo_screenwidth() - w) // 2
        y = (progress_dialog.winfo_screenheight() - h) // 2
        progress_dialog.geometry(f"{w}x{h}+{x}+{y}")

        def worker():
            try:
                temp_dir = Path(tempfile.gettempdir())
                dest_path = temp_dir / asset.get("name", f"pdf_checker_update_{latest_version}")
                downloaded = 0
                headers = {"User-Agent": f"PDF-Error-Checker-Pro/{CURRENT_VERSION}"}
                with requests.get(download_url, headers=headers, stream=True, timeout=30) as resp:
                    resp.raise_for_status()
                    with open(dest_path, "wb") as f:
                        for chunk in resp.iter_content(chunk_size=65536):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total_size:
                                    pct = downloaded / total_size
                                    self.root.after(0, lambda p=pct: bar.config(value=p * 100))

                if downloaded == 0:
                    raise IOError("Downloaded file is empty.")

                logger.info(f"Update {latest_version} downloaded to {dest_path} ({downloaded} bytes)")
                self.root.after(0, lambda: self._install_downloaded_update(progress_dialog, dest_path, latest_version))
            except Exception as e:
                logger.error(f"Update download failed: {e}")
                self.root.after(0, lambda: self._update_failed(progress_dialog, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _update_failed(self, progress_dialog, error_text):
        progress_dialog.destroy()
        messagebox.showerror("Update Failed",
                             f"Couldn't download the update automatically:\n{error_text}\n\n"
                             "You can still download it manually from the release page.")

    def _install_downloaded_update(self, progress_dialog, downloaded_path, latest_version):
        progress_dialog.destroy()
        if not messagebox.askyesno(
            "Ready to Install",
            f"Version {latest_version} has been downloaded.\n\n"
            "The app will close and restart automatically to finish installing. "
            "Save any work first.\n\nInstall now?"
        ):
            return

        try:
            if getattr(sys, "frozen", False):
                self._apply_windows_exe_update(downloaded_path)
            else:
                self._apply_source_update(downloaded_path)
        except Exception as e:
            logger.error(f"Failed to apply update: {e}")
            messagebox.showerror("Update Failed", f"Couldn't install the update:\n{e}")

    def _apply_windows_exe_update(self, new_exe_path):
        """Replaces the running .exe with the downloaded one. Windows won't
        let a running executable overwrite itself directly, so this writes
        a tiny helper batch script that: waits for the current process to
        release its file lock, moves the new exe into place, relaunches it,
        then deletes itself. We hand off to that script and exit — this is
        the standard self-update pattern used by many Windows desktop apps."""
        current_exe = Path(sys.executable)
        new_exe_path = Path(new_exe_path)
        batch_path = Path(tempfile.gettempdir()) / "pdf_checker_update.bat"

        batch_script = f"""@echo off
:wait_loop
del "{current_exe}" >NUL 2>&1
if exist "{current_exe}" (
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

        subprocess.Popen(
            ["cmd", "/c", str(batch_path)],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        self.root.destroy()
        sys.exit(0)

    def _apply_source_update(self, new_script_path):
        """Running from source: overwrite this script's file with the
        downloaded one and relaunch. No file-lock workaround is needed
        here — Python reads the script into memory at startup and doesn't
        keep it open, so the file on disk can be safely replaced while
        running."""
        current_script = Path(__file__).resolve()
        new_script_path = Path(new_script_path)

        backup_path = current_script.with_suffix(current_script.suffix + ".bak")
        shutil.copy2(current_script, backup_path)
        shutil.copy2(new_script_path, current_script)
        logger.info(f"Replaced {current_script} with downloaded update (backup at {backup_path})")

        subprocess.Popen([sys.executable, str(current_script)])
        self.root.destroy()
        sys.exit(0)

    def setup_ui(self):
        self.setup_styles()
        self.setup_header()
        self.setup_dependency_banner()
        self.setup_main_content()
        self.setup_status_bar()

    def setup_styles(self):
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.style.configure("TFrame", background="#f5f6fa")
        self.style.configure("TLabel", background="#f5f6fa", font=("Segoe UI", 10))
        self.style.configure("TProgressbar", thickness=20)

    def setup_header(self):
        header_frame = tk.Frame(self.root, bg="#2c3e50", padx=25, pady=15)
        header_frame.pack(fill=tk.X)

        title_frame = tk.Frame(header_frame, bg="#2c3e50")
        title_frame.pack(side=tk.LEFT)

        tk.Label(title_frame, text="PDF Error Checker Pro", font=("Segoe UI", 18, "bold"),
                 fg="white", bg="#2c3e50").pack(side=tk.LEFT)

        self.settings_button = RoundedButton(header_frame, text="Settings", command=self.open_settings,
                                             bg="#34495e", fg="white", font=("Segoe UI", 10),
                                             width=120, height=38)
        self.settings_button.pack(side=tk.RIGHT)

        self.history_button = RoundedButton(header_frame, text="History", command=self.open_history,
                                            bg="#34495e", fg="white", font=("Segoe UI", 10),
                                            width=100, height=38)
        self.history_button.pack(side=tk.RIGHT, padx=(0, 10))

    def open_history(self):
        ScanHistoryDialog(self.root, self.scan_history)

    def setup_dependency_banner(self):
        """Show a visible warning if PyMuPDF is missing, since two of the
        three scan checks degrade significantly without it."""
        if FITZ_AVAILABLE:
            return
        banner = tk.Frame(self.root, bg="#fdecea", padx=15, pady=8)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text=("PyMuPDF is not installed — resolution checks are disabled and the "
                  "missing-information check will use a slower fallback reader. "
                  "Install with: pip install pymupdf"),
            font=("Segoe UI", 9), bg="#fdecea", fg="#c0392b", wraplength=1000, justify=tk.LEFT,
        ).pack(anchor=tk.W)

    def setup_main_content(self):
        main_frame = tk.Frame(self.root, bg="#f5f6fa", padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel: wrapped in a scrollable canvas with a vertical scrollbar
        # so all options and action buttons remain fully visible and reachable
        # on smaller screens or displays with high DPI scaling.
        left_container = tk.Frame(main_frame, bg="#f5f6fa", width=380)
        left_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        left_container.pack_propagate(False)

        left_canvas = tk.Canvas(left_container, bg="#f5f6fa", highlightthickness=0, width=355)
        left_scrollbar = ttk.Scrollbar(left_container, orient=tk.VERTICAL, command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        left_panel = tk.Frame(left_canvas, bg="#f5f6fa")
        left_window_id = left_canvas.create_window((0, 0), window=left_panel, anchor=tk.NW, width=355)

        def _update_left_scrollregion(_event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _resize_left_inner(event):
            left_canvas.itemconfig(left_window_id, width=event.width)
            _update_left_scrollregion()

        left_panel.bind("<Configure>", _update_left_scrollregion)
        left_canvas.bind("<Configure>", _resize_left_inner)

        def _on_left_mousewheel(event):
            # Allow scrolling when content exceeds visible canvas height
            if left_canvas.winfo_height() < left_panel.winfo_reqheight():
                left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel_tree(widget):
            widget.bind("<MouseWheel>", _on_left_mousewheel, add="+")
            for child in widget.winfo_children():
                _bind_mousewheel_tree(child)

        self.setup_folder_selection(left_panel)
        self.setup_scan_mode(left_panel)
        self.setup_scan_options(left_panel)
        self.setup_action_buttons(left_panel)

        _bind_mousewheel_tree(left_panel)
        left_canvas.bind("<MouseWheel>", _on_left_mousewheel)
        left_container.bind("<MouseWheel>", _on_left_mousewheel)
        left_container.bind("<Enter>", lambda _e: self.root.bind_all("<MouseWheel>", _on_left_mousewheel))
        left_container.bind("<Leave>", lambda _e: self.root.unbind_all("<MouseWheel>"))

        right_panel = tk.Frame(main_frame, bg="#f5f6fa")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.setup_progress_section(right_panel)
        self.setup_results_section(right_panel)
        self.setup_summary_section(right_panel)

    def setup_folder_selection(self, parent):
        folder_frame = tk.LabelFrame(parent, text="Folder Selection", font=("Segoe UI", 11, "bold"),
                                     bg="#ffffff", fg="#2c3e50", padx=12, pady=10, relief=tk.FLAT)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(folder_frame, text="Parent folder (will scan all subfolders with 'open' & 'confidential'):",
                font=("Segoe UI", 9), bg="#ffffff", fg="#7f8c8d", wraplength=320).pack(anchor=tk.W, pady=(0, 6))

        folder_entry_frame = tk.Frame(folder_frame, bg="#ffffff")
        folder_entry_frame.pack(fill=tk.X)

        self.folder_entry = tk.Entry(folder_entry_frame, textvariable=self.folder_path, font=("Segoe UI", 9),
                                     bg="#ecf0f1", fg="#2c3e50", relief=tk.FLAT, state="readonly")
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        browse_btn = RoundedButton(folder_entry_frame, text="Browse", command=self.browse_folder,
                                   bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"),
                                   width=100, height=35)
        browse_btn.pack(side=tk.RIGHT, padx=(10, 0))

    def setup_scan_mode(self, parent):
        mode_frame = tk.LabelFrame(parent, text="Scan Mode", font=("Segoe UI", 11, "bold"),
                                   bg="#ffffff", fg="#2c3e50", padx=12, pady=10, relief=tk.FLAT)
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        self.scan_mode_var = tk.StringVar(value=self.settings.get("scan_mode", "project"))

        project_rb = tk.Radiobutton(
            mode_frame, text="Project Structure (target subfolders)", variable=self.scan_mode_var,
            value="project", font=("Segoe UI", 10), bg="#ffffff", fg="#34495e",
            selectcolor="#ffffff", activebackground="#ffffff", command=self._on_scan_mode_change,
        )
        project_rb.pack(anchor=tk.W, pady=(0, 2))
        Tooltip(project_rb, "Only scans subfolders matching the target names configured in "
                            "Settings → Scan Settings (default: open, confidential).")

        all_rb = tk.Radiobutton(
            mode_frame, text="All PDFs (Recursive)", variable=self.scan_mode_var,
            value="all", font=("Segoe UI", 10), bg="#ffffff", fg="#34495e",
            selectcolor="#ffffff", activebackground="#ffffff", command=self._on_scan_mode_change,
        )
        all_rb.pack(anchor=tk.W)
        Tooltip(all_rb, "Ignores subfolder naming entirely and scans every PDF found under the "
                        "selected folder and all of its subfolders.")

        self.target_subfolders_hint = tk.Label(
            mode_frame, text=self._target_subfolders_hint_text(), font=("Segoe UI", 8, "italic"),
            bg="#ffffff", fg="#95a5a6", wraplength=320, justify=tk.LEFT,
        )
        self.target_subfolders_hint.pack(anchor=tk.W, pady=(4, 0))

    def _target_subfolders_hint_text(self):
        targets = ", ".join(self._get_target_subfolder_names())
        alias_note = " (+ common variations)" if self.settings.get("use_folder_aliases", True) else ""
        return f"Project mode targets: {targets}{alias_note}. Edit in Settings → Scan Settings."

    def _on_scan_mode_change(self):
        self.settings.set("scan_mode", self.scan_mode_var.get())

    def _get_target_subfolder_names(self):
        raw = self.settings.get("target_subfolders", "open, confidential")
        names = [n.strip() for n in raw.split(",") if n.strip()]
        return names or ["open", "confidential"]

    @staticmethod
    def _expand_folder_aliases(name_lower):
        for group in FOLDER_NAME_GROUPS:
            if name_lower in group:
                return group
        return {name_lower}

    def _match_subfolders(self, dirs):
        """Given the subdirectory names found at one level of os.walk,
        return (target_name, actual_dir_name) pairs for every configured
        target subfolder that's present — matched case-insensitively, and
        via known naming aliases when enabled. A folder only needs to
        contain at least one of these, not all of them."""
        targets = self._get_target_subfolder_names()
        use_aliases = self.settings.get("use_folder_aliases", True)

        dirs_lower_map = {}
        for d in dirs:
            dirs_lower_map.setdefault(d.lower(), d)

        matches = []
        seen_actual = set()
        for target in targets:
            target_lower = target.lower()
            candidate_names = self._expand_folder_aliases(target_lower) if use_aliases else {target_lower}
            for cand in candidate_names:
                if cand in dirs_lower_map:
                    actual = dirs_lower_map[cand]
                    if actual not in seen_actual:
                        matches.append((target, actual))
                        seen_actual.add(actual)
                    break
        return matches

    def setup_scan_options(self, parent):
        options_frame = tk.LabelFrame(parent, text="Scan Options", font=("Segoe UI", 11, "bold"),
                                      bg="#ffffff", fg="#2c3e50", padx=12, pady=10, relief=tk.FLAT)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        self.check_cannot_open = tk.BooleanVar(value=self.settings.get("check_cannot_open", True))
        self.check_not_clear = tk.BooleanVar(value=self.settings.get("check_not_clear", True))
        self.check_missing_info = tk.BooleanVar(value=self.settings.get("check_missing_info", True))
        self.check_password_protected = tk.BooleanVar(value=self.settings.get("check_password_protected", True))
        self.check_duplicates = tk.BooleanVar(value=self.settings.get("check_duplicates", True))

        checks = [
            ("Cannot Open (Corrupt)", self.check_cannot_open,
             "Flags PDFs that fail to open or report zero pages — likely corrupted files."),
            ("Password Protected", self.check_password_protected,
             "Flags PDFs that require a password to open. Reported separately from 'Cannot "
             "Open' so it's clear which files just need a password versus which are truly broken."),
            ("Not Clear (Low Resolution)", self.check_not_clear,
             f"Flags scanned pages with embedded images below the DPI threshold below."
             f"{'' if FITZ_AVAILABLE else ' (Disabled: requires PyMuPDF.)'}"),
            ("Missing Information", self.check_missing_info,
             "Flags PDFs with mostly blank pages or lacking content."),
            ("Duplicate Detection", self.check_duplicates,
             "Hashes file contents to flag the same PDF appearing in more than one scanned "
             "location (e.g. both 'open' and 'confidential') — its own kind of compliance issue."),
        ]

        self._scan_check_vars = [var for _, var, _ in checks]
        for text, var, tip_text in checks:
            cb = tk.Checkbutton(options_frame, text=text, variable=var, font=("Segoe UI", 10),
                          bg="#ffffff", fg="#34495e", selectcolor="#ffffff",
                          activebackground="#ffffff", activeforeground="#34495e")
            cb.pack(anchor=tk.W, pady=2)
            Tooltip(cb, tip_text)

        toggle_frame = tk.Frame(options_frame, bg="#ffffff")
        toggle_frame.pack(anchor=tk.W, pady=(2, 0))
        tk.Button(toggle_frame, text="Select all", command=lambda: self._set_all_checks(True),
                  font=("Segoe UI", 8), bg="#ffffff", fg="#3498db", relief=tk.FLAT,
                  cursor="hand2", padx=0).pack(side=tk.LEFT)
        tk.Label(toggle_frame, text=" | ", font=("Segoe UI", 8), bg="#ffffff", fg="#bdc3c7").pack(side=tk.LEFT)
        tk.Button(toggle_frame, text="Select none", command=lambda: self._set_all_checks(False),
                  font=("Segoe UI", 8), bg="#ffffff", fg="#3498db", relief=tk.FLAT,
                  cursor="hand2", padx=0).pack(side=tk.LEFT)

        res_frame = tk.Frame(options_frame, bg="#ffffff")
        res_frame.pack(fill=tk.X, pady=(6, 0))

        tk.Label(res_frame, text="Resolution Threshold:", font=("Segoe UI", 10), bg="#ffffff",
                 fg="#34495e").pack(side=tk.LEFT)

        # Single source of truth for the DPI threshold — the Settings dialog
        # edits this same setting rather than keeping a second, separately
        # synced value.
        self.resolution_var = tk.IntVar(value=self.settings.get("resolution_threshold", 150))
        res_spinbox = tk.Spinbox(res_frame, from_=72, to=600, textvariable=self.resolution_var,
                  font=("Segoe UI", 10), width=8, state="readonly")
        res_spinbox.pack(side=tk.LEFT, padx=10)
        Tooltip(res_spinbox, "Pages with embedded images below this DPI are flagged as low resolution.")

        tk.Label(res_frame, text="DPI", font=("Segoe UI", 9), bg="#ffffff", fg="#7f8c8d").pack(side=tk.LEFT)

    def _set_all_checks(self, value):
        for var in self._scan_check_vars:
            var.set(value)

    def setup_action_buttons(self, parent):
        buttons_frame = tk.Frame(parent, bg="#f5f6fa")
        buttons_frame.pack(fill=tk.X, pady=(10, 15))

        self.scan_button = RoundedButton(buttons_frame, text="Start Scan", command=self.start_scan,
                                         bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"),
                                         width=335, height=42)
        self.scan_button.pack(fill=tk.X, pady=(0, 8))
        Tooltip(self.scan_button, "Scan every project subfolder under the selected parent folder. (Enter)")

        self.cancel_button = RoundedButton(buttons_frame, text="Cancel Scan", command=self.cancel_scan,
                                           bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"),
                                           width=335, height=38, state=tk.DISABLED)
        self.cancel_button.pack(fill=tk.X, pady=(0, 8))
        Tooltip(self.cancel_button, "Stop the scan as soon as possible, mid-file if needed. (Esc)")

        self.export_button = RoundedButton(buttons_frame, text="Export to Word", command=self.export_to_word,
                                           bg="#2980b9", fg="white", font=("Segoe UI", 10, "bold"),
                                           width=335, height=38, state=tk.DISABLED)
        self.export_button.pack(fill=tk.X)
        Tooltip(self.export_button, "Save the current results as a .docx report. (Ctrl+E)")

        clear_btn = RoundedButton(buttons_frame, text="Clear Results", command=self.clear_results,
                                  bg="#95a5a6", fg="white", font=("Segoe UI", 10),
                                  width=335, height=36)
        clear_btn.pack(fill=tk.X, pady=(8, 0))
        Tooltip(clear_btn, "Clear the results list and reset progress (does not affect saved reports).")

    def setup_progress_section(self, parent):
        progress_frame = tk.LabelFrame(parent, text="Progress", font=("Segoe UI", 11, "bold"),
                                       bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        progress_frame.pack(fill=tk.X, pady=(0, 15))

        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 10))

        status_inner = tk.Frame(progress_frame, bg="#ffffff")
        status_inner.pack(fill=tk.X)

        self.status_label = tk.Label(status_inner, text="Ready to scan", font=("Segoe UI", 10),
                                     bg="#ffffff", fg="#27ae60")
        self.status_label.pack(side=tk.LEFT)

        self.current_file_label = tk.Label(status_inner, text="", font=("Segoe UI", 9),
                                          bg="#ffffff", fg="#7f8c8d", wraplength=500)
        self.current_file_label.pack(side=tk.RIGHT)

        self.eta_label = tk.Label(progress_frame, text="", font=("Segoe UI", 9),
                                   bg="#ffffff", fg="#7f8c8d")
        self.eta_label.pack(fill=tk.X, pady=(6, 0))

    def setup_results_section(self, parent):
        results_frame = tk.LabelFrame(parent, text="Scan Results", font=("Segoe UI", 11, "bold"),
                                      bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        results_frame.pack(fill=tk.BOTH, expand=True)

        # Filter bar
        filter_frame = tk.Frame(results_frame, bg="#ffffff")
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(filter_frame, text="Filter:", font=("Segoe UI", 9), bg="#ffffff",
                 fg="#34495e").pack(side=tk.LEFT)

        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *args: self._refresh_results_tree())
        filter_entry = tk.Entry(filter_frame, textvariable=self.filter_var, font=("Segoe UI", 9),
                                bg="#ecf0f1", fg="#2c3e50", relief=tk.FLAT)
        filter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        Tooltip(filter_entry, "Filter by filename, parent folder, subfolder, or error type.")

        clear_filter_btn = tk.Button(filter_frame, text="Clear", command=lambda: self.filter_var.set(""),
                                     font=("Segoe UI", 8), bg="#ffffff", fg="#3498db", relief=tk.FLAT,
                                     cursor="hand2")
        clear_filter_btn.pack(side=tk.LEFT)

        # Sortable results table. Columns are click-to-sort; rows are
        # color-tagged by the most severe error they carry so problems are
        # scannable at a glance, and double-click / right-click open the
        # underlying file or reveal it in the file manager.
        tree_container = tk.Frame(results_frame, bg="#ffffff")
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("filename", "parent", "folder", "errors")
        self.results_tree = ttk.Treeview(tree_container, columns=columns, show="headings",
                                         selectmode="browse")
        headings = {"filename": "File", "parent": "Parent Folder", "folder": "Subfolder", "errors": "Errors"}
        widths = {"filename": 220, "parent": 140, "folder": 90, "errors": 200}
        for col in columns:
            self.results_tree.heading(col, text=headings[col],
                                      command=lambda c=col: self._sort_results_by(c))
            self.results_tree.column(col, width=widths[col], anchor=tk.W, stretch=True)

        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=vsb.set)
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Row tags: severity-based background tint (most severe error wins).
        self.results_tree.tag_configure("cannot_open", background="#fdecea")
        self.results_tree.tag_configure("password_protected", background="#f3e6fb")
        self.results_tree.tag_configure("duplicate", background="#e6f0fa")
        self.results_tree.tag_configure("not_clear", background="#fff6e5")
        self.results_tree.tag_configure("missing_info", background="#fffbe0")
        self.results_tree.tag_configure("placeholder", foreground="#95a5a6")

        self.results_tree.bind("<Double-1>", self._on_result_double_click)
        self.results_tree.bind("<Button-3>", self._on_result_right_click)   # Windows/Linux right-click
        self.results_tree.bind("<Button-2>", self._on_result_right_click)   # macOS right-click

        self._sort_state = {"column": None, "reverse": False}

        self.results_context_menu = tk.Menu(self.results_tree, tearoff=0)
        self.results_context_menu.add_command(label="Open File", command=self._open_selected_result)
        self.results_context_menu.add_command(label="Show in Folder", command=self._reveal_selected_result)
        self.results_context_menu.add_separator()
        self.results_context_menu.add_command(label="Copy Path", command=self._copy_selected_result_path)

        # Show the friendly empty-state message immediately — the panel
        # otherwise starts as just a blank box, which reads as broken on
        # first run rather than "nothing to show yet".
        self._refresh_results_tree()

    def _tag_for_errors(self, errors):
        """Most severe error determines the row's color tint."""
        if "Cannot Open" in errors:
            return "cannot_open"
        if "Password Protected" in errors:
            return "password_protected"
        if any(e.startswith("Duplicate") for e in errors):
            return "duplicate"
        if "Not Clear" in errors:
            return "not_clear"
        if "Missing Information" in errors:
            return "missing_info"
        return ""

    def _row_matches_filter(self, result, filter_text):
        if not filter_text:
            return True
        haystack = " ".join([
            result.get("filename", ""), result.get("parent", ""),
            result.get("folder", ""), ", ".join(result.get("errors", [])),
        ]).lower()
        return filter_text.lower() in haystack

    def _insert_placeholder_row(self, message):
        self.results_tree.insert("", tk.END, iid="__placeholder__",
                                 values=(message, "", "", ""), tags=("placeholder",))

    def _refresh_results_tree(self):
        """Full rebuild of the visible tree from self.results, honoring the
        current filter text and sort column. Each row's iid is the row's
        index into self.results, so double-click/right-click and re-sorting
        can always map back to the underlying result."""
        if not hasattr(self, "results_tree"):
            return
        self.results_tree.delete(*self.results_tree.get_children())

        if not self.results:
            message = ("No errors found — every scanned PDF looks good."
                       if self.has_scanned else
                       "No results yet — select a folder and start a scan.")
            self._insert_placeholder_row(message)
            return

        filter_text = self.filter_var.get().strip() if hasattr(self, "filter_var") else ""
        visible = [(i, r) for i, r in enumerate(self.results) if self._row_matches_filter(r, filter_text)]

        if not visible:
            self._insert_placeholder_row("No results match your filter.")
            return

        sort_col = self._sort_state.get("column")
        if sort_col:
            visible.sort(key=lambda item: self._sort_key(item[1], sort_col),
                        reverse=self._sort_state.get("reverse", False))

        for index, result in visible:
            values = (result["filename"], result.get("parent", "N/A"), result["folder"],
                     ", ".join(result["errors"]))
            self.results_tree.insert("", tk.END, iid=str(index), values=values,
                                     tags=(self._tag_for_errors(result["errors"]),))

    def _sort_key(self, result, column):
        if column == "filename":
            return result.get("filename", "").lower()
        if column == "parent":
            return result.get("parent", "").lower()
        if column == "folder":
            return result.get("folder", "").lower()
        if column == "errors":
            return ", ".join(result.get("errors", [])).lower()
        return ""

    def _sort_results_by(self, column):
        if self._sort_state.get("column") == column:
            self._sort_state["reverse"] = not self._sort_state["reverse"]
        else:
            self._sort_state["column"] = column
            self._sort_state["reverse"] = False
        self._refresh_results_tree()

    def _append_result_row(self, result, index):
        """Add a single new row while a scan is running, respecting the
        current filter. Sort order is only reapplied on the next full
        refresh (filter change, header click, or scan completion) rather
        than on every single insert, which would be wasteful during a big
        scan."""
        if self.results_tree.exists("__placeholder__"):
            self.results_tree.delete("__placeholder__")
        filter_text = self.filter_var.get().strip() if hasattr(self, "filter_var") else ""
        if not self._row_matches_filter(result, filter_text):
            return
        values = (result["filename"], result.get("parent", "N/A"), result["folder"],
                 ", ".join(result["errors"]))
        self.results_tree.insert("", tk.END, iid=str(index), values=values,
                                 tags=(self._tag_for_errors(result["errors"]),))

    def _get_result_for_iid(self, iid):
        try:
            return self.results[int(iid)]
        except (ValueError, IndexError):
            return None

    def _on_result_double_click(self, _event=None):
        self._open_selected_result()

    def _on_result_right_click(self, event):
        row_iid = self.results_tree.identify_row(event.y)
        if not row_iid:
            return
        self.results_tree.selection_set(row_iid)
        self.results_context_menu.tk_popup(event.x_root, event.y_root)

    def _selected_result(self):
        selection = self.results_tree.selection()
        if not selection:
            return None
        return self._get_result_for_iid(selection[0])

    def _open_selected_result(self):
        result = self._selected_result()
        if not result:
            return
        self._open_file(result["path"])

    def _reveal_selected_result(self):
        result = self._selected_result()
        if not result:
            return
        self._reveal_in_file_manager(result["path"])

    def _copy_selected_result_path(self):
        result = self._selected_result()
        if not result:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(result["path"])

    def _open_file(self, path):
        if not os.path.exists(path):
            messagebox.showwarning("File Not Found", f"This file no longer exists:\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    def _reveal_in_file_manager(self, path):
        if not os.path.exists(path):
            messagebox.showwarning("File Not Found", f"This file no longer exists:\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.run(["explorer", f'/select,"{os.path.normpath(path)}"'], check=False)
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", path], check=False)
            else:
                # Most Linux file managers have no "select this file" verb
                # over the command line; opening the containing folder is
                # the closest reliable equivalent.
                subprocess.run(["xdg-open", os.path.dirname(path)], check=False)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open containing folder:\n{e}")

    def setup_summary_section(self, parent):
        summary_frame = tk.Frame(parent, bg="#ffffff", padx=15, pady=15)
        summary_frame.pack(fill=tk.X, pady=(15, 0))

        self.summary_label = tk.Label(summary_frame, text="No scan performed yet.",
                                      font=("Segoe UI", 10), bg="#ffffff", fg="#7f8c8d")
        self.summary_label.pack(side=tk.LEFT)

    def setup_status_bar(self):
        status_bar = tk.Frame(self.root, bg="#34495e", padx=15, pady=8)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        tk.Label(status_bar, text=f"PDF Error Checker Pro v{CURRENT_VERSION}", font=("Segoe UI", 9),
                bg="#34495e", fg="#ecf0f1").pack(side=tk.LEFT)

        deps = []
        if FITZ_AVAILABLE:
            deps.append("PyMuPDF OK")
        if PYPDF2_AVAILABLE:
            deps.append("PyPDF2 OK")

        if deps:
            tk.Label(status_bar, text=" | ".join(deps), font=("Segoe UI", 9),
                    bg="#34495e", fg="#2ecc71").pack(side=tk.RIGHT)

    def open_settings(self):
        dialog = SettingsDialog(self.root, self.settings)
        dialog.transient(self.root)
        dialog.wait_window()
        # Re-sync the single shared DPI control, and apply the new window
        # size immediately instead of requiring a restart.
        self.resolution_var.set(self.settings.get("resolution_threshold", 150))
        width = self.settings.get("window_width", 1200)
        height = self.settings.get("window_height", 850)
        self.root.geometry(f"{width}x{height}")
        if hasattr(self, "target_subfolders_hint"):
            self.target_subfolders_hint.config(text=self._target_subfolders_hint_text())
        # Reload from disk: picks up a "Clear Cache" click, and ensures any
        # changed thresholds are reflected the next time the cache is
        # consulted (a stale in-memory copy could otherwise mask a clear).
        self.scan_cache = ScanCache()

    def browse_folder(self):
        folder_selected = filedialog.askdirectory(initialdir=self.settings.get("last_folder", ""))
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.settings.set("last_folder", folder_selected)

    def clear_results(self):
        if hasattr(self, "filter_var"):
            self.filter_var.set("")
        self._sort_state = {"column": None, "reverse": False}
        self.results = []
        self.has_scanned = False
        self._refresh_results_tree()
        self.summary_label.config(text="No scan performed yet.")
        self.progress["value"] = 0
        self.status_label.config(text="Ready to scan", fg="#27ae60")
        self.current_file_label.config(text="")
        self.eta_label.config(text="")
        if self.export_button:
            self.export_button.config(state=tk.DISABLED)
        if self.scan_button:
            self.scan_button.config(state=tk.NORMAL)
        if self.cancel_button:
            self.cancel_button.config(state=tk.DISABLED)

    def export_to_word(self):
        if not self.results:
            messagebox.showwarning("No Results", "No scan results to export!")
            return

        if not DOCX_AVAILABLE:
            messagebox.showerror("Missing Dependency", "python-docx is required. Install with: pip install python-docx")
            return

        doc_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
                                                title="Save Word Report",
                                                initialdir=self.settings.get("last_folder", ""))

        if not doc_path:
            return

        try:
            doc = Document()
            title = doc.add_heading("PDF Error Checker Report", level=1)
            title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            doc.add_paragraph(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            doc.add_paragraph(f"Scanned Parent Folder: {self.folder_path.get()}")

            doc.add_heading("Scan Settings", level=2)
            settings_para = doc.add_paragraph()
            settings_para.add_run("Error Types: ").bold = True
            settings = []
            if self.check_cannot_open.get():
                settings.append("Cannot Open")
            if self.check_password_protected.get():
                settings.append("Password Protected")
            if self.check_not_clear.get():
                settings.append("Not Clear")
            if self.check_missing_info.get():
                settings.append("Missing Information")
            if self.check_duplicates.get():
                settings.append("Duplicate Detection")
            settings_para.add_run(f"{', '.join(settings) if settings else 'None'}")

            doc.add_paragraph(f"Resolution Threshold: {self.resolution_var.get()} DPI")

            doc.add_heading("Summary", level=2)
            total_checked = len(self.results)
            cannot_open_count = sum(1 for r in self.results if "Cannot Open" in r["errors"])
            password_count = sum(1 for r in self.results if "Password Protected" in r["errors"])
            not_clear_count = sum(1 for r in self.results if "Not Clear" in r["errors"])
            missing_info_count = sum(1 for r in self.results if "Missing Information" in r["errors"])
            duplicate_count = sum(1 for r in self.results if any(e.startswith("Duplicate") for e in r["errors"]))
            unique_parents = set(r.get("parent", "N/A") for r in self.results)

            summary_rows = [
                ["Total PDFs With Errors", str(total_checked)],
                ["Parent Folders Affected", str(len(unique_parents))],
                ["Cannot Open", str(cannot_open_count)],
                ["Password Protected", str(password_count)],
                ["Not Clear (Low Resolution)", str(not_clear_count)],
                ["Missing Information", str(missing_info_count)],
                ["Duplicate PDFs", str(duplicate_count)],
            ]

            # Fixed: previously created a 3x2 table for 6 rows of data, which
            # silently discarded the second half. Size the table to the data.
            summary_table = doc.add_table(rows=len(summary_rows), cols=2)
            summary_table.style = "Table Grid"
            for row_idx, (label, value) in enumerate(summary_rows):
                summary_table.cell(row_idx, 0).text = label
                summary_table.cell(row_idx, 1).text = value
                summary_table.cell(row_idx, 0).paragraphs[0].runs[0].bold = True

            doc.add_heading("Detailed Results", level=2)
            for result in self.results:
                doc.add_heading(result['filename'], level=3)
                file_info = doc.add_paragraph()
                file_info.add_run("Path: ").bold = True
                file_info.add_run(f"{result['path']}\n")
                file_info.add_run("Parent: ").bold = True
                file_info.add_run(f"{result.get('parent', 'N/A')}\n")
                file_info.add_run("Subfolder: ").bold = True
                file_info.add_run(f"{result['folder']}\n")
                file_info.add_run("Errors: ").bold = True
                file_info.add_run(f"{', '.join(result['errors'])}")

            doc.save(doc_path)
            logger.info(f"Word report saved: {doc_path}")

            auto_open_pref = self.settings.get("auto_open_word_report", "ask")
            if auto_open_pref == "ask":
                if self._prompt_open_report():
                    self._open_file(doc_path)
            else:
                messagebox.showinfo("Success", f"Report saved to:\n{doc_path}")
                if auto_open_pref == "always":
                    self._open_file(doc_path)

        except Exception as e:
            logger.error(f"Failed to create Word document: {e}")
            messagebox.showerror("Error", f"Failed to create Word document:\n{str(e)}")

    def _prompt_open_report(self):
        """Small custom dialog (a plain messagebox can't host a checkbox)
        offering to open the just-saved report, with a 'don't ask again'
        option that locks in the choice via Settings → General."""
        result = {"open": False}
        dialog = tk.Toplevel(self.root)
        dialog.title("Report Saved")
        dialog.configure(bg="#f5f6fa")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = tk.Frame(dialog, bg="#f5f6fa", padx=25, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Report saved successfully.\nWould you like to open it now?",
                 font=("Segoe UI", 10), bg="#f5f6fa", fg="#2c3e50", justify=tk.LEFT).pack(anchor=tk.W)

        dont_ask_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frame, text="Don't ask again", variable=dont_ask_var, font=("Segoe UI", 9),
                      bg="#f5f6fa", selectcolor="#f5f6fa").pack(anchor=tk.W, pady=(10, 15))

        def choose(should_open):
            if dont_ask_var.get():
                self.settings.set("auto_open_word_report", "always" if should_open else "never")
                self.settings.save(self.settings.settings)
            result["open"] = should_open
            dialog.destroy()

        btn_row = tk.Frame(frame, bg="#f5f6fa")
        btn_row.pack(fill=tk.X)
        RoundedButton(btn_row, text="Open Report", command=lambda: choose(True),
                      bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"),
                      width=130, height=36).pack(side=tk.RIGHT)
        RoundedButton(btn_row, text="Not Now", command=lambda: choose(False),
                      bg="#95a5a6", fg="white", font=("Segoe UI", 10, "bold"),
                      width=100, height=36).pack(side=tk.RIGHT, padx=(0, 10))

        dialog.update_idletasks()
        w, h = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
        x = (dialog.winfo_screenwidth() - w) // 2
        y = (dialog.winfo_screenheight() - h) // 2
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        dialog.wait_window()
        return result["open"]

    def start_scan(self):
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a parent folder first!")
            return

        if not os.path.exists(folder_path):
            messagebox.showerror("Error", "Selected folder does not exist!")
            return

        scan_mode = self.scan_mode_var.get() if hasattr(self, "scan_mode_var") else "project"
        folder_matches = {}

        if scan_mode == "all":
            valid_folders = [folder_path]
            response = messagebox.askyesno(
                "Confirm Scan",
                f"This will recursively scan every PDF found under:\n{folder_path}\n\nContinue?"
            )
        else:
            valid_folders = []
            for root, dirs, files in os.walk(folder_path):
                matches = self._match_subfolders(dirs)
                if matches:
                    valid_folders.append(root)
                    folder_matches[root] = matches

            if not valid_folders:
                targets = ", ".join(self._get_target_subfolder_names())
                messagebox.showerror(
                    "Error",
                    f"No subfolders matching your configured targets ({targets}) were found!\n\n"
                    "Tip: adjust the target subfolder names under Settings → Scan Settings, "
                    "or switch to 'All PDFs (Recursive)' mode to scan every PDF under the "
                    "selected folder regardless of naming."
                )
                return

            response = messagebox.askyesno(
                "Confirm Scan", f"Found {len(valid_folders)} valid project folders.\n\nContinue?"
            )

        if not response:
            return

        self.running = True
        self.clear_results()
        self.has_scanned = True
        self.results = []

        if self.scan_button:
            self.scan_button.config(state=tk.DISABLED)
        if self.cancel_button:
            self.cancel_button.config(state=tk.NORMAL)
        if self.export_button:
            self.export_button.config(state=tk.DISABLED)

        self.scan_thread = threading.Thread(
            target=self.run_scan, args=(valid_folders, scan_mode, folder_matches), daemon=True
        )
        self.scan_thread.start()

    def run_scan(self, valid_folders, scan_mode="project", folder_matches=None):
        check_cannot_open = self.check_cannot_open.get()
        check_not_clear = self.check_not_clear.get()
        check_missing_info = self.check_missing_info.get()
        check_password_protected = self.check_password_protected.get()
        check_duplicates = self.check_duplicates.get()
        resolution_threshold = self.resolution_var.get()
        folder_matches = folder_matches or {}

        cache_enabled = self.settings.get("enable_scan_cache", True)
        max_workers = max(1, int(self.settings.get("max_workers", 4)))
        per_file_timeout = self.settings.get("per_file_timeout_seconds", 30)

        # Everything that can change a check's outcome. If any of these
        # differ from what was in effect when a file was last cached, that
        # cache entry is treated as a miss and the file is re-checked —
        # otherwise a stale hit could silently report an outdated result.
        SCAN_ENGINE_VERSION = 3  # Bumped to invalidate stale 'Not Clear' false-positive cache entries
        check_signature = [
            SCAN_ENGINE_VERSION,
            check_cannot_open, check_not_clear, check_missing_info, check_password_protected,
            resolution_threshold, self.settings.get("max_pages_check_resolution", 5),
            self.settings.get("min_text_length", 50), self.settings.get("empty_page_threshold", 0.8),
        ]

        all_pdfs = []
        if scan_mode == "all":
            # All PDFs (Recursive): no subfolder naming requirement at all —
            # every PDF under the selected folder counts, and the "Subfolder"
            # column shows the file's path relative to that root instead of
            # a fixed "open"/"confidential" label.
            base_folder = valid_folders[0]
            base_name = os.path.basename(os.path.normpath(base_folder)) or base_folder
            for root, _, files in os.walk(base_folder):
                rel_dir = os.path.relpath(root, base_folder)
                subfolder_label = "(root)" if rel_dir == "." else rel_dir.replace("\\", "/")
                for file in files:
                    if file.lower().endswith(".pdf"):
                        all_pdfs.append({
                            "path": os.path.join(root, file),
                            "folder": subfolder_label,
                            "parent": base_name,
                            "filename": file,
                        })
        else:
            for folder in valid_folders:
                parent_name = os.path.basename(folder)
                matched_subfolders = [actual for _, actual in folder_matches.get(folder, [])]
                for subfolder in matched_subfolders:
                    subfolder_path = os.path.join(folder, subfolder)
                    if not os.path.exists(subfolder_path):
                        continue
                    for root, _, files in os.walk(subfolder_path):
                        for file in files:
                            if file.lower().endswith(".pdf"):
                                all_pdfs.append({
                                    "path": os.path.join(root, file),
                                    "folder": subfolder,
                                    "parent": parent_name,
                                    "filename": file,
                                })

        if not all_pdfs:
            self.root.after(0, lambda: messagebox.showinfo("Info", "No PDF files found!"))
            self.running = False
            self.root.after(0, self.scan_complete)
            return

        total_pdfs = len(all_pdfs)
        self.scan_start_time = datetime.now()
        logger.info(f"Scan started: mode={scan_mode}, files={total_pdfs}, workers={max_workers}")

        # Fixed: progress bar defaults to maximum=100, so folders with more
        # than 100 PDFs would hit 100% long before the scan actually finished.
        self.root.after(0, lambda: self.progress.config(maximum=total_pdfs, value=0))

        cache_hits = 0
        timeouts = 0

        # File checks are I/O-bound (mostly disk reads + PDF parsing), so a
        # small worker pool gives a real speedup even on a single core.
        # Futures are consumed in the SAME order they were submitted (not
        # via as_completed) specifically so that calling future.result()
        # with a per-file timeout enforces a genuine per-file cap: by the
        # time we ask for a given future's result, it's already been
        # running in the background since submission, so the timeout here
        # is a true ceiling on how long we'll wait for that one file rather
        # than an arbitrary slice of an unrelated overall budget.
        executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="pdfcheck")
        futures = []
        try:
            for pdf_info in all_pdfs:
                if not self.running:
                    break
                future = executor.submit(
                    self._check_pdf_cached, pdf_info["path"], check_cannot_open, check_not_clear,
                    check_missing_info, check_password_protected, resolution_threshold,
                    cache_enabled, check_signature, check_duplicates,
                )
                futures.append((future, pdf_info))

            for i, (future, pdf_info) in enumerate(futures):
                if not self.running:
                    break

                pdf_path = pdf_info["path"]
                folder_name = pdf_info["folder"]
                parent_name = pdf_info["parent"]
                filename = pdf_info["filename"]

                elapsed_seconds = (datetime.now() - self.scan_start_time).total_seconds()
                avg_per_file = elapsed_seconds / max(i, 1) if i > 0 else 0
                remaining_files = total_pdfs - i
                eta_seconds = avg_per_file * remaining_files
                elapsed_str = self._format_duration(elapsed_seconds)
                eta_str = self._format_duration(eta_seconds) if i > 0 else "calculating..."

                self.root.after(0, lambda idx=i+1, total=total_pdfs, name=filename, parent=parent_name,
                                elapsed=elapsed_str, eta=eta_str: (
                    self.status_label.config(text=f"Scanning... ({idx}/{total})"),
                    self.current_file_label.config(text=f"{parent}/{name[:30]}..."),
                    self.progress.config(value=idx),
                    self.eta_label.config(text=f"Elapsed: {elapsed}  |  ETA: {eta}")
                ))

                try:
                    errors, file_hash, used_cache = future.result(timeout=per_file_timeout)
                except FuturesTimeoutError:
                    errors, file_hash, used_cache = ["Scan Timeout"], None, False
                    timeouts += 1
                    logger.warning(f"Per-file timeout ({per_file_timeout}s) exceeded: {pdf_path}")
                except Exception as e:
                    errors, file_hash, used_cache = [], None, False
                    logger.error(f"Error checking {pdf_path}: {e}")

                pdf_info["_hash"] = file_hash
                if used_cache:
                    cache_hits += 1

                if errors and self.running:
                    self.results.append({
                        "path": pdf_path,
                        "folder": folder_name,
                        "parent": parent_name,
                        "filename": filename,
                        "errors": errors,
                    })
                    result_index = len(self.results) - 1
                    result_copy = self.results[result_index]
                    self.root.after(0, lambda r=result_copy, idx=result_index: self._append_result_row(r, idx))
        finally:
            # Don't block on any still-running (e.g. timed-out or hung)
            # worker threads — Python threads can't be force-killed, so we
            # simply stop waiting on them. Anything not yet started is
            # cancelled outright. (A file that's genuinely hung will keep
            # its one worker thread alive in the background; this is a
            # documented, accepted trade-off for a pure-Python thread pool.)
            for future, _ in futures:
                future.cancel()
            executor.shutdown(wait=False)

        if check_duplicates and self.running:
            self.root.after(0, lambda: self.status_label.config(text="Checking for duplicates..."))
            self._detect_duplicates(all_pdfs)

        try:
            self.scan_cache.prune_missing({p["path"] for p in all_pdfs})
            self.scan_cache.save()
        except Exception as e:
            logger.warning(f"Failed to persist scan cache: {e}")

        self.root.after(0, lambda: self.finish_scan(all_pdfs, cache_hits, timeouts))

    def _hash_file(self, path, chunk_size=1024 * 1024):
        try:
            hasher = hashlib.sha256()
            with open(path, "rb") as f:
                while True:
                    if not self.running:
                        return None
                    chunk = f.read(chunk_size)
                    if not chunk:
                        break
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception as e:
            logger.warning(f"Could not hash file {path}: {e}")
            return None

    def _check_pdf_cached(self, pdf_path, check_cannot_open, check_not_clear, check_missing_info,
                          check_password_protected, resolution_threshold, cache_enabled,
                          signature, need_hash=False):
        """Runs on a worker thread. Consults the scan cache first (size +
        mtime + settings signature); on a miss, runs the real checks and
        records the result for next time."""
        size, mtime = None, None
        try:
            stat = os.stat(pdf_path)
            size, mtime = stat.st_size, stat.st_mtime
        except OSError as e:
            logger.warning(f"Could not stat {pdf_path}: {e}")

        if cache_enabled and size is not None:
            cached = self.scan_cache.get(pdf_path, size, mtime, signature)
            if cached is not None:
                file_hash = cached.get("hash")
                if need_hash and file_hash is None:
                    file_hash = self._hash_file(pdf_path)
                return cached.get("errors", []), file_hash, True

        errors = self.check_pdf(pdf_path, check_cannot_open, check_not_clear,
                                check_missing_info, check_password_protected, resolution_threshold)
        file_hash = self._hash_file(pdf_path) if need_hash else None

        if cache_enabled and size is not None and self.running:
            self.scan_cache.set(pdf_path, size, mtime, signature, errors, file_hash)

        return errors, file_hash, False

    def _detect_duplicates(self, all_pdfs):
        """Groups files by content hash (computed during the main pass, or
        here if that step was skipped/cached without one) and flags any
        file whose exact content also appears elsewhere in this scan —
        e.g. the same PDF present in both 'open' and 'confidential'."""
        hash_map = {}
        for pdf_info in all_pdfs:
            if not self.running:
                return
            file_hash = pdf_info.get("_hash")
            if file_hash is None:
                file_hash = self._hash_file(pdf_info["path"])
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
                self._record_duplicate(pdf_info, f"Duplicate (also in {other_desc})")

    def _record_duplicate(self, pdf_info, error_label):
        existing = next((r for r in self.results if r["path"] == pdf_info["path"]), None)
        if existing:
            if error_label not in existing["errors"]:
                existing["errors"].append(error_label)
            return
        self.results.append({
            "path": pdf_info["path"],
            "folder": pdf_info["folder"],
            "parent": pdf_info["parent"],
            "filename": pdf_info["filename"],
            "errors": [error_label],
        })

    @staticmethod
    def _format_duration(total_seconds):
        total_seconds = max(0, int(total_seconds))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"{minutes}:{seconds:02d}"

    def finish_scan(self, all_pdfs, cache_hits=0, timeouts=0):
        total_checked = len(all_pdfs)
        total_errors = len(self.results)
        scan_duration = datetime.now() - self.scan_start_time
        scan_duration_seconds = scan_duration.total_seconds()
        unique_parents = len(set(r.get("parent", "N/A") for r in self.results))

        status_word = "cancelled" if not self.running else "completed"

        summary_text = (f"Scan {status_word}: {total_checked} PDFs in {unique_parents} folders | "
                        f"Errors: {total_errors} | Time: {scan_duration_seconds:.1f}s")
        if cache_hits:
            summary_text += f" | Skipped {cache_hits} unchanged"
        if timeouts:
            summary_text += f" | {timeouts} timed out"

        self.summary_label.config(text=summary_text, fg="#27ae60" if self.running else "#e67e22")
        self.status_label.config(text=f"Scan {status_word}", fg="#27ae60" if self.running else "#e67e22")
        self.current_file_label.config(text="")
        self.eta_label.config(text=f"Total time: {self._format_duration(scan_duration_seconds)}")
        was_cancelled = not self.running
        self.running = False
        logger.info(f"Scan {status_word}: {total_checked} files, {total_errors} errors, "
                   f"{cache_hits} cache hits, {timeouts} timeouts, {scan_duration_seconds:.1f}s")

        self.scan_history.add({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "folder": self.folder_path.get(),
            "scan_mode": self.scan_mode_var.get() if hasattr(self, "scan_mode_var") else "project",
            "files_checked": total_checked,
            "errors_found": total_errors,
            "duration_seconds": round(scan_duration_seconds, 1),
            "cancelled": was_cancelled,
        })

        # Reapply the current sort (if any) now that the full result set is
        # in — during the scan itself, rows were appended in discovery order
        # to avoid re-sorting on every single new result.
        self._refresh_results_tree()

        if self.scan_button:
            self.scan_button.config(state=tk.NORMAL)
        if self.cancel_button:
            self.cancel_button.config(state=tk.DISABLED)
        if self.export_button:
            self.export_button.config(state=tk.NORMAL if self.results else tk.DISABLED)

        if was_cancelled:
            self._notify("Scan Cancelled", f"Stopped after checking {total_checked} PDFs — {total_errors} error(s) found so far.")
        elif total_errors > 0:
            self._notify("Scan Complete", f"Found {total_errors} PDFs with errors in {scan_duration_seconds:.1f}s")
        else:
            self._notify("Scan Complete", f"No errors found in {total_checked} PDFs!")

    def cancel_scan(self):
        self.running = False
        self.status_label.config(text="Cancelling...", fg="#e74c3c")
        if self.cancel_button:
            self.cancel_button.config(state=tk.DISABLED)

    def scan_complete(self):
        self.running = False
        if self.scan_button:
            self.scan_button.config(state=tk.NORMAL)
        if self.cancel_button:
            self.cancel_button.config(state=tk.DISABLED)

    def check_pdf(self, pdf_path, check_cannot_open, check_not_clear, check_missing_info,
                  check_password_protected, resolution_threshold):
        # Finer-grained cancel: previously Cancel only took effect between
        # files, so one huge multi-page PDF could stall it noticeably.
        # Checking self.running between (and inside) each sub-check lets a
        # cancel land mid-file instead of only at file boundaries.
        errors = []
        if not self.running:
            return errors

        is_protected = False
        if check_password_protected:
            is_protected = self.is_pdf_password_protected(pdf_path)
            if is_protected:
                errors.append("Password Protected")

        if not self.running:
            return errors

        # A password-protected file can't be meaningfully read further
        # without the password, so once it's flagged that way we don't also
        # pile on "Cannot Open" (or the other content checks) for what's
        # really the same underlying reason — that's a separate, more
        # actionable finding on its own.
        if is_protected:
            return errors

        if check_cannot_open and self.is_pdf_corrupt(pdf_path):
            errors.append("Cannot Open")
            return errors
        if not self.running:
            return errors
        if check_not_clear and self.is_pdf_not_clear(pdf_path, resolution_threshold):
            errors.append("Not Clear")
        if not self.running:
            return errors
        if check_missing_info and self.has_missing_information(pdf_path):
            errors.append("Missing Information")
        return errors

    def is_pdf_password_protected(self, pdf_path):
        """Detects encryption without needing the actual password — both
        PyMuPDF and PyPDF2 expose this without decrypting the content."""
        if FITZ_AVAILABLE:
            try:
                doc = fitz.open(pdf_path)
                protected = doc.needs_pass
                doc.close()
                return protected
            except Exception:
                pass
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    reader = PdfReader(f)
                    return bool(reader.is_encrypted)
            except Exception:
                pass
        return False

    def is_pdf_corrupt(self, pdf_path):
        if FITZ_AVAILABLE:
            try:
                doc = fitz.open(pdf_path)
                if len(doc) == 0:
                    doc.close()
                    return True
                doc.close()
            except Exception:
                return True
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    reader = PdfReader(f)
                    if len(reader.pages) == 0:
                        return True
            except Exception:
                return True
        return False

    def is_pdf_not_clear(self, pdf_path, resolution_threshold):
        if not FITZ_AVAILABLE:
            # No image-inspection library available — this check can't run.
            return False
        try:
            doc = fitz.open(pdf_path)
            max_pages = min(self.settings.get("max_pages_check_resolution", 5), len(doc))
            min_text = self.settings.get("min_text_length", 50)
            for page_num in range(max_pages):
                if not self.running:
                    # Cancelled mid-file: bail out of this page loop instead
                    # of only checking between whole files.
                    doc.close()
                    return False
                page = doc[page_num]
                page_text = page.get_text() or ""
                clean_text = page_text.strip()
                has_rich_text = len(clean_text) >= min_text
                page_rect = page.rect

                image_list = page.get_images(full=True)
                if image_list:
                    for img in image_list:
                        if not self.running:
                            doc.close()
                            return False
                        xref = img[0]
                        try:
                            base_image = doc.extract_image(xref)
                            if not base_image:
                                continue
                            width = base_image.get("width", 0)
                            height = base_image.get("height", 0)
                            if width <= 0 or height <= 0:
                                continue

                            # 1. Skip thin rules, dividers, gradient strips, or 1D lines (< 10 px in either dimension)
                            if width < 10 or height < 10:
                                continue

                            # Determine how the image is actually rendered on the page.
                            # get_image_rects gives the bounding box(es) where this image is placed.
                            rects = page.get_image_rects(xref)
                            if rects:
                                for rect in rects:
                                    # 2. Skip tiny decorative icons, bullets, or thin rules (< 0.5 inches in either dimension)
                                    if rect.width < 36 or rect.height < 36:
                                        continue
                                    if rect.width <= 0 or rect.height <= 0:
                                        continue
                                    # Skip decorative banners/strips with extreme aspect ratios (> 15:1)
                                    aspect = max(rect.width, rect.height) / max(min(rect.width, rect.height), 0.1)
                                    if aspect > 15:
                                        continue

                                    disp_w_inch = rect.width / 72.0
                                    disp_h_inch = rect.height / 72.0
                                    dpi_x = width / disp_w_inch
                                    dpi_y = height / disp_h_inch
                                    dpi = min(dpi_x, dpi_y)

                                    # 3. Determine effective threshold:
                                    # If the page has rich vector text, the document content is native digital text.
                                    # Standard screen graphics (>= 70 DPI) are clear on screen; only truly pixelated
                                    # images (< 70 DPI) are flagged. For scanned pages (lacking digital text), the
                                    # full resolution_threshold is enforced (with 5% nominal tolerance).
                                    if has_rich_text:
                                        effective_thresh = min(resolution_threshold, 70)
                                    else:
                                        effective_thresh = resolution_threshold * 0.95

                                    if dpi < effective_thresh:
                                        doc.close()
                                        return True
                            else:
                                # If no bounding box is exposed (e.g. nested in form XObjects),
                                # only evaluate if the pixel dimensions indicate a substantial scan or photo.
                                if width >= 400 or height >= 400:
                                    page_w_inch = page_rect.width / 72.0
                                    page_h_inch = page_rect.height / 72.0
                                    if page_w_inch > 0 and page_h_inch > 0:
                                        dpi = min(width / page_w_inch, height / page_h_inch)
                                        eff = min(resolution_threshold, 70) if has_rich_text else resolution_threshold * 0.95
                                        if dpi < eff:
                                            doc.close()
                                            return True
                        except Exception:
                            continue
            doc.close()
            return False
        except Exception as e:
            logger.warning(f"Could not check image resolution for {pdf_path}: {e}")
            return False

    def has_missing_information(self, pdf_path):
        threshold = self.settings.get("empty_page_threshold", 0.8)
        min_text = self.settings.get("min_text_length", 50)

        if FITZ_AVAILABLE:
            text_content = ""
            page_count = 0
            empty_pages = 0
            total_images = 0
            total_drawings = 0
            total_widgets = 0
            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)
                if page_count == 0:
                    doc.close()
                    return True
                for page in doc:
                    if not self.running:
                        doc.close()
                        return False

                    page_text = page.get_text() or ""
                    clean_text = page_text.strip()
                    if clean_text:
                        text_content += page_text

                    page_images = page.get_images()
                    num_images = len(page_images) if page_images else 0
                    total_images += num_images

                    page_drawings = page.get_drawings()
                    num_drawings = len(page_drawings) if page_drawings else 0
                    total_drawings += num_drawings

                    num_widgets = 0
                    try:
                        widgets = page.widgets()
                        num_widgets = len(list(widgets)) if widgets else 0
                    except Exception:
                        num_widgets = 0
                    total_widgets += num_widgets

                    has_content = bool(clean_text) or (num_images > 0) or (num_drawings > 0) or (num_widgets > 0)
                    if not has_content:
                        empty_pages += 1
                doc.close()
            except Exception as e:
                logger.warning(f"Could not inspect content of {pdf_path}: {e}")
                return False
        else:
            text_content, page_count, empty_pages, total_images = self._extract_text_fallback(pdf_path)
            total_drawings = 0
            total_widgets = 0

        if page_count == 0:
            return True
        if empty_pages > page_count * threshold:
            return True

        # Only enforce min_text if the document lacks visual content (no images, drawings, or form widgets).
        # Legitimate scanned pages, certificates, diagrams, or visual documents naturally have little or no text.
        has_visual_data = (total_images > 0) or (total_drawings > 0) or (total_widgets > 0)
        if not has_visual_data and len(text_content.strip()) < min_text:
            return True

        return False

    def _extract_text_fallback(self, pdf_path):
        """Best-effort text, page count, empty-page count, and image count when PyMuPDF
        isn't available."""
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    reader = PdfReader(f)
                    page_count = len(reader.pages)
                    if page_count == 0:
                        return "", 0, 0, 0
                    text_content = ""
                    empty_pages = 0
                    total_images = 0
                    for page in reader.pages:
                        if not self.running:
                            return text_content, page_count, empty_pages, total_images
                        page_text = ""
                        try:
                            page_text = page.extract_text() or ""
                        except Exception:
                            page_text = ""
                        page_images = 0
                        try:
                            page_images = len(page.images)
                        except Exception:
                            pass
                        total_images += page_images
                        has_content = bool(page_text.strip()) or (page_images > 0)
                        if has_content:
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


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFErrorChecker(root)
    root.mainloop()
