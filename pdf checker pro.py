import os
import sys
import subprocess
import threading
import json
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
CURRENT_VERSION = "0.1.6"


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
            print(f"Error saving settings: {e}")
            return False

    def get(self, key, default=None):
        return self.settings.get(key, default)

    def set(self, key, value):
        self.settings[key] = value


class UpdateChecker:
    """Check for updates from GitHub."""

    def __init__(self, repo, current_version):
        self.repo = repo
        self.current_version = current_version
        self.api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    def check_for_updates(self):
        """Check if new version is available."""
        try:
            response = requests.get(self.api_url, timeout=5)
            if response.status_code == 200:
                data = response.json()
                latest_version = data.get('tag_name', '').lstrip('v')
                if latest_version:
                    return self.compare_versions(latest_version, self.current_version), latest_version, data
            return False, None, None
        except Exception as e:
            print(f"Update check failed: {e}")
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


class SettingsDialog(tk.Toplevel):
    """Professional settings dialog with rounded corners."""

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title("Settings")
        # No fixed 700px height: the notebook is no longer forced to expand
        # to fill it, so the window sizes closer to its actual content and
        # the action buttons sit right under the tabs instead of stranded
        # at the bottom of empty space.
        self.geometry("650x560")
        self.minsize(560, 460)
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#f5f6fa")

        self.settings = settings
        self.modified_settings = settings.settings.copy()

        self.setup_ui()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 2
        self.geometry(f"+{x}+{y}")

    def setup_ui(self):
        main_frame = tk.Frame(self, bg="#f5f6fa", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="Application Settings", font=("Segoe UI", 18, "bold"),
                bg="#f5f6fa", fg="#2c3e50").pack(anchor=tk.W, pady=(0, 15))

        self.notebook = notebook = ttk.Notebook(main_frame)
        # fill=X (not BOTH/expand) so the notebook takes only the height its
        # content needs, keeping Save/Cancel close beneath it instead of
        # being pushed to the bottom of a mostly-empty dialog.
        notebook.pack(fill=tk.X, pady=(0, 15))

        general_frame = tk.Frame(notebook, bg="#f5f6fa", padx=20, pady=20)
        notebook.add(general_frame, text="  General  ")
        self.setup_general_tab(general_frame)

        scan_frame = tk.Frame(notebook, bg="#f5f6fa", padx=20, pady=20)
        notebook.add(scan_frame, text="  Scan Settings  ")
        self.setup_scan_tab(scan_frame)

        about_frame = tk.Frame(notebook, bg="#f5f6fa")
        notebook.add(about_frame, text="  About  ")
        self.setup_about_tab(about_frame)

        # Action buttons live in their own row directly under the tabs, and
        # are hidden while the About tab is showing — About is read-only
        # reference material, not something to Save/Cancel.
        self.button_frame = tk.Frame(main_frame, bg="#f5f6fa")
        self.button_frame.pack(fill=tk.X)

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
            self.button_frame.pack_forget()
        else:
            self.button_frame.pack(fill=tk.X)

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

    def setup_scan_tab(self, parent):
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

    def setup_about_tab(self, parent):
        # Wrapped in a canvas + scrollbar so the About content can scroll if
        # it ever grows past the available height, instead of being clipped
        # or forcing the whole dialog taller.
        canvas = tk.Canvas(parent, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        info_frame = tk.Frame(canvas, bg="#ffffff", padx=20, pady=20)
        info_frame_id = canvas.create_window((0, 0), window=info_frame, anchor="nw")

        def _update_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event):
            canvas.itemconfigure(info_frame_id, width=event.width)

        info_frame.bind("<Configure>", _update_scrollregion)
        canvas.bind("<Configure>", _resize_inner)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

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
                print(f"Error loading about icon: {e}")

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
            self.modified_settings = Settings.DEFAULT_SETTINGS.copy()
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

    def save(self):
        self.modified_settings["window_width"] = self.width_var.get()
        self.modified_settings["window_height"] = self.height_var.get()
        self.modified_settings["resolution_threshold"] = self.resolution_var.get()
        self.modified_settings["max_pages_check_resolution"] = self.max_pages_var.get()
        self.modified_settings["min_text_length"] = self.min_text_var.get()
        self.modified_settings["empty_page_threshold"] = self.empty_ratio_var.get()
        self.modified_settings["auto_check_updates"] = self.auto_update_var.get()

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
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(1000, 650)
        self.root.configure(bg="#f5f6fa")
        self.root.resizable(True, True)

        self.folder_path = tk.StringVar(value=self.settings.get("last_folder", ""))
        self.results = []
        self.running = False
        self.scan_start_time = None
        self.scan_thread = None

        self.export_button = None
        self.scan_button = None
        self.cancel_button = None
        self.settings_button = None

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
        if has_update and data:
            release_url = data.get('html_url', '')
            release_notes = data.get('body') or 'No release notes available.'

            msg = (f"A new version ({latest}) is available!\n\n"
                  f"Current version: {CURRENT_VERSION}\n"
                  f"Latest version: {latest}\n\n"
                  f"Release notes:\n{release_notes[:500]}\n\n"
                  f"Would you like to download it now?")

            if messagebox.askyesno("Update Available", msg):
                if release_url:
                    webbrowser.open(release_url)

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
        main_frame = tk.Frame(self.root, bg="#f5f6fa", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_panel = tk.Frame(main_frame, bg="#f5f6fa", width=380)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_panel.pack_propagate(False)

        self.setup_folder_selection(left_panel)
        self.setup_scan_options(left_panel)
        self.setup_action_buttons(left_panel)

        right_panel = tk.Frame(main_frame, bg="#f5f6fa")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.setup_progress_section(right_panel)
        self.setup_results_section(right_panel)
        self.setup_summary_section(right_panel)

    def setup_folder_selection(self, parent):
        folder_frame = tk.LabelFrame(parent, text="Folder Selection", font=("Segoe UI", 11, "bold"),
                                     bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        folder_frame.pack(fill=tk.X, pady=(0, 15))

        tk.Label(folder_frame, text="Parent folder (will scan all subfolders with 'open' & 'confidential'):",
                font=("Segoe UI", 9), bg="#ffffff", fg="#7f8c8d", wraplength=320).pack(anchor=tk.W, pady=(0, 10))

        folder_entry_frame = tk.Frame(folder_frame, bg="#ffffff")
        folder_entry_frame.pack(fill=tk.X)

        self.folder_entry = tk.Entry(folder_entry_frame, textvariable=self.folder_path, font=("Segoe UI", 9),
                                     bg="#ecf0f1", fg="#2c3e50", relief=tk.FLAT, state="readonly")
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        browse_btn = RoundedButton(folder_entry_frame, text="Browse", command=self.browse_folder,
                                   bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"),
                                   width=100, height=35)
        browse_btn.pack(side=tk.RIGHT, padx=(10, 0))

    def setup_scan_options(self, parent):
        options_frame = tk.LabelFrame(parent, text="Scan Options", font=("Segoe UI", 11, "bold"),
                                      bg="#ffffff", fg="#2c3e50", padx=15, pady=15, relief=tk.FLAT)
        options_frame.pack(fill=tk.X, pady=(0, 15))

        self.check_cannot_open = tk.BooleanVar(value=self.settings.get("check_cannot_open", True))
        self.check_not_clear = tk.BooleanVar(value=self.settings.get("check_not_clear", True))
        self.check_missing_info = tk.BooleanVar(value=self.settings.get("check_missing_info", True))

        checks = [
            ("Cannot Open (Corrupt)", self.check_cannot_open,
             "Flags PDFs that fail to open or report zero pages — likely corrupted files."),
            ("Not Clear (Low Resolution)", self.check_not_clear,
             f"Flags scanned pages with embedded images below the DPI threshold below."
             f"{'' if FITZ_AVAILABLE else ' (Disabled: requires PyMuPDF.)'}"),
            ("Missing Information", self.check_missing_info,
             "Flags PDFs with little or no extractable text, or mostly blank pages."),
        ]

        self._scan_check_vars = [var for _, var, _ in checks]
        for text, var, tip_text in checks:
            cb = tk.Checkbutton(options_frame, text=text, variable=var, font=("Segoe UI", 10),
                          bg="#ffffff", fg="#34495e", selectcolor="#ffffff",
                          activebackground="#ffffff", activeforeground="#34495e")
            cb.pack(anchor=tk.W, pady=5)
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
        res_frame.pack(fill=tk.X, pady=(10, 0))

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
        buttons_frame.pack(fill=tk.X, pady=(15, 0))

        self.scan_button = RoundedButton(buttons_frame, text="Start Scan", command=self.start_scan,
                                         bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"),
                                         width=340, height=45)
        self.scan_button.pack(fill=tk.X, pady=(0, 10))
        Tooltip(self.scan_button, "Scan every project subfolder under the selected parent folder. (Enter)")

        self.cancel_button = RoundedButton(buttons_frame, text="Cancel Scan", command=self.cancel_scan,
                                           bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"),
                                           width=340, height=40, state=tk.DISABLED)
        self.cancel_button.pack(fill=tk.X, pady=(0, 10))
        Tooltip(self.cancel_button, "Stop the scan as soon as possible, mid-file if needed. (Esc)")

        self.export_button = RoundedButton(buttons_frame, text="Export to Word", command=self.export_to_word,
                                           bg="#2980b9", fg="white", font=("Segoe UI", 10, "bold"),
                                           width=340, height=40, state=tk.DISABLED)
        self.export_button.pack(fill=tk.X)
        Tooltip(self.export_button, "Save the current results as a .docx report. (Ctrl+E)")

        clear_btn = RoundedButton(buttons_frame, text="Clear Results", command=self.clear_results,
                                  bg="#95a5a6", fg="white", font=("Segoe UI", 10),
                                  width=340, height=38)
        clear_btn.pack(fill=tk.X, pady=(10, 0))
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
        self.results_tree.tag_configure("not_clear", background="#fff6e5")
        self.results_tree.tag_configure("missing_info", background="#fffbe0")

        self.results_tree.bind("<Double-1>", self._on_result_double_click)
        self.results_tree.bind("<Button-3>", self._on_result_right_click)   # Windows/Linux right-click
        self.results_tree.bind("<Button-2>", self._on_result_right_click)   # macOS right-click

        self._sort_state = {"column": None, "reverse": False}

        self.results_context_menu = tk.Menu(self.results_tree, tearoff=0)
        self.results_context_menu.add_command(label="Open File", command=self._open_selected_result)
        self.results_context_menu.add_command(label="Show in Folder", command=self._reveal_selected_result)
        self.results_context_menu.add_separator()
        self.results_context_menu.add_command(label="Copy Path", command=self._copy_selected_result_path)

    def _tag_for_errors(self, errors):
        """Most severe error determines the row's color tint."""
        if "Cannot Open" in errors:
            return "cannot_open"
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

    def _refresh_results_tree(self):
        """Full rebuild of the visible tree from self.results, honoring the
        current filter text and sort column. Each row's iid is the row's
        index into self.results, so double-click/right-click and re-sorting
        can always map back to the underlying result."""
        if not hasattr(self, "results_tree"):
            return
        self.results_tree.delete(*self.results_tree.get_children())

        filter_text = self.filter_var.get().strip() if hasattr(self, "filter_var") else ""
        visible = [(i, r) for i, r in enumerate(self.results) if self._row_matches_filter(r, filter_text)]

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

    def browse_folder(self):
        folder_selected = filedialog.askdirectory(initialdir=self.settings.get("last_folder", ""))
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.settings.set("last_folder", folder_selected)

    def clear_results(self):
        self.results_tree.delete(*self.results_tree.get_children())
        if hasattr(self, "filter_var"):
            self.filter_var.set("")
        self._sort_state = {"column": None, "reverse": False}
        self.summary_label.config(text="No scan performed yet.")
        self.progress["value"] = 0
        self.status_label.config(text="Ready to scan", fg="#27ae60")
        self.current_file_label.config(text="")
        self.eta_label.config(text="")
        self.results = []
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
            if self.check_not_clear.get():
                settings.append("Not Clear")
            if self.check_missing_info.get():
                settings.append("Missing Information")
            settings_para.add_run(f"{', '.join(settings) if settings else 'None'}")

            doc.add_paragraph(f"Resolution Threshold: {self.resolution_var.get()} DPI")

            doc.add_heading("Summary", level=2)
            total_checked = len(self.results)
            cannot_open_count = sum(1 for r in self.results if "Cannot Open" in r["errors"])
            not_clear_count = sum(1 for r in self.results if "Not Clear" in r["errors"])
            missing_info_count = sum(1 for r in self.results if "Missing Information" in r["errors"])
            unique_parents = set(r.get("parent", "N/A") for r in self.results)

            summary_rows = [
                ["Total PDFs With Errors", str(total_checked)],
                ["Parent Folders Affected", str(len(unique_parents))],
                ["Cannot Open", str(cannot_open_count)],
                ["Not Clear (Low Resolution)", str(not_clear_count)],
                ["Missing Information", str(missing_info_count)],
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
            messagebox.showinfo("Success", f"Report saved to:\n{doc_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Word document:\n{str(e)}")

    def start_scan(self):
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a parent folder first!")
            return

        if not os.path.exists(folder_path):
            messagebox.showerror("Error", "Selected folder does not exist!")
            return

        valid_folders = []
        for root, dirs, files in os.walk(folder_path):
            if "open" in dirs and "confidential" in dirs:
                valid_folders.append(root)

        if not valid_folders:
            messagebox.showerror("Error", "No subfolders with both 'open' and 'confidential' found!\n\n"
                                           "Tip: Each project folder should contain both subfolders.")
            return

        response = messagebox.askyesno("Confirm Scan", f"Found {len(valid_folders)} valid project folders.\n\nContinue?")
        if not response:
            return

        self.running = True
        self.clear_results()
        self.results = []

        if self.scan_button:
            self.scan_button.config(state=tk.DISABLED)
        if self.cancel_button:
            self.cancel_button.config(state=tk.NORMAL)
        if self.export_button:
            self.export_button.config(state=tk.DISABLED)

        self.scan_thread = threading.Thread(target=self.run_scan, args=(valid_folders,), daemon=True)
        self.scan_thread.start()

    def run_scan(self, valid_folders):
        check_cannot_open = self.check_cannot_open.get()
        check_not_clear = self.check_not_clear.get()
        check_missing_info = self.check_missing_info.get()
        resolution_threshold = self.resolution_var.get()

        all_pdfs = []
        for folder in valid_folders:
            parent_name = os.path.basename(folder)
            for subfolder in ["open", "confidential"]:
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

        # Fixed: progress bar defaults to maximum=100, so folders with more
        # than 100 PDFs would hit 100% long before the scan actually finished.
        self.root.after(0, lambda: self.progress.config(maximum=total_pdfs, value=0))

        for i, pdf_info in enumerate(all_pdfs):
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

            errors = self.check_pdf(pdf_path, check_cannot_open, check_not_clear, check_missing_info, resolution_threshold)

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

        self.root.after(0, lambda: self.finish_scan(all_pdfs))

    @staticmethod
    def _format_duration(total_seconds):
        total_seconds = max(0, int(total_seconds))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"{minutes}:{seconds:02d}"

    def finish_scan(self, all_pdfs):
        total_checked = len(all_pdfs)
        total_errors = len(self.results)
        scan_duration = datetime.now() - self.scan_start_time
        scan_duration_seconds = scan_duration.total_seconds()
        unique_parents = len(set(r.get("parent", "N/A") for r in self.results))

        status_word = "cancelled" if not self.running else "completed"

        summary_text = (f"Scan {status_word}: {total_checked} PDFs in {unique_parents} folders | "
                        f"Errors: {total_errors} | Time: {scan_duration_seconds:.1f}s")

        self.summary_label.config(text=summary_text, fg="#27ae60" if self.running else "#e67e22")
        self.status_label.config(text=f"Scan {status_word}", fg="#27ae60" if self.running else "#e67e22")
        self.current_file_label.config(text="")
        self.eta_label.config(text=f"Total time: {self._format_duration(scan_duration_seconds)}")
        was_cancelled = not self.running
        self.running = False

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

    def check_pdf(self, pdf_path, check_cannot_open, check_not_clear, check_missing_info, resolution_threshold):
        # Finer-grained cancel: previously Cancel only took effect between
        # files, so one huge multi-page PDF could stall it noticeably.
        # Checking self.running between (and inside) each sub-check lets a
        # cancel land mid-file instead of only at file boundaries.
        errors = []
        if not self.running:
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
            for page_num in range(max_pages):
                if not self.running:
                    # Cancelled mid-file: bail out of this page loop instead
                    # of only checking between whole files.
                    doc.close()
                    return False
                page = doc[page_num]
                image_list = page.get_images(full=True)
                if image_list:
                    for img in image_list:
                        if not self.running:
                            doc.close()
                            return False
                        xref = img[0]
                        try:
                            base_image = doc.extract_image(xref)
                            if base_image:
                                width = base_image.get("width", 0)
                                height = base_image.get("height", 0)
                                page_rect = page.rect
                                page_width_inch = page_rect.width / 72
                                page_height_inch = page_rect.height / 72
                                if page_width_inch > 0 and page_height_inch > 0:
                                    dpi = min(width/page_width_inch, height/page_height_inch)
                                    if dpi < resolution_threshold:
                                        doc.close()
                                        return True
                        except Exception:
                            continue
            doc.close()
            return False
        except Exception:
            return False

    def has_missing_information(self, pdf_path):
        threshold = self.settings.get("empty_page_threshold", 0.8)
        min_text = self.settings.get("min_text_length", 50)

        if FITZ_AVAILABLE:
            text_content = ""
            page_count = 0
            empty_pages = 0
            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)
                if page_count == 0:
                    doc.close()
                    return True
                for page in doc:
                    if not self.running:
                        # Cancelled mid-file — stop scanning pages; the scan
                        # loop will discard this result anyway.
                        doc.close()
                        return False
                    text = page.get_text()
                    if text:
                        text_content += text
                    else:
                        empty_pages += 1
                doc.close()
            except Exception:
                return True
        else:
            # Fixed: previously there was no fallback at all when PyMuPDF was
            # missing, so page_count stayed 0 and every PDF was flagged as
            # "Missing Information" regardless of its actual content — and
            # even after adding a fallback, the empty-page-ratio half of this
            # check was still being skipped. It now runs here too, using
            # real per-page results from the fallback extractor.
            text_content, page_count, empty_pages = self._extract_text_fallback(pdf_path)

        if page_count == 0:
            return True
        if empty_pages > page_count * threshold:
            return True
        if len(text_content.strip()) < min_text:
            return True
        return False

    def _extract_text_fallback(self, pdf_path):
        """Best-effort text, page count, and empty-page count when PyMuPDF
        isn't available.

        Prefers PyPDF2 because it can iterate pages individually, giving a
        real page count and a real empty-page count (needed for the
        empty-page-ratio check) rather than one blob of text for the whole
        document. Falls back to pdfminer only for the min-text-length part
        of the check if PyPDF2 isn't installed either — pdfminer's simple
        API doesn't expose an efficient per-page count, so in that case the
        whole document is treated as a single "page" and the empty-page-ratio
        check is effectively a no-op (documented here rather than silently
        skipped)."""
        if PYPDF2_AVAILABLE:
            try:
                with open(pdf_path, "rb") as f:
                    reader = PdfReader(f)
                    page_count = len(reader.pages)
                    if page_count == 0:
                        return "", 0, 0
                    text_content = ""
                    empty_pages = 0
                    for page in reader.pages:
                        if not self.running:
                            # Cancelled mid-file — return what's been read so
                            # far; the caller discards this result anyway.
                            return text_content, page_count, empty_pages
                        try:
                            page_text = page.extract_text() or ""
                        except Exception:
                            page_text = ""
                        if page_text.strip():
                            text_content += page_text
                        else:
                            empty_pages += 1
                    return text_content, page_count, empty_pages
            except Exception:
                pass

        if PDFMINER_AVAILABLE:
            try:
                text = pdfminer_extract_text(pdf_path) or ""
                # No cheap per-page count available here; treat the whole
                # document as one unit so only the min-text-length check
                # is meaningful in this path.
                page_count = 1 if text.strip() else 0
                empty_pages = 0 if text.strip() else 1
                return text, page_count, empty_pages
            except Exception:
                pass

        return "", 0, 0


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFErrorChecker(root)
    root.mainloop()
