"""
pdf_checker_qt.py — PySide6 front end for PDF Error Checker Pro.

Runs on pdf_checker_core.PDFScanEngine, which contains all scanning logic
and knows nothing about Qt. This file is UI only: layout, styling, threading
glue (QThread + signals), and the results table's virtualized model.

Install: pip install PySide6 (plus the same optional deps as before:
PyMuPDF, PyPDF2, pdfminer.six, python-docx, Pillow).
"""
import os
import sys
import subprocess
import threading
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QObject, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QTimer
from PySide6.QtGui import QIcon, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QRadioButton, QCheckBox, QSpinBox, QDoubleSpinBox,
    QGroupBox, QTableView, QProgressBar, QFileDialog, QMessageBox, QTabWidget, QDialog,
    QMenu, QSystemTrayIcon, QHeaderView, QComboBox, QDialogButtonBox, QScrollArea,
    QSplitter, QAbstractItemView, QSizePolicy, QProgressDialog, QInputDialog, QStackedWidget,
)

from pdf_checker_core import (
    Settings, ScanCache, ScanHistory, UpdateChecker, PDFScanEngine, ScanResult,
    export_results_to_word, download_asset, apply_windows_exe_update, apply_source_update,
    CURRENT_VERSION, GITHUB_REPO,
    FITZ_AVAILABLE, PYPDF2_AVAILABLE, PDFMINER_AVAILABLE, DOCX_AVAILABLE, logger,
)

# icon.ico/icon.png are read-only bundled assets, unlike the persistent
# settings/cache/history/log files in pdf_checker_core.py — sys._MEIPASS
# (PyInstaller's onefile extraction temp dir) is exactly where such
# read-only resources are meant to be looked up at runtime when frozen.
APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))

# A single flat, rounded, slightly darker-neutral palette applied app-wide —
# this is the "reskin" half of the modernization; the responsiveness half is
# the QThread + virtualized-model architecture below.
MODERN_QSS = """
QWidget { background-color: #f5f6fa; color: #2c3e50; font-family: 'Segoe UI'; font-size: 10pt; }
QMainWindow { background-color: #f5f6fa; }
QGroupBox {
    background-color: #ffffff; border-radius: 10px; border: 1px solid #e6e8ec;
    margin-top: 14px; padding: 12px; font-weight: 600;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; color: #2c3e50; }
QPushButton {
    border-radius: 10px; padding: 8px 16px; font-weight: 600; color: white;
    background-color: #3498db; border: none;
}
QPushButton:hover { background-color: #55acee; }
QPushButton:pressed { background-color: #2c81ba; }
QPushButton:disabled { background-color: #dbe9f5; color: #a9c6dd; }
QPushButton#scanButton { background-color: #27ae60; }
QPushButton#scanButton:hover { background-color: #37c172; }
QPushButton#scanButton:disabled { background-color: #d6f0e0; color: #9fd9bb; }
QPushButton#cancelButton { background-color: #e74c3c; }
QPushButton#cancelButton:hover { background-color: #f0685a; }
QPushButton#cancelButton:disabled { background-color: #fbdedb; color: #f0aca3; }
QPushButton#secondaryButton { background-color: #95a5a6; }
QPushButton#secondaryButton:hover { background-color: #aab7b8; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
    border: 1px solid #dfe3e8; border-radius: 8px; padding: 6px 8px; background: #ffffff;
}
QTableView {
    border: 1px solid #e6e8ec; border-radius: 10px; background: #ffffff;
    gridline-color: #f0f1f3; selection-background-color: #dceefc; selection-color: #2c3e50;
}
QHeaderView::section {
    background-color: #f5f6fa; border: none; border-bottom: 1px solid #e6e8ec;
    padding: 6px; font-weight: 600;
}
QProgressBar { border-radius: 8px; background: #eef0f3; text-align: center; height: 18px; }
QProgressBar::chunk { border-radius: 8px; background-color: #3498db; }
QTabWidget::pane { border: 1px solid #e6e8ec; border-radius: 8px; }
QTabBar::tab { padding: 8px 16px; border-top-left-radius: 8px; border-top-right-radius: 8px; }
QTabBar::tab:selected { background: #ffffff; font-weight: 600; }
"""

SEVERITY_COLORS = {
    "Cannot Open": QColor("#fdecea"),
    "Password Protected": QColor("#f3e6fb"),
    "Not Clear": QColor("#fff6e5"),
    "Missing Information": QColor("#fffbe0"),
}


def _severity_color(errors):
    for label in ("Cannot Open", "Password Protected"):
        if label in errors:
            return SEVERITY_COLORS[label]
    if any(e.startswith("Duplicate") for e in errors):
        return QColor("#e6f0fa")
    for label in ("Not Clear", "Missing Information"):
        if label in errors:
            return SEVERITY_COLORS[label]
    return None


# ---------------------------------------------------------------------------
# Results table: a real virtualized model, so scrolling/sorting/filtering
# thousands of results stays smooth instead of degrading the way a widget-
# per-row Tkinter Treeview would.
# ---------------------------------------------------------------------------

class ResultsTableModel(QAbstractTableModel):
    HEADERS = ["File", "Parent Folder", "Subfolder", "Errors"]

    def __init__(self):
        super().__init__()
        self._results: list[ScanResult] = []

    def rowCount(self, parent=QModelIndex()):
        return len(self._results)

    def columnCount(self, parent=QModelIndex()):
        return len(self.HEADERS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        result = self._results[index.row()]
        col = index.column()
        if role == Qt.DisplayRole:
            return [result.filename, result.parent, result.folder, ", ".join(result.errors)][col]
        if role == Qt.BackgroundRole:
            return _severity_color(result.errors)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEADERS[section]
        return None

    def add_result(self, result: ScanResult):
        row = len(self._results)
        self.beginInsertRows(QModelIndex(), row, row)
        self._results.append(result)
        self.endInsertRows()

    def update_result(self, result: ScanResult):
        """Re-renders the row for a ScanResult whose .errors list changed
        in place after it was already added (duplicate detection appending
        a "Duplicate" error to a file that already had one from pass 1).
        Without this, the underlying data changes but the view — which
        only repaints in response to dataChanged — never finds out, so the
        new error is invisible on screen despite being present in the
        object and correctly showing up in the exported Word report."""
        try:
            row = self._results.index(result)
        except ValueError:
            return
        top_left = self.index(row, 0)
        bottom_right = self.index(row, self.columnCount() - 1)
        self.dataChanged.emit(top_left, bottom_right)

    def clear(self):
        self.beginResetModel()
        self._results = []
        self.endResetModel()

    def result_at_source_row(self, row):
        return self._results[row]

    def all_results(self):
        return list(self._results)


# ---------------------------------------------------------------------------
# Scan worker: runs PDFScanEngine.run_scan() on a background QThread.
# should_continue is a plain threading.Event so the engine (which knows
# nothing about Qt) can be polled cheaply and thread-safely from the
# scanning thread while Cancel is clicked on the main thread.
# ---------------------------------------------------------------------------

class ScanWorker(QObject):
    progress = Signal(object)   # ScanProgress
    result_found = Signal(object)  # ScanResult
    result_updated = Signal(object)  # ScanResult whose errors changed in place (e.g. a duplicate found after the fact)
    status = Signal(str)
    finished = Signal(object, list)  # ScanSummary, list[ScanResult]
    failed = Signal(str)  # error message, if run_scan itself raised

    def __init__(self, engine: PDFScanEngine, folder, scan_mode, check_flags, resolution_threshold):
        super().__init__()
        self.engine = engine
        self.folder = folder
        self.scan_mode = scan_mode
        self.check_flags = check_flags
        self.resolution_threshold = resolution_threshold

    def run(self):
        try:
            summary, results = self.engine.run_scan(
                self.folder, self.scan_mode, self.check_flags, self.resolution_threshold,
                on_progress=self.progress.emit, on_result=self.result_found.emit,
                on_status=self.status.emit, on_result_updated=self.result_updated.emit,
            )
        except Exception as e:
            # Fixed: an unhandled exception here (e.g. a directory
            # permission error, a drive disconnecting mid-scan) used to
            # kill this thread silently — finished never fired, so the
            # MainWindow's buttons stayed disabled and the progress bar
            # stayed frozen indefinitely, with no way to recover short of
            # restarting the whole app. Catching it and emitting a
            # dedicated failed signal lets the UI unlock and tell the user
            # what happened instead.
            logger.error(f"Scan worker crashed: {e}")
            self.failed.emit(str(e))
            return
        self.finished.emit(summary, results)


class UpdateCheckWorker(QObject):
    """Runs the one-shot GitHub release check off the main thread."""
    result_ready = Signal(object)  # (has_update, latest_version, release_data)

    def run(self):
        checker = UpdateChecker(GITHUB_REPO, CURRENT_VERSION)
        self.result_ready.emit(checker.check_for_updates())


class DownloadWorker(QObject):
    """Streams a release asset to disk, reporting fractional progress."""
    progress = Signal(float)
    finished = Signal(str)   # downloaded file path
    failed = Signal(str)     # error message

    def __init__(self, url, dest_path, size_hint=0):
        super().__init__()
        self.url = url
        self.dest_path = dest_path
        self.size_hint = size_hint

    def run(self):
        try:
            path = download_asset(self.url, self.dest_path, on_progress=self.progress.emit,
                                  size_hint=self.size_hint)
            self.finished.emit(str(path))
        except Exception as e:
            self.failed.emit(str(e))


# ---------------------------------------------------------------------------
# Settings dialog
# ---------------------------------------------------------------------------

class SettingsDialog(QDialog):
    def __init__(self, parent, settings: Settings, cache: ScanCache):
        super().__init__(parent)
        self.settings = settings
        self.cache = cache
        self.setWindowTitle("Settings")
        self.resize(560, 520)

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs, 1)

        tabs.addTab(self._build_general_tab(), "General")
        tabs.addTab(self._build_scan_tab(), "Scan Settings")
        tabs.addTab(self._build_about_tab(), "About")

        buttons = QDialogButtonBox()
        reset_btn = buttons.addButton("Reset to Defaults", QDialogButtonBox.ResetRole)
        save_btn = buttons.addButton("Save", QDialogButtonBox.AcceptRole)
        cancel_btn = buttons.addButton("Cancel", QDialogButtonBox.RejectRole)
        reset_btn.clicked.connect(self._reset_defaults)
        save_btn.clicked.connect(self._save)
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(buttons)

        # Hide the button row on the About tab — it's read-only reference
        # material, not something to Save/Cancel.
        tabs.currentChanged.connect(lambda i: buttons.setVisible(tabs.tabText(i) != "About"))

    def _build_general_tab(self):
        widget = QWidget()
        form = QFormLayout(widget)

        self.auto_update_cb = QCheckBox("Automatically check for updates on startup")
        self.auto_update_cb.setChecked(self.settings.get("auto_check_updates", True))
        form.addRow(self.auto_update_cb)

        self.enable_cache_cb = QCheckBox("Skip re-checking unchanged files (cache by size + mtime)")
        self.enable_cache_cb.setChecked(self.settings.get("enable_scan_cache", True))
        self.enable_cache_cb.setToolTip(
            "If a file's size and modified-date haven't changed since the last scan (with "
            "the same settings), its previous result is reused instead of re-checking it.")
        form.addRow(self.enable_cache_cb)

        clear_cache_btn = QPushButton("Clear Cache")
        clear_cache_btn.setObjectName("secondaryButton")
        clear_cache_btn.clicked.connect(self._clear_cache)
        form.addRow(clear_cache_btn)

        # Restored (items 3 & 7): window size is both configurable and
        # persisted here, matching v1. The main window applies it
        # immediately on save (see MainWindow.open_settings) and saves the
        # user's current size back on close (see MainWindow.closeEvent) —
        # neither existed at all in the initial Qt port.
        self.width_spin = QSpinBox()
        self.width_spin.setRange(800, 3840)
        self.width_spin.setValue(self.settings.get("window_width", 1200))
        self.height_spin = QSpinBox()
        self.height_spin.setRange(500, 2160)
        self.height_spin.setValue(self.settings.get("window_height", 850))
        size_row = QHBoxLayout()
        size_row.addWidget(self.width_spin)
        size_row.addWidget(QLabel("×"))
        size_row.addWidget(self.height_spin)
        size_note = QLabel("Applies immediately on save. Also updated automatically when you resize the window.")
        size_note.setStyleSheet("color:#95a5a6; font-size:8pt; font-style:italic;")
        size_note.setWordWrap(True)
        form.addRow("Window size:", size_row)
        form.addRow(size_note)
        return widget

    def _clear_cache(self):
        self.cache.clear()
        QMessageBox.information(self, "Cache Cleared",
                                "The scan cache has been cleared. The next scan will re-check every file.")

    def _profile_values_from_ui(self):
        return {
            "resolution_threshold": self.dpi_spin.value(),
            "max_pages_check_resolution": self.pages_spin.value(),
            "min_text_length": self.min_text_spin.value(),
            "empty_page_threshold": self.empty_ratio_spin.value(),
            "target_subfolders": self.target_subfolders_edit.text(),
            "use_folder_aliases": self.use_aliases_cb.isChecked(),
        }

    def _apply_profile_values(self, values):
        self.dpi_spin.setValue(values.get("resolution_threshold", self.dpi_spin.value()))
        self.pages_spin.setValue(values.get("max_pages_check_resolution", self.pages_spin.value()))
        self.min_text_spin.setValue(values.get("min_text_length", self.min_text_spin.value()))
        self.empty_ratio_spin.setValue(values.get("empty_page_threshold", self.empty_ratio_spin.value()))
        self.target_subfolders_edit.setText(values.get("target_subfolders", self.target_subfolders_edit.text()))
        self.use_aliases_cb.setChecked(values.get("use_folder_aliases", self.use_aliases_cb.isChecked()))

    def _persist_profiles(self, profiles):
        self.settings.set("profiles", profiles)
        self.settings.save(self.settings.settings)
        self.profile_combo.clear()
        self.profile_combo.addItems(list(profiles.keys()))

    def _save_profile_as(self):
        name, ok = QInputDialog.getText(self, "Save Profile", "Profile name:")
        if not ok or not name.strip():
            return
        name = name.strip()
        profiles = dict(self.settings.get("profiles", {}))
        profiles[name] = self._profile_values_from_ui()
        self._persist_profiles(profiles)
        self.profile_combo.setCurrentText(name)
        QMessageBox.information(self, "Profile Saved", f"Saved profile '{name}'.")

    def _load_profile(self):
        name = self.profile_combo.currentText()
        if not name:
            QMessageBox.warning(self, "No Profile Selected", "Choose a profile to load first.")
            return
        values = self.settings.get("profiles", {}).get(name)
        if values is None:
            QMessageBox.critical(self, "Not Found", f"Profile '{name}' no longer exists.")
            return
        self._apply_profile_values(values)

    def _delete_profile(self):
        name = self.profile_combo.currentText()
        if not name:
            return
        if QMessageBox.question(self, "Delete Profile", f"Delete profile '{name}'?") != QMessageBox.Yes:
            return
        profiles = dict(self.settings.get("profiles", {}))
        profiles.pop(name, None)
        self._persist_profiles(profiles)

    def _build_scan_tab(self):
        widget = QScrollArea()
        widget.setWidgetResizable(True)
        inner = QWidget()
        form = QFormLayout(inner)
        widget.setWidget(inner)

        if not FITZ_AVAILABLE:
            warn = QLabel("PyMuPDF is not installed. Resolution checks are disabled and "
                          "text-based checks fall back to a slower reader.\nInstall with: pip install pymupdf")
            warn.setStyleSheet("background:#fdecea; color:#c0392b; padding:8px; border-radius:8px;")
            warn.setWordWrap(True)
            form.addRow(warn)

        # Scan profiles: saved and reloaded independently of Save/Cancel —
        # a saved profile shouldn't vanish just because the user later
        # cancels an unrelated threshold edit in this same dialog session.
        profile_row = QHBoxLayout()
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(list(self.settings.get("profiles", {}).keys()))
        load_profile_btn = QPushButton("Load")
        save_profile_btn = QPushButton("Save As...")
        save_profile_btn.setObjectName("secondaryButton")
        delete_profile_btn = QPushButton("Delete")
        delete_profile_btn.setObjectName("cancelButton")
        load_profile_btn.clicked.connect(self._load_profile)
        save_profile_btn.clicked.connect(self._save_profile_as)
        delete_profile_btn.clicked.connect(self._delete_profile)
        profile_row.addWidget(self.profile_combo, 1)
        profile_row.addWidget(load_profile_btn)
        profile_row.addWidget(save_profile_btn)
        profile_row.addWidget(delete_profile_btn)
        form.addRow("Scan Profiles:", profile_row)

        self.target_subfolders_edit = QLineEdit(self.settings.get("target_subfolders", "open, confidential"))
        self.target_subfolders_edit.setToolTip(
            "e.g. 'open, confidential' or 'public, internal'. A project folder is scanned "
            "if it contains at least one of these.")
        form.addRow("Target subfolders:", self.target_subfolders_edit)

        self.use_aliases_cb = QCheckBox("Also match common naming variations")
        self.use_aliases_cb.setChecked(self.settings.get("use_folder_aliases", True))
        self.use_aliases_cb.setToolTip(
            "Recognizes variations like 'opened', 'public', 'conf', 'restricted', 'private', etc.")
        form.addRow(self.use_aliases_cb)

        subfolder_note = QLabel(
            "Matching is always case-insensitive, and a project folder no longer needs "
            "every target subfolder to be scanned — just one.")
        subfolder_note.setWordWrap(True)
        subfolder_note.setStyleSheet("color:#95a5a6; font-size:8pt; font-style:italic;")
        form.addRow(subfolder_note)

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(72, 600)
        self.dpi_spin.setValue(self.settings.get("resolution_threshold", 150))
        self.dpi_spin.setToolTip(
            "Pages with embedded images below this DPI are flagged as low resolution.")
        form.addRow("Minimum DPI:", self.dpi_spin)

        self.pages_spin = QSpinBox()
        self.pages_spin.setRange(1, 20)
        self.pages_spin.setValue(self.settings.get("max_pages_check_resolution", 5))
        form.addRow("Pages to check (resolution):", self.pages_spin)

        self.min_text_spin = QSpinBox()
        self.min_text_spin.setRange(0, 1000)
        self.min_text_spin.setValue(self.settings.get("min_text_length", 50))
        form.addRow("Minimum text length (chars):", self.min_text_spin)

        self.empty_ratio_spin = QDoubleSpinBox()
        self.empty_ratio_spin.setRange(0.1, 1.0)
        self.empty_ratio_spin.setSingleStep(0.05)
        self.empty_ratio_spin.setValue(self.settings.get("empty_page_threshold", 0.8))
        form.addRow("Empty page ratio threshold:", self.empty_ratio_spin)

        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 16)
        self.workers_spin.setValue(self.settings.get("max_workers", 4))
        self.workers_spin.setToolTip(
            "How many PDFs to check at once. Higher values speed up large folders on "
            "multi-core machines, but too many can thrash a slow disk.")
        form.addRow("Parallel workers:", self.workers_spin)

        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(5, 300)
        self.timeout_spin.setValue(self.settings.get("per_file_timeout_seconds", 30))
        self.timeout_spin.setToolTip(
            "If a single PDF takes longer than this to check, it's marked 'Scan Timeout' "
            "and the scan moves on instead of stalling on it.")
        form.addRow("Per-file timeout (sec):", self.timeout_spin)

        return widget

    def _build_about_tab(self):
        widget = QScrollArea()
        widget.setWidgetResizable(True)
        inner = QWidget()
        layout = QVBoxLayout(inner)
        widget.setWidget(inner)

        title = QLabel("PDF Error Checker Pro")
        title.setStyleSheet("font-size:16pt; font-weight:700;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Check .png first (better for scaled display) but fall back to
        # .ico — the window/taskbar icon already uses .ico, so checking
        # only .png meant a project shipping just an icon.ico showed an
        # icon everywhere EXCEPT here.
        icon_path = next((p for p in (APP_DIR / "icon.png", APP_DIR / "icon.ico") if p.exists()), None)
        if icon_path is not None:
            pixmap = QIcon(str(icon_path)).pixmap(64, 64)
            if not pixmap.isNull():
                icon_label = QLabel()
                icon_label.setPixmap(pixmap)
                icon_label.setAlignment(Qt.AlignCenter)
                layout.addWidget(icon_label)

        version = QLabel(f"Version {CURRENT_VERSION}")
        version.setAlignment(Qt.AlignCenter)
        layout.addWidget(version)

        credit_intro = QLabel("Designed and developed by")
        credit_intro.setAlignment(Qt.AlignCenter)
        credit_intro.setStyleSheet("color:#95a5a6; font-size:9pt;")
        layout.addWidget(credit_intro)

        credit_names = QLabel("Brandon & Ian")
        credit_names.setAlignment(Qt.AlignCenter)
        credit_names.setStyleSheet("color:#3498db; font-size:13pt; font-weight:700;")
        layout.addWidget(credit_names)

        desc = QLabel(
            "A professional tool for scanning and validating PDF files across multiple "
            "project folders.\n\nFeatures: multi-folder scanning, parallel background "
            "processing with per-file timeouts, scan result caching, duplicate detection, "
            "password-protection detection, scan history, Word report export, and "
            "automatic update checking."
        )
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc)

        deps = "\n".join([
            f"- PyMuPDF: {'OK' if FITZ_AVAILABLE else 'MISSING'}",
            f"- PyPDF2: {'OK' if PYPDF2_AVAILABLE else 'MISSING'}",
            f"- pdfminer: {'OK' if PDFMINER_AVAILABLE else 'MISSING'}",
            f"- python-docx: {'OK' if DOCX_AVAILABLE else 'MISSING'}",
        ])
        deps_label = QLabel(deps)
        deps_label.setStyleSheet("background:#ecf0f1; padding:10px; border-radius:8px;")
        layout.addWidget(deps_label)

        footer = QLabel("Built with Python & Qt (PySide6)")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color:#95a5a6; font-style:italic; font-size:9pt;")
        layout.addWidget(footer)

        layout.addStretch()
        return widget

    def _reset_defaults(self):
        if QMessageBox.question(self, "Reset Settings", "Are you sure?") == QMessageBox.Yes:
            defaults = Settings.DEFAULT_SETTINGS
            self.auto_update_cb.setChecked(defaults["auto_check_updates"])
            self.enable_cache_cb.setChecked(defaults["enable_scan_cache"])
            self.width_spin.setValue(defaults["window_width"])
            self.height_spin.setValue(defaults["window_height"])
            self.target_subfolders_edit.setText(defaults["target_subfolders"])
            self.use_aliases_cb.setChecked(defaults["use_folder_aliases"])
            self.dpi_spin.setValue(defaults["resolution_threshold"])
            self.pages_spin.setValue(defaults["max_pages_check_resolution"])
            self.min_text_spin.setValue(defaults["min_text_length"])
            self.empty_ratio_spin.setValue(defaults["empty_page_threshold"])
            self.workers_spin.setValue(defaults["max_workers"])
            self.timeout_spin.setValue(defaults["per_file_timeout_seconds"])

    def _save(self):
        self.settings.set("auto_check_updates", self.auto_update_cb.isChecked())
        self.settings.set("enable_scan_cache", self.enable_cache_cb.isChecked())
        self.settings.set("window_width", self.width_spin.value())
        self.settings.set("window_height", self.height_spin.value())
        self.settings.set("target_subfolders", self.target_subfolders_edit.text())
        self.settings.set("use_folder_aliases", self.use_aliases_cb.isChecked())
        self.settings.set("resolution_threshold", self.dpi_spin.value())
        self.settings.set("max_pages_check_resolution", self.pages_spin.value())
        self.settings.set("min_text_length", self.min_text_spin.value())
        self.settings.set("empty_page_threshold", self.empty_ratio_spin.value())
        self.settings.set("max_workers", self.workers_spin.value())
        self.settings.set("per_file_timeout_seconds", self.timeout_spin.value())
        if self.settings.save(self.settings.settings):
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to save settings.")


# ---------------------------------------------------------------------------
# History dialog
# ---------------------------------------------------------------------------

class HistoryDialog(QDialog):
    def __init__(self, parent, history: ScanHistory):
        super().__init__(parent)
        self.history = history
        self.setWindowTitle("Scan History")
        self.resize(760, 420)

        layout = QVBoxLayout(self)
        from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Date", "Folder", "Mode", "Files Checked", "Errors Found", "Duration (s)"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)
        self._populate()

        btn_row = QHBoxLayout()
        clear_btn = QPushButton("Clear History")
        clear_btn.setObjectName("cancelButton")
        clear_btn.clicked.connect(self._clear)
        close_btn = QPushButton("Close")
        close_btn.setObjectName("secondaryButton")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _populate(self):
        from PySide6.QtWidgets import QTableWidgetItem
        self.table.setRowCount(0)
        for entry in self.history.entries:
            row = self.table.rowCount()
            self.table.insertRow(row)
            mode = "All PDFs" if entry.get("scan_mode") == "all" else "Project"
            if entry.get("cancelled"):
                mode += " (cancelled)"
            values = [entry.get("timestamp", ""), entry.get("folder", ""), mode,
                     str(entry.get("files_checked", 0)), str(entry.get("errors_found", 0)),
                     str(entry.get("duration_seconds", 0))]
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))

    def _clear(self):
        if QMessageBox.question(self, "Clear History", "Remove all recorded scan history?") == QMessageBox.Yes:
            self.history.clear()
            self._populate()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Error Checker Pro")

        self.settings = Settings()

        icon_path = APP_DIR / "icon.ico"
        if not icon_path.exists():
            icon_path = APP_DIR / "icon.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        # Restored (items 3 & 7): the Tkinter version explicitly clamped
        # startup size to the screen and enforced a minimum, specifically
        # to keep action buttons from being pushed off-screen on smaller
        # displays (1366x768 was called out by name) — the initial Qt port
        # dropped this entirely in favor of a flat resize(1200, 850) with
        # no clamp and no minimum at all.
        width = self.settings.get("window_width", 1200)
        height = self.settings.get("window_height", 850)
        screen = QApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()
            width = min(width, max(850, available.width() - 60))
            height = min(height, max(500, available.height() - 90))
        self.resize(width, height)
        self.setMinimumSize(800, 500)
        self.cache = ScanCache()
        self.history = ScanHistory()

        # A threading.Event, not a plain bool: this is what should_continue
        # polls from the scan worker thread, and it's what Cancel sets from
        # the main thread. Using Event (rather than a bare attribute) makes
        # the intent explicit and is trivially thread-safe.
        self._cancel_event = threading.Event()
        self._cancel_event.set()  # "should continue" == event is set
        self.engine = PDFScanEngine(self.settings, self.cache, should_continue=self._cancel_event.is_set)

        self.thread = None
        self.worker = None
        self._update_thread = None
        self._update_worker = None
        self._download_thread = None
        self._download_worker = None
        self.results_model = ResultsTableModel()
        self.has_scanned = False

        self.tray_icon = None
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray_icon = QSystemTrayIcon(self.windowIcon(), self)
            self.tray_icon.show()

        self._build_ui()
        self._setup_shortcuts()
        QTimer.singleShot(1000, self._check_updates_async)

    # ---- layout -----------------------------------------------------------

    def _build_status_bar(self):
        bar = self.statusBar()
        version_label = QLabel(f"PDF Error Checker Pro v{CURRENT_VERSION}")
        bar.addWidget(version_label)

        deps = []
        if FITZ_AVAILABLE:
            deps.append("PyMuPDF OK")
        if PYPDF2_AVAILABLE:
            deps.append("PyPDF2 OK")
        if deps:
            deps_label = QLabel(" | ".join(deps))
            deps_label.setStyleSheet("color:#27ae60;")
            bar.addPermanentWidget(deps_label)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)

        self._build_status_bar()

        header = QHBoxLayout()
        title = QLabel("PDF Error Checker Pro")
        title.setStyleSheet("font-size:16pt; font-weight:700;")
        header.addWidget(title)
        header.addStretch()
        history_btn = QPushButton("History")
        history_btn.setObjectName("secondaryButton")
        history_btn.clicked.connect(self.open_history)
        settings_btn = QPushButton("Settings")
        settings_btn.setObjectName("secondaryButton")
        settings_btn.clicked.connect(self.open_settings)
        self.history_btn, self.settings_btn = history_btn, settings_btn
        header.addWidget(history_btn)
        header.addWidget(settings_btn)
        outer.addLayout(header)

        if not FITZ_AVAILABLE:
            banner = QLabel("PyMuPDF is not installed — resolution checks are disabled and the "
                            "missing-information check uses a slower fallback reader.")
            banner.setStyleSheet("background:#fdecea; color:#c0392b; padding:8px; border-radius:8px;")
            banner.setWordWrap(True)
            outer.addWidget(banner)

        splitter = QSplitter(Qt.Horizontal)
        outer.addWidget(splitter, 1)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFixedWidth(360)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_scroll.setWidget(left_panel)
        splitter.addWidget(left_scroll)

        left_layout.addWidget(self._build_folder_group())
        left_layout.addWidget(self._build_scan_mode_group())
        left_layout.addWidget(self._build_scan_options_group())
        left_layout.addWidget(self._build_action_buttons())
        left_layout.addStretch()

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

        right_layout.addWidget(self._build_progress_group())
        right_layout.addWidget(self._build_results_group(), 1)
        self.summary_label = QLabel("No scan performed yet.")
        self.summary_label.setStyleSheet("color:#7f8c8d; padding:8px;")
        right_layout.addWidget(self.summary_label)

    def _build_folder_group(self):
        box = QGroupBox("Folder Selection")
        layout = QVBoxLayout(box)
        hint = QLabel("Parent folder (will scan matching subfolders, or every PDF in 'All PDFs' mode):")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#7f8c8d; font-size:9pt;")
        layout.addWidget(hint)
        row = QHBoxLayout()
        self.folder_edit = QLineEdit(self.settings.get("last_folder", ""))
        self.folder_edit.setReadOnly(True)
        self.browse_btn = QPushButton("Browse")
        self.browse_btn.clicked.connect(self.browse_folder)
        row.addWidget(self.folder_edit)
        row.addWidget(self.browse_btn)
        layout.addLayout(row)
        return box

    def _build_scan_mode_group(self):
        box = QGroupBox("Scan Mode")
        layout = QVBoxLayout(box)
        self.project_radio = QRadioButton("Project Structure (target subfolders)")
        self.project_radio.setToolTip(
            "Only scans subfolders matching the target names configured in "
            "Settings → Scan Settings (default: open, confidential).")
        self.all_radio = QRadioButton("All PDFs (Recursive)")
        self.all_radio.setToolTip(
            "Ignores subfolder naming entirely and scans every PDF found under the "
            "selected folder and all of its subfolders.")
        if self.settings.get("scan_mode", "project") == "all":
            self.all_radio.setChecked(True)
        else:
            self.project_radio.setChecked(True)
        # Fixed (item 6): the radio read the saved scan_mode on startup but
        # never wrote it back, so switching to "All PDFs" and relaunching
        # silently reverted to "Project Structure" every time.
        self.project_radio.toggled.connect(self._on_scan_mode_toggled)
        self.all_radio.toggled.connect(self._on_scan_mode_toggled)
        layout.addWidget(self.project_radio)
        layout.addWidget(self.all_radio)

        # Restored: a live hint showing exactly which subfolder names
        # "Project Structure" mode will currently match, since that's
        # configurable in Settings and otherwise invisible from this panel.
        self.target_subfolders_hint = QLabel(self._target_subfolders_hint_text())
        self.target_subfolders_hint.setWordWrap(True)
        self.target_subfolders_hint.setStyleSheet("color:#95a5a6; font-size:8pt; font-style:italic;")
        layout.addWidget(self.target_subfolders_hint)
        return box

    def _on_scan_mode_toggled(self, _checked=None):
        self.settings.set("scan_mode", "all" if self.all_radio.isChecked() else "project")

    def _target_subfolders_hint_text(self):
        targets = ", ".join(self.engine.get_target_subfolder_names())
        alias_note = " (+ common variations)" if self.settings.get("use_folder_aliases", True) else ""
        return f"Project mode targets: {targets}{alias_note}. Edit in Settings → Scan Settings."

    def _build_scan_options_group(self):
        box = QGroupBox("Scan Options")
        layout = QVBoxLayout(box)
        self.check_cannot_open = QCheckBox("Cannot Open (Corrupt)")
        self.check_cannot_open.setChecked(self.settings.get("check_cannot_open", True))
        self.check_cannot_open.setToolTip(
            "Flags PDFs that fail to open or report zero pages — likely corrupted files.")
        self.check_password = QCheckBox("Password Protected")
        self.check_password.setChecked(self.settings.get("check_password_protected", True))
        self.check_password.setToolTip(
            "Flags PDFs that require a password to open. Reported separately from 'Cannot "
            "Open' so it's clear which files just need a password versus which are truly broken.")
        self.check_not_clear = QCheckBox("Not Clear (Low Resolution)")
        self.check_not_clear.setChecked(self.settings.get("check_not_clear", True))
        not_clear_tip = "Flags scanned pages with embedded images below the DPI threshold below."
        if not FITZ_AVAILABLE:
            not_clear_tip += " (Disabled: requires PyMuPDF.)"
            self.check_not_clear.setText("Not Clear (Low Resolution) — disabled, requires PyMuPDF")
        self.check_not_clear.setToolTip(not_clear_tip)
        self.check_missing_info = QCheckBox("Missing Information")
        self.check_missing_info.setChecked(self.settings.get("check_missing_info", True))
        self.check_missing_info.setToolTip(
            "Flags PDFs with mostly blank pages or lacking content.")
        self.check_duplicates = QCheckBox("Duplicate Detection")
        self.check_duplicates.setChecked(self.settings.get("check_duplicates", True))
        self.check_duplicates.setToolTip(
            "Hashes file contents to flag the same PDF appearing in more than one scanned "
            "location (e.g. both 'open' and 'confidential') — its own kind of compliance issue.")
        for cb in (self.check_cannot_open, self.check_password, self.check_not_clear,
                  self.check_missing_info, self.check_duplicates):
            layout.addWidget(cb)

        toggle_row = QHBoxLayout()
        select_all = QPushButton("Select all")
        select_all.setObjectName("secondaryButton")
        select_none = QPushButton("Select none")
        select_none.setObjectName("secondaryButton")
        select_all.clicked.connect(lambda: self._set_all_checks(True))
        select_none.clicked.connect(lambda: self._set_all_checks(False))
        toggle_row.addWidget(select_all)
        toggle_row.addWidget(select_none)
        layout.addLayout(toggle_row)

        dpi_row = QHBoxLayout()
        dpi_row.addWidget(QLabel("Resolution Threshold:"))
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(72, 600)
        self.dpi_spin.setValue(self.settings.get("resolution_threshold", 150))
        self.dpi_spin.setToolTip(
            "Pages with embedded images below this DPI are flagged as low resolution.")
        dpi_row.addWidget(self.dpi_spin)
        dpi_row.addWidget(QLabel("DPI"))
        layout.addLayout(dpi_row)
        return box

    def _set_all_checks(self, value):
        for cb in (self.check_cannot_open, self.check_password, self.check_not_clear,
                  self.check_missing_info, self.check_duplicates):
            cb.setChecked(value)

    def _build_action_buttons(self):
        box = QWidget()
        layout = QVBoxLayout(box)
        self.scan_btn = QPushButton("Start Scan  (Enter)")
        self.scan_btn.setObjectName("scanButton")
        self.scan_btn.clicked.connect(self.start_scan)
        self.cancel_btn = QPushButton("Cancel Scan  (Esc)")
        self.cancel_btn.setObjectName("cancelButton")
        self.cancel_btn.clicked.connect(self.cancel_scan)
        self.cancel_btn.setEnabled(False)
        self.export_btn = QPushButton("Export to Word  (Ctrl+E)")
        self.export_btn.clicked.connect(self.export_to_word)
        self.export_btn.setEnabled(False)
        clear_btn = QPushButton("Clear Results")
        clear_btn.setObjectName("secondaryButton")
        clear_btn.clicked.connect(self.clear_results)
        for b in (self.scan_btn, self.cancel_btn, self.export_btn, clear_btn):
            layout.addWidget(b)
        return box

    def _build_progress_group(self):
        box = QGroupBox("Progress")
        layout = QVBoxLayout(box)
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        status_row = QHBoxLayout()
        self.status_label = QLabel("Ready to scan")
        self.status_label.setStyleSheet("color:#27ae60;")
        self.current_file_label = QLabel("")
        self.current_file_label.setStyleSheet("color:#7f8c8d; font-size:9pt;")
        status_row.addWidget(self.status_label)
        status_row.addStretch()
        status_row.addWidget(self.current_file_label)
        layout.addLayout(status_row)
        self.eta_label = QLabel("")
        self.eta_label.setStyleSheet("color:#7f8c8d; font-size:9pt;")
        layout.addWidget(self.eta_label)
        return box

    def _build_results_group(self):
        box = QGroupBox("Scan Results")
        layout = QVBoxLayout(box)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Filter:"))
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filter by filename, folder, or error type...")
        self.filter_edit.textChanged.connect(self._apply_filter)
        filter_row.addWidget(self.filter_edit)
        clear_filter_btn = QPushButton("Clear")
        clear_filter_btn.setObjectName("secondaryButton")
        clear_filter_btn.clicked.connect(lambda: self.filter_edit.setText(""))
        filter_row.addWidget(clear_filter_btn)
        layout.addLayout(filter_row)

        self.proxy_model = QSortFilterProxyModel()
        self.proxy_model.setSourceModel(self.results_model)
        self.proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy_model.setFilterKeyColumn(-1)  # search all columns

        self.results_table = QTableView()
        self.results_table.setModel(self.proxy_model)
        self.results_table.setSortingEnabled(True)
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.doubleClicked.connect(self._on_result_double_clicked)
        self.results_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._on_result_context_menu)

        # Empty-state overlay: without this the panel is just a blank grid
        # before any scan, after a clean scan, or when a filter excludes
        # everything — which reads as broken rather than "nothing to show
        # yet". A QStackedWidget swaps between the real table and this
        # message depending on row counts, updated from _update_empty_state.
        self.empty_state_label = QLabel("No results yet — select a folder and start a scan.")
        self.empty_state_label.setAlignment(Qt.AlignCenter)
        self.empty_state_label.setStyleSheet("color:#95a5a6; font-size:10pt;")
        self.empty_state_label.setWordWrap(True)

        self.results_stack = QStackedWidget()
        self.results_stack.addWidget(self.results_table)
        self.results_stack.addWidget(self.empty_state_label)
        layout.addWidget(self.results_stack)

        self.results_model.rowsInserted.connect(self._update_empty_state)
        self.results_model.rowsRemoved.connect(self._update_empty_state)
        self.results_model.modelReset.connect(self._update_empty_state)
        self.proxy_model.rowsInserted.connect(self._update_empty_state)
        self.proxy_model.rowsRemoved.connect(self._update_empty_state)
        self._update_empty_state()
        return box

    def _update_empty_state(self, *args):
        """Decides which of the three empty-state messages (if any) to
        show, and switches the stacked widget accordingly. Called whenever
        rows are added/removed on either the source or filtered model."""
        total = self.results_model.rowCount()
        visible = self.proxy_model.rowCount()

        if total == 0:
            message = ("No errors found — every scanned PDF looks good."
                       if self.has_scanned else
                       "No results yet — select a folder and start a scan.")
            self.empty_state_label.setText(message)
            self.results_stack.setCurrentWidget(self.empty_state_label)
        elif visible == 0:
            self.empty_state_label.setText("No results match your filter.")
            self.results_stack.setCurrentWidget(self.empty_state_label)
        else:
            self.results_stack.setCurrentWidget(self.results_table)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence(Qt.Key_Return), self, activated=self._shortcut_start_scan)
        QShortcut(QKeySequence(Qt.Key_Enter), self, activated=self._shortcut_start_scan)
        QShortcut(QKeySequence(Qt.Key_Escape), self, activated=self._shortcut_cancel_scan)
        QShortcut(QKeySequence("Ctrl+E"), self, activated=self._shortcut_export)

    def _shortcut_start_scan(self):
        # Don't hijack Enter while focus is in a text field.
        focus_widget = QApplication.focusWidget()
        if isinstance(focus_widget, (QLineEdit,)):
            return
        if self.scan_btn.isEnabled():
            self.start_scan()

    def _shortcut_cancel_scan(self):
        if self.cancel_btn.isEnabled():
            self.cancel_scan()

    def _shortcut_export(self):
        if self.export_btn.isEnabled():
            self.export_to_word()

    # ---- results table interactions ---------------------------------------

    def _apply_filter(self, text):
        self.proxy_model.setFilterFixedString(text)

    def _selected_result(self):
        indexes = self.results_table.selectionModel().selectedRows()
        if not indexes:
            return None
        source_index = self.proxy_model.mapToSource(indexes[0])
        return self.results_model.result_at_source_row(source_index.row())

    def _on_result_double_clicked(self, index):
        source_index = self.proxy_model.mapToSource(index)
        result = self.results_model.result_at_source_row(source_index.row())
        self._open_file(result.path)

    def _on_result_context_menu(self, pos):
        index = self.results_table.indexAt(pos)
        if not index.isValid():
            return
        self.results_table.selectRow(index.row())
        result = self._selected_result()
        if not result:
            return
        menu = QMenu(self)
        menu.addAction("Open File", lambda: self._open_file(result.path))
        menu.addAction("Show in Folder", lambda: self._reveal_in_file_manager(result.path))
        menu.addSeparator()
        menu.addAction("Copy Path", lambda: QApplication.clipboard().setText(result.path))
        menu.exec(self.results_table.viewport().mapToGlobal(pos))

    def _open_file(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found", f"This file no longer exists:\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open file:\n{e}")

    def _reveal_in_file_manager(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found", f"This file no longer exists:\n{path}")
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.run(["explorer", f'/select,"{os.path.normpath(path)}"'], check=False)
            elif sys.platform == "darwin":
                subprocess.run(["open", "-R", path], check=False)
            else:
                subprocess.run(["xdg-open", os.path.dirname(path)], check=False)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open containing folder:\n{e}")

    # ---- settings / history -------------------------------------------------

    def open_settings(self):
        dialog = SettingsDialog(self, self.settings, self.cache)
        if dialog.exec():
            self.dpi_spin.setValue(self.settings.get("resolution_threshold", 150))
            self.target_subfolders_hint.setText(self._target_subfolders_hint_text())
            self.resize(self.settings.get("window_width", 1200), self.settings.get("window_height", 850))

    def open_history(self):
        HistoryDialog(self, self.history).exec()

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", self.settings.get("last_folder", ""))
        if folder:
            self.folder_edit.setText(folder)
            self.settings.set("last_folder", folder)

    # ---- scan lifecycle -------------------------------------------------

    def clear_results(self):
        self.filter_edit.setText("")
        self.has_scanned = False
        self.results_model.clear()
        self.summary_label.setText("No scan performed yet.")
        self.progress_bar.setValue(0)
        self.status_label.setText("Ready to scan")
        self.status_label.setStyleSheet("color:#27ae60;")
        self.current_file_label.setText("")
        self.eta_label.setText("")
        self.export_btn.setEnabled(False)
        self.scan_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

    def _check_flags(self):
        return {
            "cannot_open": self.check_cannot_open.isChecked(),
            "password_protected": self.check_password.isChecked(),
            "not_clear": self.check_not_clear.isChecked(),
            "missing_info": self.check_missing_info.isChecked(),
            "duplicates": self.check_duplicates.isChecked(),
        }

    def start_scan(self):
        folder = self.folder_edit.text()
        if not folder:
            QMessageBox.critical(self, "Error", "Please select a parent folder first!")
            return
        if not os.path.exists(folder):
            QMessageBox.critical(self, "Error", "Selected folder does not exist!")
            return

        scan_mode = "all" if self.all_radio.isChecked() else "project"
        if scan_mode == "all":
            msg = f"This will recursively scan every PDF found under:\n{folder}\n\nContinue?"
        else:
            valid_folders, _ = self.engine.find_valid_folders(folder, scan_mode)
            if not valid_folders:
                targets = ", ".join(self.engine.get_target_subfolder_names())
                QMessageBox.critical(self, "Error",
                    f"No subfolders matching your configured targets ({targets}) were found!\n\n"
                    "Adjust target subfolder names in Settings, or switch to 'All PDFs' mode.")
                return
            msg = f"Found {len(valid_folders)} valid project folders.\n\nContinue?"

        if QMessageBox.question(self, "Confirm Scan", msg) != QMessageBox.Yes:
            return

        self.clear_results()
        self.has_scanned = True
        self._cancel_event.set()  # should_continue() == True
        self._set_controls_enabled_during_scan(False)

        self.thread = QThread()
        self.worker = ScanWorker(self.engine, folder, scan_mode, self._check_flags(), self.dpi_spin.value())
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.result_found.connect(self.results_model.add_result)
        self.worker.result_updated.connect(self.results_model.update_result)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self._on_scan_finished)
        self.worker.failed.connect(self._on_scan_failed)
        self.thread.start()

    def cancel_scan(self):
        self._cancel_event.clear()  # should_continue() == False
        self.status_label.setText("Cancelling...")
        self.status_label.setStyleSheet("color:#e74c3c;")
        self.cancel_btn.setEnabled(False)

    def _set_controls_enabled_during_scan(self, enabled):
        self.scan_btn.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)
        self.export_btn.setEnabled(False)
        self.browse_btn.setEnabled(enabled)
        self.settings_btn.setEnabled(enabled)
        self.history_btn.setEnabled(enabled)

    def _on_progress(self, progress):
        self.progress_bar.setMaximum(progress.total)
        self.progress_bar.setValue(progress.index)
        self.status_label.setText(f"Scanning... ({progress.index}/{progress.total})")
        self.current_file_label.setText(f"{progress.parent}/{progress.filename[:30]}...")
        elapsed_str = self._format_duration(progress.elapsed_seconds)
        eta_str = self._format_duration(progress.eta_seconds) if progress.eta_seconds is not None else "calculating..."
        self.eta_label.setText(f"Elapsed: {elapsed_str}  |  ETA: {eta_str}")

    @staticmethod
    def _format_duration(total_seconds):
        total_seconds = max(0, int(total_seconds))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"{minutes}:{seconds:02d}"

    def _on_scan_failed(self, error_message):
        """Fixed (Bug 3): previously an unhandled exception in the scan
        worker meant `finished` never fired, so the UI stayed permanently
        locked (Scan/Cancel/Export/Browse/Settings/History all disabled,
        progress bar frozen) until the whole app was restarted. This
        handler runs the same cleanup _on_scan_finished would have, plus
        tells the user what happened."""
        self._cleanup_thread("thread", "worker")
        self.status_label.setText("Scan failed")
        self.status_label.setStyleSheet("color:#e74c3c;")
        self.summary_label.setText(f"Scan failed: {error_message}")
        self.summary_label.setStyleSheet("color:#e74c3c; padding:8px;")
        self._set_controls_enabled_during_scan(True)
        self._results_for_export = self.results_model.all_results()
        self.export_btn.setEnabled(self.results_model.rowCount() > 0)
        logger.error(f"Scan failed: {error_message}")
        QMessageBox.critical(self, "Scan Failed",
                             f"The scan stopped unexpectedly:\n{error_message}\n\n"
                             "Any results found before the error are still shown below.")

    def _on_scan_finished(self, summary, results):
        # See _cleanup_thread's docstring for why this quit()+wait() is
        # both safe (brief) and necessary here.
        self._cleanup_thread("thread", "worker")

        status_word = "cancelled" if summary.cancelled else "completed"
        text = (f"Scan {status_word}: {summary.total_checked} PDFs in {summary.unique_parents} folders | "
               f"Errors: {summary.total_errors} | Time: {summary.duration_seconds:.1f}s")
        if summary.cache_hits:
            text += f" | Skipped {summary.cache_hits} unchanged"
        if summary.timeouts:
            text += f" | {summary.timeouts} timed out"

        color = "#27ae60" if not summary.cancelled else "#e67e22"
        self.summary_label.setText(text)
        self.summary_label.setStyleSheet(f"color:{color}; padding:8px;")
        self.status_label.setText(f"Scan {status_word}")
        self.status_label.setStyleSheet(f"color:{color};")
        self.current_file_label.setText("")
        self.eta_label.setText(f"Total time: {self._format_duration(summary.duration_seconds)}")
        # A clean scan (zero errors) never triggers a rowsInserted event to
        # refresh the empty-state message on its own, since nothing was
        # added — this makes sure "every scanned PDF looks good" actually
        # replaces the "select a folder and start a scan" message.
        self._update_empty_state()

        self.history.add({
            "timestamp": __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "folder": self.folder_edit.text(),
            "scan_mode": "all" if self.all_radio.isChecked() else "project",
            "files_checked": summary.total_checked,
            "errors_found": summary.total_errors,
            "duration_seconds": round(summary.duration_seconds, 1),
            "cancelled": summary.cancelled,
        })

        self._set_controls_enabled_during_scan(True)
        self.export_btn.setEnabled(bool(results))
        self._results_for_export = results

        if summary.cancelled:
            self._notify("Scan Cancelled",
                        f"Stopped after checking {summary.total_checked} PDFs — "
                        f"{summary.total_errors} error(s) found so far.")
        elif summary.total_errors:
            self._notify("Scan Complete", f"Found {summary.total_errors} PDFs with errors "
                        f"in {summary.duration_seconds:.1f}s")
        else:
            self._notify("Scan Complete", f"No errors found in {summary.total_checked} PDFs!")

    def _notify(self, title, message):
        # Native OS notification via Qt's system tray API — no third-party
        # notification library needed the way the Tkinter version required
        # `plyer` as an optional dependency.
        if self.tray_icon and self.tray_icon.isSystemTrayAvailable():
            self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 6000)
        else:
            QMessageBox.information(self, title, message)

    def closeEvent(self, event):
        # If a scan is mid-flight when the window is closed, signal it to
        # stop and wait for the thread to actually exit before allowing the
        # close to proceed — same reasoning as the fix in _on_scan_finished.
        self._cancel_event.clear()
        for thread_attr in ("thread", "_update_thread", "_download_thread"):
            thread = getattr(self, thread_attr, None)
            if thread is not None and thread.isRunning():
                thread.quit()
                thread.wait(5000)

        # Restored (item 7): the window size the user actually ends up
        # with — whether from Settings or just dragging the corner — is
        # saved here so it's remembered next launch, instead of always
        # starting at the same size regardless of what was set before.
        self.settings.set("window_width", self.width())
        self.settings.set("window_height", self.height())
        self.settings.save(self.settings.settings)
        event.accept()

    def _cleanup_thread(self, thread_attr, worker_attr, timeout_ms=3000):
        """Quit-and-wait a finished worker's thread, then clear both
        references. Safe to call from a slot handling that worker's
        finished/result signal: by the time such a slot runs, the worker's
        run() method has already returned (the emit was its last action),
        so this wait is normally near-instant. It's bounded rather than
        unconditional, though — blocking forever on thread shutdown would
        trade one failure mode (a destroyed-while-running crash) for a
        worse one (an app that can never close). If the thread doesn't
        stop within the timeout, we log it and move on; the thread object
        is intentionally NOT cleared in that case, so a later cleanup
        attempt (e.g. on close) can still find and wait on it again."""
        thread = getattr(self, thread_attr, None)
        if thread is None:
            return
        thread.quit()
        if thread.wait(timeout_ms):
            setattr(self, thread_attr, None)
            setattr(self, worker_attr, None)
        else:
            logger.warning(f"{thread_attr} did not stop within {timeout_ms}ms of quit()")

    def export_to_word(self):
        results = getattr(self, "_results_for_export", [])
        if not results:
            QMessageBox.warning(self, "No Results", "No scan results to export!")
            return
        if not DOCX_AVAILABLE:
            QMessageBox.critical(self, "Missing Dependency",
                                 "python-docx is required. Install with: pip install python-docx")
            return
        doc_path, _ = QFileDialog.getSaveFileName(self, "Save Word Report",
                                                   self.settings.get("last_folder", ""),
                                                   "Word Document (*.docx)")
        if not doc_path:
            return
        try:
            export_results_to_word(doc_path, results, self.folder_edit.text(),
                                   self.dpi_spin.value(), self._check_flags())
        except Exception as e:
            logger.error(f"Failed to create Word document: {e}")
            QMessageBox.critical(self, "Error", f"Failed to create Word document:\n{e}")
            return

        # Restored (item 4): auto_open_word_report existed in
        # DEFAULT_SETTINGS from the start but nothing in the Qt UI ever
        # read it — export always just showed a plain success box with no
        # way to open the report or remember a preference either way.
        auto_open_pref = self.settings.get("auto_open_word_report", "ask")
        if auto_open_pref == "ask":
            if self._prompt_open_report():
                self._open_file(doc_path)
        else:
            QMessageBox.information(self, "Success", f"Report saved to:\n{doc_path}")
            if auto_open_pref == "always":
                self._open_file(doc_path)

    def _prompt_open_report(self):
        """A plain QMessageBox can't host a checkbox, so this is a small
        custom dialog offering to open the just-saved report, with a
        'Don't ask again' option that locks in the choice via
        auto_open_word_report ('always'/'never')."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Report Saved")
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("Report saved successfully.\nWould you like to open it now?"))

        dont_ask_cb = QCheckBox("Don't ask again")
        layout.addWidget(dont_ask_cb)

        result = {"open": False}

        def choose(should_open):
            result["open"] = should_open
            if dont_ask_cb.isChecked():
                self.settings.set("auto_open_word_report", "always" if should_open else "never")
                self.settings.save(self.settings.settings)
            dialog.accept()

        btn_row = QHBoxLayout()
        not_now_btn = QPushButton("Not Now")
        not_now_btn.setObjectName("secondaryButton")
        not_now_btn.clicked.connect(lambda: choose(False))
        open_btn = QPushButton("Open Report")
        open_btn.clicked.connect(lambda: choose(True))
        btn_row.addWidget(not_now_btn)
        btn_row.addStretch()
        btn_row.addWidget(open_btn)
        layout.addLayout(btn_row)

        dialog.exec()
        return result["open"]

    # ---- auto-updater -------------------------------------------------

    def _check_updates_async(self):
        if not self.settings.get("auto_check_updates", True):
            return
        self._update_thread = QThread()
        self._update_worker = UpdateCheckWorker()
        self._update_worker.moveToThread(self._update_thread)
        self._update_thread.started.connect(self._update_worker.run)
        self._update_worker.result_ready.connect(self._handle_update_result)
        self._update_thread.start()

    def _handle_update_result(self, result):
        self._cleanup_thread("_update_thread", "_update_worker")

        has_update, latest, data = result
        if not (has_update and data):
            return

        release_url = data.get("html_url", "")
        release_notes = data.get("body") or "No release notes available."
        asset = UpdateChecker.find_installable_asset(data)

        dialog = QDialog(self)
        dialog.setWindowTitle("Update Available")
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"<b>A new version ({latest}) is available.</b>"))
        layout.addWidget(QLabel(f"You have {CURRENT_VERSION}."))
        notes = QLabel(release_notes[:500])
        notes.setWordWrap(True)
        notes.setStyleSheet("background:#ffffff; padding:10px; border-radius:8px;")
        layout.addWidget(notes)

        if not asset:
            hint = QLabel("No auto-installable update is available for this build — "
                         "you'll need to download it yourself.")
            hint.setWordWrap(True)
            hint.setStyleSheet("color:#95a5a6; font-style:italic;")
            layout.addWidget(hint)

        btn_row = QHBoxLayout()
        later_btn = QPushButton("Later")
        later_btn.setObjectName("secondaryButton")
        later_btn.clicked.connect(dialog.reject)
        btn_row.addWidget(later_btn)
        btn_row.addStretch()

        if release_url:
            view_btn = QPushButton("View Release Page")
            view_btn.setObjectName("secondaryButton")
            view_btn.clicked.connect(lambda: __import__("webbrowser").open(release_url))
            btn_row.addWidget(view_btn)

        if asset:
            install_btn = QPushButton("Download && Install")
            install_btn.setObjectName("scanButton")
            install_btn.clicked.connect(lambda: (dialog.accept(), self._start_update_download(latest, asset)))
            btn_row.addWidget(install_btn)
        elif release_url:
            open_btn = QPushButton("Open Release Page")
            open_btn.clicked.connect(lambda: (dialog.accept(), __import__("webbrowser").open(release_url)))
            btn_row.addWidget(open_btn)

        layout.addLayout(btn_row)
        dialog.exec()

    def _start_update_download(self, latest_version, asset):
        download_url = asset.get("browser_download_url")
        if not download_url:
            QMessageBox.critical(self, "Update Failed", "The release asset has no download URL.")
            return

        import tempfile
        dest_path = Path(tempfile.gettempdir()) / asset.get("name", f"pdf_checker_update_{latest_version}")
        size_hint = asset.get("size") or 0

        progress = QProgressDialog(f"Downloading version {latest_version}...", None, 0, 100, self)
        progress.setWindowTitle("Downloading Update")
        progress.setWindowModality(Qt.WindowModal)
        progress.setCancelButton(None)  # no cancelling mid-download, matching the Tkinter version
        progress.setMinimumDuration(0)
        if not size_hint:
            # Fixed: previously always a determinate 0-100 bar, which sat
            # frozen at 0% with no visual sign of activity whenever the
            # live response had no Content-Length (download_asset never
            # calls on_progress in that case). The release's own metadata
            # size (known before the download even starts) is used first;
            # only fall back to an indeterminate/marquee style when that's
            # unavailable too.
            progress.setRange(0, 0)
        else:
            progress.setValue(0)
        progress.show()

        self._download_thread = QThread()
        self._download_worker = DownloadWorker(download_url, str(dest_path), size_hint=size_hint)
        self._download_worker.moveToThread(self._download_thread)
        self._download_thread.started.connect(self._download_worker.run)
        self._download_worker.progress.connect(lambda frac: progress.setValue(int(frac * 100)))
        self._download_worker.finished.connect(lambda path: self._on_update_downloaded(progress, path, latest_version))
        self._download_worker.failed.connect(lambda err: self._on_update_download_failed(progress, err))
        self._download_thread.start()

    def _on_update_download_failed(self, progress, error_text):
        self._cleanup_thread("_download_thread", "_download_worker")
        progress.close()
        QMessageBox.critical(self, "Update Failed",
                             f"Couldn't download the update automatically:\n{error_text}\n\n"
                             "You can still download it manually from the release page.")

    def _on_update_downloaded(self, progress, downloaded_path, latest_version):
        self._cleanup_thread("_download_thread", "_download_worker")
        progress.close()
        if QMessageBox.question(
            self, "Ready to Install",
            f"Version {latest_version} has been downloaded.\n\n"
            "The app will close and restart automatically to finish installing. "
            "Save any work first.\n\nInstall now?"
        ) != QMessageBox.Yes:
            return
        try:
            if getattr(sys, "frozen", False):
                apply_windows_exe_update(downloaded_path)
            else:
                apply_source_update(downloaded_path, current_script_path=__file__)
        except Exception as e:
            logger.error(f"Failed to apply update: {e}")
            QMessageBox.critical(self, "Update Failed", f"Couldn't install the update:\n{e}")
            return
        QApplication.quit()


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(MODERN_QSS)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
