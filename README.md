# PDF Error Checker Pro

A professional desktop application for scanning and validating PDF files across multiple project folders.

## Features

- 🔍 **Multi-folder scanning** - Recursively finds all folders with 'open' and 'confidential' subfolders
- ⚡ **Fast processing** - Background threading keeps UI responsive
- 📊 **Detailed reports** - Export professional Word documents with error summaries
- ⚙️ **Configurable** - Professional settings dialog with persistent configuration
- 🎨 **Modern UI** - Clean, professional interface with custom styling

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

## Usage

1. Click **Browse** to select a parent folder
2. The app will automatically find all subfolders containing both 'open' and 'confidential' folders
3. Configure scan options (resolution threshold, error types to check)
4. Click **Start Scan**
5. Review results and export to Word if needed

## Settings

Access settings via the ⚙️ button in the header:

- **General**: Window size, theme (coming soon)
- **Scan Settings**: Resolution threshold, page limits, detection thresholds
- **Export**: Auto-export options, format preferences
- **About**: Version info and dependency status

## Requirements

- Python 3.8+
- PyMuPDF (fitz)
- PyPDF2
- pdfminer.six
- python-docx (for Word export)

## License

MIT License
