@echo off
echo ========================================
echo Building PDF Error Checker Pro EXE (Qt Edition)
echo ========================================
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
)

REM Sanity check: the two application files must be present and in the
REM same folder as this script, since pdf_checker_qt.py imports directly
REM from pdf_checker_core.py.
if not exist "pdf_checker_qt.py" (
    echo ERROR: pdf_checker_qt.py not found in this folder.
    echo Run this script from the same folder as pdf_checker_qt.py and pdf_checker_core.py.
    pause
    exit /b 1
)
if not exist "pdf_checker_core.py" (
    echo ERROR: pdf_checker_core.py not found in this folder.
    echo pdf_checker_qt.py imports from it directly — both files must sit together.
    pause
    exit /b 1
)

echo.
echo Building executable...
echo.

REM icon.ico/icon.png are optional. If present, both the .exe's own icon
REM (--icon) and the in-app icon lookup (--add-data, read via sys._MEIPASS
REM at runtime — see APP_DIR in pdf_checker_qt.py) are wired up. If neither
REM file exists, the app still builds and runs fine with no custom icon.
set ICON_FLAG=
set ADD_DATA_FLAGS=
if exist "icon.ico" (
    set ICON_FLAG=--icon=icon.ico
    set ADD_DATA_FLAGS=%ADD_DATA_FLAGS% --add-data "icon.ico;."
)
if exist "icon.png" (
    set ADD_DATA_FLAGS=%ADD_DATA_FLAGS% --add-data "icon.png;."
)

REM --onefile: single portable .exe (matches the previous Tkinter build).
REM --windowed: no console window behind the GUI.
REM pdf_checker_core.py needs no explicit --add-data — PyInstaller's static
REM analysis follows the "from pdf_checker_core import ..." statement and
REM bundles it as compiled code automatically. The --hidden-import flags
REM below ARE needed though: PyMuPDF, pdfminer.six, and python-docx all
REM load some of their own submodules dynamically in ways PyInstaller's
REM static analyzer can miss, especially inside try/except ImportError
REM blocks like this app's optional-dependency pattern. These specific
REM names matched what the original Tkinter build's .spec file required.
pyinstaller --clean --noconfirm --onefile --windowed ^
    --name "PDF Error Checker Pro" ^
    --hidden-import=fitz ^
    --hidden-import=pymupdf ^
    --hidden-import=pdfminer ^
    --hidden-import=pdfminer.high_level ^
    --hidden-import=PyPDF2 ^
    --hidden-import=docx ^
    %ICON_FLAG% %ADD_DATA_FLAGS% ^
    pdf_checker_qt.py

echo.
echo ========================================
if exist "dist\PDF Error Checker Pro.exe" (
    echo SUCCESS! EXE created in dist\ folder
    echo.
    echo NOTE: settings, cache, history, and the log file are created next
    echo to the .exe itself the first time it runs — not inside the dist
    echo folder's build artifacts, and not lost between runs.
    echo ========================================
) else (
    echo FAILED! Check error messages above.
    echo.
    echo Common causes: PySide6 not installed for the Python environment
    echo running this script, or a missing Qt plugin at runtime — if the
    echo built exe launches but shows a blank/broken window, try adding
    echo --collect-all PySide6 to the pyinstaller command above.
    echo ========================================
)
echo.
pause
