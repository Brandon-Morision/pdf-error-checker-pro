@echo off
echo ========================================
echo Building PDF Error Checker Pro EXE
echo ========================================
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
)

echo.
echo Building executable...
echo.

REM Build with PyInstaller
pyinstaller --clean "pdf checker pro.spec"

echo.
echo ========================================
if exist "dist\PDF_Error_Checker_Pro.exe" (
    echo SUCCESS! EXE created in dist\ folder
    echo ========================================
) else (
    echo FAILED! Check error messages above
    echo ========================================
)
echo.
pause
