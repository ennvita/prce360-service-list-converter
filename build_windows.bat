@echo off
setlocal enabledelayedexpansion

echo ============================================
echo  PRCe360 Service List Converter - Windows build
echo ============================================
echo.

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: python not found on PATH.
    echo Install Python 3.12+ from https://python.org and ensure "Add to PATH" is checked.
    pause
    exit /b 1
)

echo Installing dependencies...
python -m pip install --quiet --upgrade openpyxl python-docx lxml pillow "cx_Freeze>=7.2"
if %errorlevel% neq 0 (
    echo ERROR: pip install failed.
    pause
    exit /b 1
)

echo.
echo Generating application icon...
python -c "from PIL import Image; img = Image.open('brand/NMPRC_logo-1.png').convert('RGBA'); sizes = [(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)]; imgs = [img.resize(s, Image.LANCZOS) for s in sizes]; imgs[0].save('brand/NMPRC_logo.ico', format='ICO', append_images=imgs[1:])"
if %errorlevel% neq 0 (
    echo ERROR: Icon generation failed. Ensure brand\NMPRC_logo-1.png exists.
    pause
    exit /b 1
)

echo.
echo Building PRCe360ServiceListConverter.msi ...
python setup_msi.py bdist_msi
if %errorlevel% neq 0 (
    echo.
    echo Build failed. Check the output above for errors.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Done!
echo  Installer: dist\*.msi
echo.
echo  Run the .msi on any Windows machine to install.
echo  - No Python required on the target machine.
echo  - Per-user install (no admin needed).
echo  - Adds Desktop + Start Menu shortcuts.
echo  - Uninstalls cleanly via Add/Remove Programs.
echo ============================================
pause
