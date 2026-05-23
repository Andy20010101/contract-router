@echo off
setlocal

cd /d "%~dp0"

py -3.12 --version >nul 2>&1
if errorlevel 1 (
    echo Python 3.12 was not found. Please install Python 3.12 and enable the py launcher.
    exit /b 1
)

py -3.12 -m venv .venv
if errorlevel 1 exit /b 1

call .venv\Scripts\activate.bat
if errorlevel 1 exit /b 1

python -m pip install --upgrade pip
if errorlevel 1 exit /b 1

pip install -r requirements-build.txt
if errorlevel 1 exit /b 1

python -m PyInstaller --clean --noconfirm file_classifier.spec
if errorlevel 1 exit /b 1

echo.
echo Build complete: dist\contract-router.exe
