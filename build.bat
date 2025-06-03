@echo off
echo ========================================
echo RTF2PDF Build Script
echo ========================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check if pip is available
pip --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: pip is not installed or not in PATH
    pause
    exit /b 1
)

echo Installing/updating dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ========================================
echo Building GUI version...
echo ========================================
pyinstaller --clean build_gui.spec
if errorlevel 1 (
    echo ERROR: Failed to build GUI version
    pause
    exit /b 1
)

echo.
echo ========================================
echo Building CLI version...
echo ========================================
pyinstaller --clean build_cli.spec
if errorlevel 1 (
    echo ERROR: Failed to build CLI version
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo GUI executable: dist\RTF2PDF_GUI.exe
echo CLI executable: dist\RTF2PDF_CLI.exe
echo.
echo Note: Make sure to copy the 'docs' folder alongside
echo the executable if your application requires it.
echo.
pause 