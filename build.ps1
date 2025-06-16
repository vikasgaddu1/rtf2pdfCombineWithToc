# RTF2PDF Build Script (PowerShell)
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RTF2PDF Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Check if virtual environment exists
if (-not (Test-Path "venv")) {
    Write-Host "ERROR: Virtual environment 'venv' not found!" -ForegroundColor Red
    Write-Host "Please create a virtual environment first:" -ForegroundColor Yellow
    Write-Host "  python -m venv venv" -ForegroundColor White
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
try {
    & ".\venv\Scripts\Activate.ps1"
    Write-Host "Virtual environment activated successfully" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to activate virtual environment" -ForegroundColor Red
    Write-Host "You may need to run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Verify we're using the right Python
$pythonPath = Get-Command python | Select-Object -ExpandProperty Source
Write-Host "Using Python from: $pythonPath" -ForegroundColor Green

# Check if Python is available from venv
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Found Python: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not available in virtual environment" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if pip is available from venv
try {
    $pipVersion = pip --version 2>&1
    Write-Host "Found pip: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: pip is not available in virtual environment" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`nInstalling/updating dependencies in virtual environment..." -ForegroundColor Yellow
try {
    pip install -r requirements.txt
    Write-Host "Dependencies installed successfully" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to install dependencies" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Building GUI version..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
try {
    pyinstaller --clean build_gui.spec
    Write-Host "GUI build completed successfully" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to build GUI version" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Building CLI version..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
try {
    pyinstaller --clean build_cli.spec
    Write-Host "CLI build completed successfully" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to build CLI version" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Build completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "GUI executable: dist\RTF2PDF_GUI.exe" -ForegroundColor White
Write-Host "CLI executable: dist\RTF2PDF_CLI.exe" -ForegroundColor White
Write-Host "`nNote: Make sure to copy the 'docs' folder alongside" -ForegroundColor Yellow
Write-Host "the executable if your application requires it." -ForegroundColor Yellow

# Deactivate virtual environment
Write-Host "`nDeactivating virtual environment..." -ForegroundColor Yellow
deactivate

Read-Host "`nPress Enter to exit" 