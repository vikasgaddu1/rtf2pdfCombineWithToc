# Building Executable Files

This document explains how to create standalone executable files (.exe) from the RTF2PDF Python project.

## Prerequisites

1. **Python 3.8+** installed on your system
2. **pip** package manager
3. All project dependencies installed

## Quick Build

### Option 1: Using Batch Script (Recommended for Windows)
```cmd
build.bat
```

### Option 2: Using PowerShell Script
```powershell
.\build.ps1
```

## Manual Build Process

If you prefer to build manually or the scripts don't work:

### 1. Install Dependencies
```cmd
pip install -r requirements.txt
```

### 2. Build GUI Version
```cmd
pyinstaller --clean build_gui.spec
```

### 3. Build CLI Version
```cmd
pyinstaller --clean build_cli.spec
```

## Output Files

After successful build, you'll find:
- **GUI executable**: `dist\RTF2PDF_GUI.exe`
- **CLI executable**: `dist\RTF2PDF_CLI.exe`

## Distribution

### For End Users
1. Copy the executable file to the target system
2. Copy the `docs` folder (if your application uses it) alongside the executable
3. Ensure the target system has Microsoft Word installed (for RTF processing)

### Folder Structure for Distribution
```
MyApp/
├── RTF2PDF_GUI.exe
├── docs/                    # Include if needed
│   ├── filename_section_mapping.xlsx
│   └── ich_categories.xlsx
├── input/                   # Create empty folders for user
└── output/                  # Create empty folders for user
```

## Troubleshooting

### Common Issues

1. **Missing Dependencies Error**
   - Ensure all packages in `requirements.txt` are installed
   - Try: `pip install --upgrade -r requirements.txt`

2. **Build Fails with Import Error**
   - Check if all `hiddenimports` are listed in the .spec files
   - Add missing imports to the respective .spec file

3. **Executable Runs but Features Don't Work**
   - Ensure all data files are included in the `datas` section of .spec files
   - Copy required folders (like `docs`) alongside the executable

4. **Large Executable Size**
   - The executables will be large (50-200MB) due to included libraries
   - This is normal for PyInstaller builds with scientific libraries

### Advanced Configuration

To customize the build, edit the `.spec` files:

- **build_gui.spec**: Configuration for GUI version
- **build_cli.spec**: Configuration for CLI version

Key sections to modify:
- `datas`: Include additional files/folders
- `hiddenimports`: Add missing Python modules
- `icon`: Set custom application icon
- `console`: Show/hide console window

## Build Environment

- **Tested on**: Windows 10/11
- **Python**: 3.8+
- **PyInstaller**: 5.10.0+

## Notes

- The first build may take several minutes
- Subsequent builds are faster due to caching
- Build artifacts are excluded from git via `.gitignore`
- Both GUI and CLI versions are built by default 