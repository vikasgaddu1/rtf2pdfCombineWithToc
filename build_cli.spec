# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Define the analysis for the CLI version
a = Analysis(
    ['src/cli.py'],  # Changed from main.py to src/cli.py
    pathex=[],
    binaries=[],
    datas=[
        ('docs', 'docs'),  # Include docs folder if it contains required files
    ],
    hiddenimports=[
        # Windows COM
        'win32com.client',
        'pythoncom',
        'pywintypes',
        # Data processing
        'pandas',
        'pandas.core',
        'pandas.io',
        'pandas.io.excel',
        'xlrd',
        'xlrd.biffh',
        'openpyxl',
        'openpyxl.reader.excel',
        # Text processing
        'striprtf.striprtf',
        # PDF processing
        'fpdf',  # fpdf2 is imported as fpdf
        'pypdf',
        'pypdf.generic',
        'fitz',  # PyMuPDF is imported as fitz
        # Image processing
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        # Report generation
        'reportlab',
        'reportlab.lib',
        'reportlab.lib.colors',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.platypus',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        # Standard library
        'dataclasses',
        'threading',
        'queue',
        'gc',
        'json',
        'pathlib',
        'logging',
        'os',
        'sys',
        'argparse'  # Added for CLI argument parsing
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='RTF2PDF_CLI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Keep console window for CLI
    disable_windowed_traceback=False,
    icon=None,  # Add path to .ico file if you have one
) 