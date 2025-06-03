# RTF to PDF Converter with TOC

A Python application that converts RTF files to PDF and combines them into a single PDF document with a table of contents and bookmarks.

## Features

- Convert multiple RTF files to PDF
- Automatic or manual section organization
- Generate table of contents with page numbers
- Create PDF bookmarks for easy navigation
- Modern GUI interface with progress tracking
- Configurable PDF settings
- Real-time conversion progress
- Detailed logging
- Support for both automatic and manual section organization
- All settings configured through the GUI (no configuration files)

## Requirements

- Windows OS (uses Word COM automation)
- Python 3.6 or higher
- Microsoft Word installed
- Required Python packages (install via `pip install -r requirements.txt`):
  - pandas
  - fpdf2
  - pypdf
  - PyMuPDF (fitz)
  - pywin32

## Installation

1. Clone the repository:
```bash
git clone https://github.com/vikasgaddu1/rtf2pdfCombineWithToc.git
```

2. Create and activate a virtual environment (recommended):
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

1. Launch the GUI:
```bash
python run_gui.py
```

2. Configure the settings in the GUI:
   - **Input Folder**: Select the folder containing RTF files
   - **Output Folder**: Choose where to save the generated PDFs
   - **Output Filename**: Name for the final combined PDF
   - **Section Settings**:
     - Automatic: Files are organized based on filename prefixes (t, f, l)
     - Manual: Use an Excel file to define sections
   - **PDF Options**:
     - Page Width (mm): Default 210 (A4)
     - Margin (mm): Default 15
     - Font Size: Default 8
     - Header Font Size: Default 10

3. Click "Process Files" to start the conversion

### File Organization

#### Automatic Section Mode
Files are automatically organized based on their filename prefixes:
- `t*`: Tables (Section 1)
- `f*`: Figures (Section 2)
- `l*`: Listings (Section 3)

#### Manual Section Mode
Use an Excel file (`filename_section.xlsx`) with the following columns:
- `filename`: RTF filename (without extension)
- `section_number`: Section number (e.g., "14.1")
- `ICH_section_name`: Section name

### Output

The application generates:
1. Individual PDF files in the `_pdf` subfolder
2. A combined PDF with table of contents and bookmarks
3. Log output in the GUI showing conversion progress

## Project Structure

```
rtf2pdfcombine/
├── run_gui.py              # GUI launcher
├── main.py                 # Main processing logic
├── requirements.txt        # Python dependencies
├── input/                  # Input RTF files
├── output/                 # Generated PDFs
│   └── _pdf/              # Individual PDFs
├── docs/                   # Documentation and mapping files
└── src/
    ├── gui.py             # GUI implementation
    ├── gui_config.py      # GUI configuration holder
    ├── rtf_converter.py   # RTF to PDF conversion
    ├── rtf_parser.py      # RTF title extraction
    ├── data_processing.py # Data handling and validation
    └── pdf_utils.py       # PDF generation utilities
```

## GUI Settings

All application settings are configured through the GUI interface:

- **Input/Output Paths**: Browse and select folders directly
- **Section Mode**: Toggle between automatic and manual modes
- **PDF Settings**: Adjust page layout and font sizes
- **No Configuration Files**: All settings are passed directly from the GUI to the processing engine

## Troubleshooting

1. **Word COM Automation Issues**
   - Ensure Microsoft Word is installed
   - Close any open Word instances
   - Run the application with administrator privileges
   - Check if Word is properly registered in the system

2. **File Conversion Failures**
   - Check file permissions
   - Verify RTF files are not corrupted
   - Ensure filenames follow the required format
   - Check if the output directory is writable

3. **Section Mapping Issues**
   - Verify Excel file format
   - Check section numbers match ICH categories
   - Ensure filenames match exactly
   - Validate Excel file column names

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

[Your License Here]

## Acknowledgments

- Microsoft Word COM automation
- FPDF2 for PDF generation
- PyMuPDF for PDF manipulation
