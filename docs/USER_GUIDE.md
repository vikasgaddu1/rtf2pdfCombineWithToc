---
title: RTF to PDF Converter with Table of Contents
description: A tool that converts RTF files to a single PDF document with table of contents and bookmarks
css:
  - https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css
  - https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css
js:
  - https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js
---

# RTF to PDF Converter with Table of Contents
## User Guide

This tool converts RTF files to a single PDF document with an automatically generated table of contents and bookmarks. It is available in both GUI and command-line versions.

## Table of Contents
1. [Installation](#installation)
2. [Quick Start](#quick-start)
3. [Using the GUI Version](#using-the-gui-version)
4. [Using the Command-Line Version](#using-the-command-line-version)
5. [Input Requirements](#input-requirements)
6. [Output Structure](#output-structure)
7. [Troubleshooting](#troubleshooting)

## Installation

The tool is provided as standalone executables in the `dist` folder:
- `RTF2PDF_GUI.exe` - Graphical user interface version
- `RTF2PDF_CLI.exe` - Command-line interface version

No additional installation is required. Simply run the executable of your choice.

## Quick Start

1. Place your RTF files in the `input` folder
2. Run either `RTF2PDF_GUI.exe` or `RTF2PDF_CLI.exe`
3. The combined PDF will be generated in the `output` folder

## Using the GUI Version

1. Launch `RTF2PDF_GUI.exe`
2. The interface will show:
   - Input folder selection
   - Output folder selection
   - Section mapping options:
     - Option to use automatic section organization
     - Option to use a section mapping Excel file
     - Section mapping file selection (if enabled)
   - Progress bar
3. Click "Start Processing" to begin the conversion
4. The tool will:
   - Convert RTF files to PDF
   - Generate a table of contents
   - Combine all PDFs into a single document
   - Add bookmarks for easy navigation

## Using the Command-Line Version

Run `RTF2PDF_CLI.exe` from the command line with the following options:

### Basic Options
- `--input-folder PATH`: Folder containing RTF files (default: ./input)
- `--output-folder PATH`: Folder for output files (default: ./output)
- `--output-filename NAME`: Name of the final combined PDF (default: final_document_with_toc.pdf)

### Section Mapping Options
- `--use-section-file`: Use a section mapping Excel file instead of automatic section organization
- `--section-file PATH`: Path to the section mapping Excel file (required if --use-section-file is set)

### PDF Settings
- `--page-width MM`: Page width in millimeters (default: 210.0)
- `--margin MM`: Page margin in millimeters (default: 15.0)
- `--font-size SIZE`: Base font size (default: 8.0)
- `--header-font-size SIZE`: Header font size (default: 10.0)

### Logging Options
- `--log-level LEVEL`: Set the logging level (choices: DEBUG, INFO, WARNING, ERROR, CRITICAL, default: INFO)

### Examples

1. Basic usage with default settings:
```bash
RTF2PDF_CLI.exe
```

2. Custom input and output folders:
```bash
RTF2PDF_CLI.exe --input-folder "C:\My RTF Files" --output-folder "C:\Output"
```

3. Using a section mapping file:
```bash
RTF2PDF_CLI.exe --use-section-file --section-file "docs\filename_section.xlsx"
```

4. Custom PDF settings:
```bash
RTF2PDF_CLI.exe --page-width 297 --margin 20 --font-size 10 --header-font-size 12
```

5. Full example with all options:
```bash
RTF2PDF_CLI.exe --input-folder "C:\My RTF Files" --output-folder "C:\Output" --output-filename "final.pdf" --use-section-file --section-file "docs\filename_section.xlsx" --page-width 297 --margin 20 --font-size 10 --header-font-size 12 --log-level DEBUG
```

## Input Requirements

### File Structure
- Place all RTF files in the `input` folder
- Files should be properly named according to your section structure
- Supported file format: RTF (Rich Text Format)

### Section Mapping (Optional)
Both GUI and CLI versions support section mapping. You can choose between:

1. **Automatic Section Organization**:
   - The tool will automatically organize sections based on file names
   - Files starting with 't' or 'f' will be placed in section 14.x
   - Files starting with 'l' will be placed in section 16.x

2. **Manual Section Mapping**:
   - Create a section mapping Excel file
   - Place it in the `docs` folder
   - In GUI: Enable the section mapping option and select the file
   - In CLI: Use the `--use-section-file` and `--section-file` options

### Section Mapping Excel File Structure

The section mapping Excel file should contain the following columns:

1. **filename**: The base name of your RTF file (without extension)
2. **section_number**: The section number in the format "X.Y" (e.g., "14.1", "14.3")

The file should be placed in the `docs` folder and named `filename_section.xlsx`.

#### Example Section Mapping

| filename | section_number |
|----------|----------------|
| tsidm01  | 14.1          |
| tsids01  | 14.3          |
| fefos01a | 14.2          |

#### File Naming Convention
- Your RTF files should be named exactly as specified in the `filename` column
- Example: If `filename` is "tsidm01", your file should be named "tsidm01.rtf"
- The tool will match your RTF files with the mapping based on these names

#### Section Numbering Rules
- Use numbers in the format "X.Y" (e.g., "14.1", "14.3")
- The numbers correspond to ICH categories (e.g., 14.x for tables)
- The tool will automatically sort sections based on these numbers
- Files starting with 't' should have section numbers starting with '14'
- Files starting with 'f' should have section numbers starting with '14'
- Files starting with 'l' should have section numbers starting with '16'

#### ICH Categories
The tool uses an additional file `iche3_categories.xlsx` in the `docs` folder that maps section numbers to ICH category names. For example:
- 14.1 might map to "Demographic Data"
- 14.3 might map to "Clinical Study Reports"

## Output Structure

The tool generates the following in the `output` folder:
1. `final_document_with_toc.pdf` - The final combined PDF document
2. `file_mismatch_report.txt` (if applicable) - Report of any file mapping issues
3. Intermediate files (automatically cleaned up)

The final PDF includes:
- Table of contents with page numbers
- Bookmarks for easy navigation
- All RTF files converted to PDF and combined in the correct order

## Troubleshooting

1. **Word COM Automation Issues**
   - Ensure Microsoft Word is installed
   - Close any open Word instances
   - Run the application with administrator privileges
   - Check if Word is properly registered in the system

2. **File Mapping Issues**
   - Check the `file_mismatch_report.txt` for details
   - Ensure RTF filenames match exactly with the mapping file
   - Verify section numbers are in the correct format

3. **PDF Generation Issues**
   - Check if the input RTF files are valid
   - Ensure there is enough disk space
   - Try running with `--log-level DEBUG` for more detailed error messages

## System Requirements

- Windows operating system
- Sufficient disk space for PDF generation
- Write permissions in the output directory