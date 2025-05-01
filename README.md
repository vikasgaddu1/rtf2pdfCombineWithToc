# RTF2CombinePDFwithTOC

A utility for converting RTF files to PDF, generating a table of contents, and combining them into a single hyperlinked PDF document with bookmarks.

## Features

- Converts RTF files to PDFs
- Extracts titles from RTF files automatically
- Maps files to their respective sections using configuration files
- Generates a Table of Contents (TOC) based on document structure
- Combines individual PDFs into a single document
- Adds hyperlinks from TOC entries to corresponding content pages
- Creates hierarchical bookmarks for easy navigation

## Requirements

- Python 3.8 or higher
- Dependencies (see requirements.txt):
  - pandas
  - xlrd
  - striprtf (for extracting text from RTF)
  - fpdf2
  - pypdf
  - PyMuPDF (fitz)
  - win32com (for Windows users, to handle Word processes)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/RTF2CombinePDFwithTOC.git
   cd RTF2CombinePDFwithTOC
   ```

2. Create a virtual environment (recommended):
   ```
   python -m venv venv
   # On Windows
   venv\Scripts\activate
   # On macOS/Linux
   source venv/bin/activate
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Directory Structure

The utility expects the following directory structure:

```
rtf2pdfcombine/
├── main.py             # Main script
├── input/              # Place your RTF files here
├── output/             # Final and intermediate PDFs will be created here
│   └── _pdf/           # Individual converted PDFs (intermediate)
├── docs/               # Configuration files
│   ├── filename_section.xlsx  # Maps filenames to sections
│   └── iche3_categories.xlsx  # Defines ICH section categories
└── src/                # Source code modules
    ├── rtf_parser.py   # Extracts titles from RTF files
    ├── data_processing.py  # Handles data merging and validation
    └── pdf_utils.py    # PDF generation utilities
```

## Configuration Files

1. **filename_section.xlsx**: Maps each RTF filename to its section number
   - Should contain at least columns for filename and section_number

2. **iche3_categories.xlsx**: Contains information about ICH section categories
   - Used to create proper section headers in TOC and bookmarks

## Usage

1. Place all your RTF files in the `input` directory
2. Ensure your configuration files are correctly set up in the `docs` directory
3. Run the main script:
   ```
   python main.py
   ```
4. The final PDF with TOC and hyperlinks will be created in the `output` directory as `final_document_with_toc.pdf`

## Processing Steps

1. Scans RTF files and extracts titles
2. Loads mapping files for section organization
3. Merges and validates data
4. Creates TOC data structure
5. Converts RTF files to individual PDFs
6. Combines PDFs and creates bookmarks
7. Generates final TOC with hyperlinks
8. Prepends TOC to content PDF to create final document
9. Cleans up intermediate files

## Customization

To modify the default behavior:

1. Edit paths in `main.py` to change input/output directories
2. Adjust PDF layout constants in `src/pdf_utils.py`
3. Modify TOC title and formatting in `generate_toc_pdf()`

## Troubleshooting

- If hyperlinks don't work, ensure you have the latest version of PyMuPDF installed
- For RTF conversion issues, check that:
  - Your RTF files are properly formatted
  - You have proper access permissions to read/write files
  - On Windows, no Word processes are running (the utility attempts to close them)
- If TOC links aren't working, verify that:
  - The `filepath` column in your data matches the actual file paths
  - The page mapping is correctly generated

## Advanced Usage

For more advanced usage or integration into other workflows, you can import and use the individual functions:

```python
from src.rtf_parser import build_title_dataframe
from src.data_processing import load_filename_section_map, load_ich_categories_map, merge_and_validate, create_toc_structure, convert_all
from src.pdf_utils import combine_pdfs, generate_toc_pdf, prepend_toc_to_pdf

# Your custom implementation here
```

## License

[Your license information here]

## Credits

[Your credits information here] 