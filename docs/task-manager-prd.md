# Product Requirements Document (PRD): RTF to Combined PDF with TOC Tool

## 1. Introduction

### Purpose  
To automate the process of converting a collection of RTF documents into a single PDF file, preceded by a generated Table of Contents (TOC) based on document titles and predefined section categorizations.

### Scope  
The tool will:  
- Read RTF files from a specified folder  
- Process and merge them with section information from external Excel files  
- Generate a TOC  
- Combine the individual PDFs  
- Update the TOC with correct page numbers  
- Save the final consolidated PDF  

It will operate as a Command-Line Interface (CLI) tool.

### Goals  
- Significantly reduce manual effort in compiling reports or documents from multiple RTF sources  
- Ensure consistent formatting and structure for the combined PDF  
- Provide an accurate TOC with bookmarks for easy navigation  
- Handle potential errors gracefully and report discrepancies  

### Target Audience  
Users who need to regularly combine multiple RTF reports/documents (e.g., regulatory submissions, project documentation) into a single, navigable PDF.

---

## 2. Functional Requirements

### FR1: Input Processing  
The tool must accept the following inputs:  
1. Path to the input folder containing RTF files  
2. Path to the user-provided Excel file mapping output_name to section_number  
3. Path to the output folder for the final PDF and any reports  
4. *(Optional)* Path to the ICH section mapping Excel file (if not assumed to be in a fixed location)

### FR2: RTF Parsing  
For each `.rtf` file in the input folder:  
- Read the file content using appropriate error handling
- Extract plain text using the `striprtf` library  
- Determine the document title  
  - Use the first non-empty line of the extracted text as the title  
- Determine the output_name  
  - Use the RTF filename without the `.rtf` extension 

### FR3: RTF → PDF Conversion  
- Convert each RTF file into an individual PDF using Microsoft Word COM automation
- Save each PDF into a `_pdf` subfolder within the main output directory  
- **Dependency:** Requires Microsoft Word for RTF-to-PDF conversion via COM automation

### FR4: Data Management & Merging  
1. Build an in-memory structure mapping output_name → title for all RTFs  
2. Read the user-provided Excel file (output_name, section_number)  
3. Read the ICH section mapping Excel file (section_number, ICH_section_name)  
4. Merge the two Excel data sources on section_number  
5. Merge the RTF title data with the combined Excel data on output_name  
6. Retain only records present in both sources for TOC generation  
7. Generate a mismatch report listing output_names present in only one source

### FR5: TOC Generation  
- Generate a separate TOC PDF using FPDF
- Order TOC by section_number, then by RTF order within each section 
- For each section, add a header with section number and ICH section name
- Under each section, list each document's title with page numbers
- Store TOC entry information for later link creation

### FR6: PDF Combination  
1. Combine all individual PDFs into an intermediate combined PDF following TOC order
2. Create a page map tracking the start page of each document
3. Generate the TOC with accurate page numbers using the page map
4. Prepend the TOC to the combined content PDF

### FR7: Bookmarks & Navigation  
- Set PDF bookmarks for both TOC section headers and document entries
- Ensure bookmarks link to the correct pages in the combined document
- Add navigational links from TOC entries to the corresponding content pages

### FR8: Output Generation  
- Save the final PDF (with updated TOC & bookmarks) in the output folder  
- Clean up temporary files and intermediate PDFs
- Provide detailed logging throughout the process

---

## 3. Non-Functional Requirements

### Modular Design
- Code is organized into separate modules with clear responsibilities:
  - `rtf_parser.py`: Handles RTF file reading and title extraction
  - `rtf_converter.py`: Manages RTF to PDF conversion via Word COM
  - `data_processing.py`: Handles data loading, merging, and validation
  - `pdf_utils.py`: Provides PDF generation, combination, and bookmarking

### Reliability
- Robust error handling with appropriate try/except blocks
- Comprehensive logging of operations, warnings, and errors
- Graceful handling of partial failures (e.g., if some RTFs fail to convert)

### Performance
- Efficient PDF processing using PyMuPDF (fitz) and pypdf libraries
- Two-pass approach for TOC generation to ensure accurate page numbers

### Maintainability
- Clear code structure with well-named functions and variables
- Consistent logging at appropriate levels
- Type hints for improved code readability and IDE support

### Dependencies
- **Python Libraries:**
  - `pandas`: For data manipulation and Excel file processing
  - `striprtf`: For extracting text from RTF files
  - `fpdf`: For TOC PDF generation
  - `pypdf`: For PDF merging and manipulation
  - `PyMuPDF` (fitz): For handling PDF bookmarks
  - `pywin32`: For Windows COM automation (on Windows only)
- **External Dependencies:**
  - Microsoft Word: Required for RTF-to-PDF conversion
  - Windows OS: For Word COM automation
---

