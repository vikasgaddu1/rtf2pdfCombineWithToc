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
The tool must accept the following inputs via CLI arguments:  
1. Path to the input folder containing RTF files  
2. Path to the user-provided Excel file mapping output_name to section_number  
3. Path to the output folder for the final PDF and any reports  
4. *(Optional)* Path to the ICH section mapping Excel file (if not assumed to be in a fixed location)

### FR2: RTF Parsing  
For each `.rtf` file in the input folder:  
- Read the file content  
- Extract plain text using the `striprtf` library  
- Determine the document title  
  - Use the first non-empty line of the extracted text as the title  
- Determine the output_name  
  - Use the RTF filename without the `.rtf` extension 

### FR3: RTF → PDF Conversion  
- Convert each RTF file into an individual PDF  
- Save each PDF into a `_pdf` subfolder within the main output directory  
- Embed a PDF bookmark in each individual PDF using the extracted title  
- **Dependency:** Requires Microsoft Word for RTF-to-PDF conversion via COM automation or similar interface

### FR4: Data Management & Merging  
1. Build an in-memory structure (e.g., `pandas.DataFrame rtf_title`) mapping output_name → title for all RTFs  
2. Read the user-provided Excel file (output_name, section_number)  
3. Read the ICH section mapping Excel file (section_number, ICH_section_name) stored in docs folder.  
4. Merge the two Excel data sources on section_number → `ich_number_categories`  
5. Merge `rtf_title` with `ich_number_categories` on output_name  
6. Retain only records present in both sources for TOC generation  
7. Generate a mismatch report (`mismatches.csv`) listing output_names present in only one source

### FR5: TOC Generation  
- Generate a separate TOC PDF (`toc_temp.pdf`)  
- Order TOC by section_number, then by RTF order within each section 
- For each section, add a header:  
- Under each section, list each document’s title  
- Use placeholder page numbers (e.g., “XX”) initially  
- Add PDF bookmarks in the TOC pointing to each section header

### FR6: PDF Combination  
1. Merge all individual PDFs from `_pdf/` into `combined_pdf_temp.pdf` following TOC order  
2. Concatenate `toc_temp.pdf` + `combined_pdf_temp.pdf` → `final_temp.pdf`

### FR7: TOC Finalization & Bookmarking  
- Calculate actual page numbers for each TOC entry in `final_temp.pdf` (account for TOC length)  
- Replace placeholder page numbers in the TOC  
- Update section-header bookmarks to point to the corresponding section header on toc page in the final document  
- Ensure individual document bookmarks from FR3 are preserved and functional

### FR8: Output Generation  
- Save the final PDF (with updated TOC & bookmarks) as `Combined_Document.pdf` in the output folder  
- Save the mismatch report (`mismatches.csv`) in the output folder  
- Clean up temporary files (`_pdf/`, `*_temp.pdf`), unless a debug flag is set

---

## 3. Non-Functional Requirements

- **Reliability:** Robust error handling (try/except) for file I/O, parsing, conversion, merging; clear error logs  
- **Maintainability:** Modular code (separate functions/classes); use OOP where beneficial; adhere to PEP 8  
- **Usability:** Clear CLI instructions (`--help`); informative status messages during processing  
- **Dependencies:** List all external dependencies in `requirements.txt` (e.g., `pandas`, `striprtf`, `pypdf`/`PyPDF2`, `reportlab`; or external tools like `unoconv`)
- **Python PDF libraries:** FPDF and Fitz will be used.
---

