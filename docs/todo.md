# Project Management: Task Breakdown

## Phase 1: Setup & Core Processing (COMPLETED)

- [x] **Task 1.1: Project Setup**  
  Create directory structure (`src`, `docs`, `tests`, `output`, `input`), create virtual environment, add `requirements.txt`.

- [x] **Task 1.2: CLI Argument Parsing**  
  Implement argument parsing (`argparse`) for:
  1. Input folder by default input folder created above
  2. Output folder by default output folder created above
  3. User Excel path  by default docs folder created above
  4. *(Optional)* ICH section Excel path by default docs/iche3_categories.xlsx

- [x] **Task 1.3: RTF File Discovery**  
  Implement a function to scan the input directory and identify all `.rtf` files.

- [x] **Task 1.4: RTF Parsing Module**  
  - Create a function/class to accept an RTF path  
  - Use `striprtf` to extract text  
  - Implement title extraction logic
  - Implement output-name generation
  - Return `(output_name, title, text_content)` with error handling

- [x] **Task 1.5: RTF → PDF Conversion Module**  
  - Choose and implement the conversion method
  - Create a function/class to accept RTF path, output PDF path, and title  
  - Perform conversion and add a PDF bookmark using the title
  - Save to the `_pdf` directory with error handling

- [x] **Task 1.6: Initial Data Collection**  
  Iterate through RTF files, call parsing (Task 1.4) and conversion (Task 1.5) modules, and populate the data structures.

---

## Phase 2: Data Merging & TOC Generation (COMPLETED)

- [x] **Task 2.1: Excel Data Loading Module**  
  - Create function(s) to load user Excel and ICH section Excel into pandas DataFrames  
  - Merge them on section number → `ich_number_categories`  
  - Add error handling for file I/O and parsing

- [x] **Task 2.2: Data Merging & Reporting**  
  - Merge document data with ICH section data on output name  
  - Identify matches and mismatches  
  - Create the sorted list/DataFrame for TOC generation
  - Write mismatches report if needed

- [x] **Task 2.3: TOC Generation Module**  
  - Use ReportLab/FPDF to generate TOC PDF
  - Add section headers (`[section number] [ICH section name]`)  
  - List document titles under each section with calculated page numbers
  - Add bookmarks for section headers within the TOC PDF

---

## Phase 3: PDF Assembly & Finalization (COMPLETED)

- [x] **Task 3.1: PDF Merging Module**  
  - Use PyPDF/PyMuPDF to merge PDFs from `_pdf/` into combined PDF in the TOC order  
  - Preserve individual document bookmarks

- [x] **Task 3.2: TOC Integration**  
  Concatenate TOC PDF and combined document PDF into final output PDF

- [x] **Task 3.3: Page Number Calculation & Update Module**  
  1. Determine TOC length
  2. Build page number mapping for documents
  3. Update page numbers in TOC with correct offsets
  4. Update bookmarks to correct final pages

- [x] **Task 3.4: Final Output Saving**  
  Save the final PDF in the specified output folder and clean up temporary files/folders.

---

## Phase 4: Refinement & Testing (Remaining Tasks)

- [ ] **Task 4.1: Error Handling & Logging Improvements**  
  Review modules, enhance error handling, and optimize logging for better traceability and user feedback.

- [ ] **Task 4.2: Performance Optimization**  
  Identify and address performance bottlenecks, especially for large document sets.

- [ ] **Task 4.3: Testing & QA**  
  - Create more comprehensive test cases for varied input scenarios
  - Test with large document sets
  - Test with uncommon RTF formatting
  - Verify consistent bookmark and TOC behavior across different PDF viewers

- [ ] **Task 4.4: Documentation**  
  - Update README.md with detailed installation and usage instructions
  - Add troubleshooting section for common issues
  - Document the internal architecture and data flow
  - Add inline code documentation where needed

- [ ] **Task 4.5: User Experience Improvements**  
  - Add progress indicators for long-running operations
  - Improve error messages and warnings to be more user-friendly
  - Consider adding a simple GUI wrapper (optional)
