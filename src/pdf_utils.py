#!/usr/bin/env python3
import logging
from pathlib import Path
import pandas as pd
from fpdf import FPDF
from pypdf import PdfWriter, PdfReader
import fitz  # Import PyMuPDF

# Import the GUI configuration
from src.gui_config import GUIConfig

# Configure logging (can be configured globally in main if preferred)
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


# --- Constants for PDF Layout ---
PAGE_WIDTH_MM = 210 # A4 width
MARGIN_MM = 15
CONTENT_WIDTH_MM = PAGE_WIDTH_MM - 2 * MARGIN_MM
LINE_HEIGHT = 6
FONT_SIZE = 8
HEADER_FONT_SIZE = 10
# PLACEHOLDER_PAGE_NUM = "XX" # No longer needed
FONT = 'Arial'

# --------------------------------

# Placeholder: Needs toc_data to include 'filepath' column corresponding to page_map keys
def generate_toc_pdf(toc_data: pd.DataFrame, page_map: dict[str, int], output_path: Path, config: GUIConfig = None) -> tuple[Path | None, int | None]:
    """Generates a PDF file for the Table of Contents with actual page numbers.
    No links are added at this stage - they will be added to the final document.

    Args:
        toc_data: DataFrame with TOC structure (columns: level, text, type, filepath).
        page_map: Dictionary mapping filepath strings (lowercase) to their 1-based starting
                  page number in the content PDF.
        output_path: The path where the final TOC PDF will be saved.
        config: GUIConfig object containing PDF settings (optional, uses defaults if None)

    Returns:
        A tuple containing:
            - The path to the generated TOC PDF if successful, None otherwise.
            - The number of pages in the generated TOC PDF, None otherwise.
    """
    # Use config values or defaults
    if config:
        PAGE_WIDTH_MM = config.page_width_mm
        MARGIN_MM = config.margin_mm
        FONT_SIZE = config.font_size
        HEADER_FONT_SIZE = config.header_font_size
    else:
        PAGE_WIDTH_MM = 210  # A4 width
        MARGIN_MM = 15
        FONT_SIZE = 8
        HEADER_FONT_SIZE = 10
    
    CONTENT_WIDTH_MM = PAGE_WIDTH_MM - 2 * MARGIN_MM
    LINE_HEIGHT = 6
    FONT = 'Arial'
    
    logging.info(f"--- Generating Final Table of Contents PDF to {output_path.name} ---")
    if toc_data.empty:
        logging.warning("TOC data is empty, skipping TOC PDF creation.")
        return None, None
    if 'filepath' not in toc_data.columns:
         logging.error("TOC data DataFrame must include a 'filepath' column.")
         # Consider raising an error instead or handling differently
         return None, None

    # Detect if we're in automatic or manual mode by examining section headers
    is_automatic_mode = False
    if not toc_data.empty:
        # Look for section headers in the TOC data
        headers = toc_data[toc_data['type'] == 'header']
        if not headers.empty:
            # Check if headers follow automatic pattern (e.g., "1  Tables", "2  Figures", "3  Listings")
            first_header = headers.iloc[0]['text']
            # Automatic mode headers start with single digit followed by section name
            if (isinstance(first_header, str) and 
                len(first_header.split()) >= 2 and 
                first_header.split()[0].isdigit() and 
                int(first_header.split()[0]) <= 10):
                is_automatic_mode = True
    
    # Determine TOC title based on mode
    if is_automatic_mode:
        toc_title = "Table of Contents"
        title_align = 'C'  # Center align for automatic mode
        logging.info("Detected automatic mode - using 'Table of Contents' title")
    else:
        toc_title = "14. TABLES, FIGURES AND GRAPHS REFERRED TO BUT NOT INCLUDED IN THE TEXT"
        title_align = 'L'  # Left align for manual mode
        logging.info("Detected manual mode - using ICH-specific title")

    try:
        # --- First Pass: Calculate TOC page count ---
        pdf_calc = FPDF(orientation='P', unit='mm', format='A4')
        pdf_calc.set_auto_page_break(auto=True, margin=MARGIN_MM)
        pdf_calc.set_margins(left=MARGIN_MM, top=MARGIN_MM, right=MARGIN_MM)
        pdf_calc.add_page()
        pdf_calc.set_font(FONT, '', FONT_SIZE)
        pdf_calc.set_text_color(0, 0, 255)  # Set text color to blue
        pdf_calc.set_font_size(12)
        pdf_calc.cell(CONTENT_WIDTH_MM, 10, toc_title, 0, 1, title_align)
        pdf_calc.ln(5)

        for _, row in toc_data.iterrows():
            if row['type'] == 'header':
                pdf_calc.set_font(FONT, 'B', HEADER_FONT_SIZE)
                pdf_calc.ln(LINE_HEIGHT * 0.25)
                pdf_calc.multi_cell(CONTENT_WIDTH_MM, LINE_HEIGHT, str(row['text']), 0, 'L')
                pdf_calc.ln(LINE_HEIGHT * 0.25)
                pdf_calc.set_font(FONT, '', FONT_SIZE)
            elif row['type'] == 'entry':
                # For calculation pass, estimate space needed for wrapped entries
                pdf_calc.set_font(FONT, '', FONT_SIZE)
                pdf_calc.set_text_color(0, 0, 255)
                indent = "  " * (row['level'] - 1)
                formatted_text = indent + str(row['text'])
                
                # Check if text will fit on one line (leaving space for dots and page number)
                text_width = pdf_calc.get_string_width(formatted_text)
                page_num_width = pdf_calc.get_string_width("999")  # Estimate
                reserved_space = page_num_width + 30  # Space for page number and dots
                
                if text_width > (CONTENT_WIDTH_MM - reserved_space):
                    # Text needs wrapping - use multi_cell to calculate height
                    wrap_width = CONTENT_WIDTH_MM - reserved_space
                    pdf_calc.multi_cell(wrap_width, LINE_HEIGHT, formatted_text, 0, 'L')
                    # No need for extra space since dots go on the last line
                    pdf_calc.ln(LINE_HEIGHT * 0.25)
                else:
                    # Single line entry
                    pdf_calc.cell(CONTENT_WIDTH_MM * 0.8, LINE_HEIGHT, formatted_text, 0, 0)
                    pdf_calc.cell(CONTENT_WIDTH_MM * 0.2, LINE_HEIGHT, "999", 0, 1, 'R')
                    pdf_calc.ln(LINE_HEIGHT / 4)

        toc_page_count = pdf_calc.page_no()
        logging.info(f"Calculated TOC will require {toc_page_count} page(s).")

        # --- Second Pass: Generate actual TOC PDF without links ---
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=MARGIN_MM)
        pdf.set_margins(left=MARGIN_MM, top=MARGIN_MM, right=MARGIN_MM)
        pdf.add_page()
        pdf.set_font(FONT, '', FONT_SIZE) # Use standard font
        pdf.set_text_color(0, 0, 255)  # Set text color to blue

        # Add TOC Title
        pdf.set_font_size(12)
        pdf.cell(CONTENT_WIDTH_MM, 10, toc_title, 0, 1, title_align)
        pdf.ln(5)

        # Write TOC Entries without links - we'll add them in the final document
        dot_char = "."
        dot_width = pdf.get_string_width(dot_char)
        max_page_num_str = str(len(page_map) * 10 + toc_page_count) # Estimate max page num width reasonably
        page_num_width = pdf.get_string_width(max_page_num_str) + 1 # Add small buffer

        # Store TOC entry info for later link creation
        toc_entries = []

        for _, row in toc_data.iterrows():
            level = row['level']
            text = str(row['text']) # Ensure text is string
            entry_type = row['type']
            file_path_key = str(row['filepath']) # Ensure key is string, use lowercase 'filepath'

            if entry_type == 'header':
                pdf.set_font(FONT, 'B', HEADER_FONT_SIZE) # Bold for headers
                pdf.set_text_color(0, 0, 0)  # Black color for headers
                pdf.ln(LINE_HEIGHT * 0.25) # Reduced space before header
                pdf.multi_cell(CONTENT_WIDTH_MM, LINE_HEIGHT, text, 0, 'L')
                pdf.ln(LINE_HEIGHT * 0.25) # Reduced space after header
                pdf.set_font(FONT, '', FONT_SIZE) # Reset to normal font
                
                # Store header information for use in bookmark creation
                # Headers don't have target pages in content, but we'll record their position in TOC
                # Clean text before storing
                clean_header_text = clean_text(text)
                
                toc_entries.append({
                    'toc_page': pdf.page_no(),
                    'target_page': None,  # No target for headers
                    'text': clean_header_text,
                    'original_text': text,  # Keep original for debugging
                    'page_num_str': '',
                    'is_header': True,
                    'y_position': pdf.get_y()  # Store y position
                })

            elif entry_type == 'entry':
                pdf.set_font(FONT, '', FONT_SIZE) # Ensure normal font for entries
                pdf.set_text_color(0, 0, 255)  # Blue color for entries
                indent = "  " * (level - 1)
                formatted_text = indent + text

                # Get original page number and calculate final page number
                original_page_num = page_map.get(file_path_key)
                if original_page_num is None:
                    logging.warning(f"File path '{file_path_key}' not found in page map for entry '{text}'. Skipping page number.")
                    final_page_num_str = "N/A"
                    final_page_num = None
                else:
                    final_page_num = original_page_num + toc_page_count
                    final_page_num_str = str(final_page_num)

                # Calculate if text needs wrapping
                text_width = pdf.get_string_width(formatted_text)
                current_page_num_width = pdf.get_string_width(final_page_num_str)
                # Reserve space for page number and some dots
                reserved_space = current_page_num_width + 30  # Increased buffer for dots
                
                # Store the starting position for hyperlink creation
                start_y = pdf.get_y()
                start_page = pdf.page_no()
                
                if text_width > (CONTENT_WIDTH_MM - reserved_space):
                    # Text needs wrapping
                    logging.debug(f"Wrapping long title: {formatted_text[:50]}...")
                    
                    # Store current X position before multi_cell
                    start_x = pdf.get_x()
                    
                    # Use multi_cell for the title with a specific width
                    wrap_width = CONTENT_WIDTH_MM - reserved_space
                    pdf.multi_cell(wrap_width, LINE_HEIGHT, formatted_text, 0, 'L')
                    
                    # Get position after multi_cell
                    after_text_y = pdf.get_y()
                    
                    # For wrapped text, we need to find where the last line ends
                    # Split the text into words and estimate the last line
                    words = formatted_text.split()
                    last_line_text = ""
                    temp_text = ""
                    
                    # Build lines word by word to find what's on the last line
                    for word in words:
                        test_text = temp_text + " " + word if temp_text else word
                        if pdf.get_string_width(test_text) <= wrap_width:
                            temp_text = test_text
                        else:
                            last_line_text = temp_text
                            temp_text = word
                    # Don't forget the remaining text
                    if temp_text:
                        last_line_text = temp_text
                    
                    # Calculate where the last line of text ends
                    last_line_text_width = pdf.get_string_width(last_line_text)
                    
                    # Move to the last line
                    pdf.set_y(after_text_y - LINE_HEIGHT)
                    pdf.set_x(MARGIN_MM)
                    
                    # Add a small gap after the text
                    gap_width = 5  # 5mm gap between text and dots
                    
                    # Position after the last line text plus gap
                    text_end_x = MARGIN_MM + last_line_text_width + gap_width
                    
                    # Calculate space for dots from text end to page number
                    available_for_dots = CONTENT_WIDTH_MM - last_line_text_width - current_page_num_width - gap_width
                    
                    if available_for_dots > 0 and dot_width > 0:
                        num_dots = int(available_for_dots / dot_width)
                        dots = dot_char * num_dots
                        
                        # Position at the end of text plus gap
                        pdf.set_x(text_end_x)
                        # Add dots
                        pdf.cell(available_for_dots, LINE_HEIGHT, dots, 0, 0, 'R')
                    else:
                        # No room for dots, just add page number
                        pdf.set_x(MARGIN_MM + CONTENT_WIDTH_MM - current_page_num_width)
                    
                    # Add page number at the end
                    pdf.cell(current_page_num_width, LINE_HEIGHT, final_page_num_str, 0, 1, 'R')
                    
                    # Store entry info with multi-line flag
                    if final_page_num is not None:
                        clean_formatted_text = clean_text(formatted_text)
                        
                        toc_entries.append({
                            'toc_page': start_page,
                            'target_page': final_page_num,
                            'text': clean_formatted_text,
                            'original_text': formatted_text,
                            'page_num_str': final_page_num_str,
                            'is_header': False,
                            'y_position': start_y,
                            'end_y_position': pdf.get_y(),
                            'is_multiline': True,
                            'first_words': ' '.join(formatted_text.split()[:5])  # Store first 5 words for matching
                        })
                    
                else:
                    # Single line entry - improved logic
                    # Add a small gap between text and dots for better readability
                    gap_width = 3  # 3mm gap
                    
                    # Calculate available space for dots with gap
                    available_dots_width = CONTENT_WIDTH_MM - text_width - current_page_num_width - gap_width

                    dots = ""
                    if available_dots_width > 0 and dot_width > 0:
                        num_dots = int(available_dots_width / dot_width)
                        dots = dot_char * num_dots

                    # Record this entry's info
                    if final_page_num is not None:
                        clean_formatted_text = clean_text(formatted_text)
                        
                        toc_entries.append({
                            'toc_page': pdf.page_no(),
                            'target_page': final_page_num,
                            'text': clean_formatted_text,
                            'original_text': formatted_text,
                            'page_num_str': final_page_num_str,
                            'is_header': False,
                            'y_position': start_y,
                            'end_y_position': start_y + LINE_HEIGHT,
                            'is_multiline': False,
                            'first_words': ' '.join(formatted_text.split()[:5])
                        })

                    # Add cells with gap
                    pdf.cell(text_width, LINE_HEIGHT, formatted_text, 0, 0)
                    pdf.cell(gap_width, LINE_HEIGHT, "", 0, 0)  # Gap between text and dots
                    pdf.cell(available_dots_width, LINE_HEIGHT, dots, 0, 0, 'R')
                    pdf.cell(current_page_num_width, LINE_HEIGHT, final_page_num_str, 0, 1, 'R')
                
                pdf.ln(LINE_HEIGHT / 4) # Keep small space between entries

        # --- Save PDF ---
        output_path.parent.mkdir(parents=True, exist_ok=True) # Ensure output dir exists
        pdf.output(str(output_path), "F")
        logging.info(f"Successfully generated TOC PDF: {output_path} with {len(toc_entries)} entries")
        
        # Create a metadata file with TOC entries for later link creation
        toc_info_path = output_path.with_suffix('.json')
        import json
        with open(toc_info_path, 'w') as f:
            json.dump(toc_entries, f)
        logging.debug(f"Saved TOC entry information to {toc_info_path}")
        
        # Return the actual page count of the generated TOC
        return output_path, pdf.page_no()

    except ImportError:
         logging.error("FPDF library not found. Please install it: pip install fpdf2")
         return None, None
    except KeyError as ke:
        logging.error(f"KeyError during TOC generation, likely missing 'filepath' in toc_data or page_map: {ke}")
        return None, None
    except Exception as toc_err:
        logging.error(f"Failed to generate final TOC PDF: {toc_err}", exc_info=True)
        return None, None


def combine_pdfs(final_df: pd.DataFrame, output_pdf_folder: Path, output_path: Path) -> tuple[Path | None, dict[str, int] | None]:
    """Combines PDF files specified in final_df into a single PDF with bookmarks.

    Args:
        final_df: DataFrame sorted in the desired order, containing at least
                  'filepath' (lowercase, relative path from workspace root or absolute)
                  and 'title' (lowercase) columns.
        output_pdf_folder: Path to the folder containing the individual PDF files
                           (generated from RTFs). The filenames in this folder should
                           match the basename of the 'filepath' in final_df.
        output_path: The path where the combined PDF (without TOC) will be saved.

    Returns:
        A tuple containing:
            - The path to the combined PDF if successful, None otherwise.
            - A dictionary mapping filepath strings (lowercase) to their 1-based starting
              page number in the combined PDF, None otherwise.
    """
    logging.info(f"--- Combining PDFs from {output_pdf_folder.name} into {output_path.name} with bookmarks ---")

    if not {'filepath', 'title'}.issubset(final_df.columns):
        logging.error("final_df DataFrame must include 'filepath' and 'title' columns.")
        return None, None
    if final_df.empty:
        logging.warning("final_df is empty, nothing to combine.")
        return None, None
        
    # Ensure final_df is sorted by section_number then filename_stem (same as TOC and bookmarks)
    if 'section_number' in final_df.columns and 'filename_stem' in final_df.columns:
        logging.info("Sorting PDFs to match TOC and bookmark order...")
        final_df = final_df.sort_values(by=['section_number', 'filename_stem'])
        logging.info(f"Sorted {len(final_df)} files by section_number and filename_stem")

    writer = PdfWriter()
    page_map = {}
    current_page_number = 0 # 0-based index for PyPDF outline/pages

    try:
        for index, row in final_df.iterrows():
            file_path_str = str(row['filepath'])
            pdf_filename = Path(file_path_str).name.replace('.rtf', '.pdf') # Assume conversion replaces ext
            pdf_file_to_combine = output_pdf_folder / pdf_filename
            # Use lowercase 'title' for bookmark
            base_title = str(row['title']) # Bookmark title
            filename_stem = Path(file_path_str).stem # Get stem for uniqueness
            
            # Clean the title text
            base_title = clean_text(base_title)
            bookmark_title = f"{base_title} ({filename_stem})" # Make title unique

            if not pdf_file_to_combine.is_file():
                logging.warning(f"PDF file not found: {pdf_file_to_combine}. Skipping.")
                continue # Or handle as error depending on requirements

            try:
                reader = PdfReader(str(pdf_file_to_combine))
                num_pages = len(reader.pages)
                if num_pages == 0:
                     logging.warning(f"PDF file {pdf_filename} has 0 pages. Skipping.")
                     continue

                # Append the pages from the current PDF
                writer.append(str(pdf_file_to_combine))

                # Store the 1-based starting page number for TOC generation
                # Use the original filepath (lowercase) from the dataframe as the key
                page_map[file_path_str] = current_page_number + 1
                
                # Special logging for FEFOS01A
                if "fefos01a" in file_path_str.lower():
                    logging.info(f"FEFOS01A page mapping: {file_path_str} -> page {current_page_number + 1}")

                # Update the current page number for the next iteration
                current_page_number += num_pages

                logging.debug(f"Appended {pdf_filename} ({num_pages} pages). Current total pages: {current_page_number}.")

            except Exception as append_err:
                logging.error(f"Failed to process or append {pdf_filename}: {append_err}")
                # Decide whether to abort or continue
                # writer.close() # Close writer if aborting
                # return None, None # Abort

        if current_page_number == 0:
            logging.error("No pages were added to the combined PDF. Aborting.")
            writer.close()
            return None, None

        # Write the combined PDF (without TOC)
        output_path.parent.mkdir(parents=True, exist_ok=True) # Ensure output dir exists
        with open(output_path, "wb") as fp:
            writer.write(fp)
        logging.info(f"Successfully combined {len(page_map)} PDFs into {output_path}")
        return output_path, page_map

    except Exception as merge_err:
        logging.error(f"Failed to combine PDFs: {merge_err}", exc_info=True)
        return None, None
    finally:
        writer.close() # Ensure the writer is closed


def prepend_toc_to_pdf(toc_pdf_path: Path, content_pdf_path: Path, final_output_path: Path, final_df: pd.DataFrame, page_map: dict[str, int]) -> Path | None:
    """Merges the TOC PDF and the main content PDF using PyMuPDF (fitz).
    
    This uses a simpler approach to avoid incompatibility issues between links.
    
    Args:
        toc_pdf_path: Path to the generated TOC PDF.
        content_pdf_path: Path to the combined content PDF (intermediate).
        final_output_path: The path for the final output PDF.
        final_df: DataFrame containing the sorted order, 'filepath', and 'title' for bookmarks.
        page_map: Dictionary mapping filepath strings to their 1-based starting page
                  number in the content_pdf (before TOC is prepended).

    Returns:
        The path to the final PDF if successful, None otherwise.
    """
    logging.info(f"--- Prepending TOC ({toc_pdf_path.name}) to Content ({content_pdf_path.name}) ---")

    try:
        # Detect if we're in automatic or manual mode by examining section numbers
        is_automatic_mode = False
        if not final_df.empty and 'section_number' in final_df.columns:
            sample_section_numbers = final_df['section_number'].unique()[:3]
            # If section numbers are simple digits (1, 2, 3) rather than decimal format (14.1, 14.3)
            is_automatic_mode = all(str(sn).isdigit() and int(str(sn)) <= 10 for sn in sample_section_numbers if pd.notna(sn))
        
        mode_name = "Automatic" if is_automatic_mode else "Manual"
        logging.info(f"Detected {mode_name} section mode for hyperlink creation")

        # Use PyPDF to combine PDFs (simpler, more reliable)
        merger = PdfWriter()
        
        # Add TOC PDF
        merger.append(str(toc_pdf_path))
        num_toc_pages = len(PdfReader(str(toc_pdf_path)).pages)
        logging.debug(f"Added TOC PDF with {num_toc_pages} pages")
        
        # Add content PDF
        merger.append(str(content_pdf_path))
        num_content_pages = len(PdfReader(str(content_pdf_path)).pages)
        logging.debug(f"Added content PDF with {num_content_pages} pages")
        
        # Create temporary merged PDF
        temp_merged_path = final_output_path.with_name(f"temp_{final_output_path.name}")
        with open(temp_merged_path, "wb") as f:
            merger.write(f)
        merger.close()
        logging.debug(f"Created temporary merged PDF at {temp_merged_path}")
        
        # Now use PyMuPDF to add links and bookmarks
        doc = fitz.open(str(temp_merged_path))
        
        # Try to load TOC entry information from JSON file
        toc_info_path = toc_pdf_path.with_suffix('.json')
        toc_entries = []
        if toc_info_path.exists():
            import json
            try:
                with open(toc_info_path, 'r') as f:
                    toc_entries = json.load(f)
                logging.debug(f"Loaded {len(toc_entries)} TOC entries from {toc_info_path}")
            except json.JSONDecodeError:
                logging.error(f"Failed to load TOC entry information from {toc_info_path}")
        
        # If we have TOC entries, create links
        if toc_entries:
            logging.debug(f"Creating {len(toc_entries)} links in the TOC using {mode_name} mode logic...")
            # Log some sample page mappings for debugging
            logging.info(f"TOC page count: {num_toc_pages}, Content page count: {num_content_pages}")
            logging.info(f"Sample of page_map entries: {list(page_map.items())[:3]}...")
            
            # Count how many headers and entries we have
            headers = [e for e in toc_entries if e.get('is_header', False)]
            entries = [e for e in toc_entries if not e.get('is_header', False)]
            logging.info(f"TOC contains {len(headers)} section headers and {len(entries)} document entries")
            
            # Initialize variables that will be used later in bookmark generation
            main_title_line = None
            
            # Create hyperlinks using mode-specific logic
            if is_automatic_mode:
                # Automatic mode: sections are "1 Tables", "2 Figures", "3 Listings"
                logging.info("Using automatic mode hyperlink creation logic")
                
                # Find main title line for bookmark generation
                for page_idx in range(min(num_toc_pages, 3)):  # Check first 3 pages max
                    page = doc[page_idx]
                    text_blocks = page.get_text("dict")["blocks"]
                    for block in text_blocks:
                        for line in block.get("lines", []):
                            line_text = "".join(span.get("text", "") for span in line.get("spans", []))
                            line_text_stripped = line_text.strip()
                            
                            # Check for main title
                            if (line_text_stripped.startswith("14. TABLES") or 
                                "TABLES, FIGURES AND GRAPHS" in line_text_stripped or
                                line_text_stripped == "Table of Contents"):
                                main_title_line = {
                                    'page': page_idx,
                                    'rect': fitz.Rect(line["bbox"]),
                                    'text': line_text_stripped
                                }
                                logging.info(f"Identified main title on page {page_idx+1}: '{line_text_stripped}'")
                                break
                        if main_title_line:  # Break out of outer loop too
                            break
                    if main_title_line:  # Break out of page loop
                        break
                
                for entry in toc_entries:
                    # Skip header entries - they don't get hyperlinks in the TOC
                    if entry.get('is_header', False):
                        logging.debug(f"Skipping link creation for header: {entry['text']}")
                        continue
                        
                    # Skip entries with no target page
                    if entry.get('target_page') is None:
                        logging.debug(f"Skipping link creation for entry with no target page: {entry['text']}")
                        continue
                        
                    toc_page_idx = entry['toc_page'] - 1  # Convert 1-based to 0-based
                    target_page_idx = entry['target_page'] - 1  # Convert 1-based to 0-based
                    
                    # Get the page
                    page = doc[toc_page_idx]
                    
                    # Enhanced matching for wrapped titles
                    is_multiline = entry.get('is_multiline', False)
                    page_num_str = entry['page_num_str']
                    first_words = entry.get('first_words', '')
                    
                    # Find the line(s) with this entry
                    text_blocks = page.get_text("dict")["blocks"]
                    entry_found = False
                    entry_rect = None
                    
                    for block in text_blocks:
                        for line in block.get("lines", []):
                            line_text = "".join(span.get("text", "") for span in line.get("spans", []))
                            line_text_stripped = line_text.strip()
                            
                            # Skip section headers in automatic mode
                            if (len(line_text_stripped.split()) >= 2 and
                                line_text_stripped.split()[0].isdigit() and
                                any(word.lower() in ['tables', 'figures', 'listings'] for word in line_text_stripped.split()[1:])):
                                continue
                            
                            # Skip main title
                            if ("TABLES, FIGURES AND GRAPHS" in line_text_stripped or
                                line_text_stripped == "Table of Contents"):
                                continue
                            
                            # For multi-line entries, match by first words or page number
                            if is_multiline:
                                # Check if this line contains the beginning of our entry
                                if first_words and first_words in line_text:
                                    # This is the start of our multi-line entry
                                    entry_rect = fitz.Rect(line["bbox"])
                                    entry_found = True
                                    logging.debug(f"Found start of multi-line entry: '{line_text_stripped[:50]}...'")
                                # Or check if it's a line with just dots and page number
                                elif page_num_str in line_text and ('.' * 3) in line_text and line_text.strip().endswith(page_num_str):
                                    # This is the end line with page number
                                    if entry_rect:
                                        # Expand rect to include this line
                                        entry_rect = entry_rect | fitz.Rect(line["bbox"])
                                    else:
                                        entry_rect = fitz.Rect(line["bbox"])
                                    
                                    # Create hyperlink for the entire multi-line entry
                                    expanded_rect = fitz.Rect(
                                        MARGIN_MM,
                                        entry_rect.y0,
                                        page.rect.width - MARGIN_MM,
                                        entry_rect.y1
                                    )
                                    
                                    page.insert_link({
                                        "kind": fitz.LINK_GOTO,
                                        "from": expanded_rect,
                                        "page": target_page_idx,
                                        "zoom": 0
                                    })
                                    
                                    logging.info(f"Added multi-line link for entry ending with page {page_num_str}")
                                    entry_found = False  # Reset for next entry
                                    entry_rect = None
                                    break
                                # Continue building the rect if we're in a multi-line entry
                                elif entry_found and entry_rect:
                                    entry_rect = entry_rect | fitz.Rect(line["bbox"])
                            else:
                                # Single line entry - use existing logic
                                if (page_num_str in line_text and 
                                    not line_text_stripped.split()[0].isdigit() and
                                    ('.' * 3) in line_text and
                                    line_text.strip().endswith(page_num_str)):
                                    
                                    logging.info(f"Auto mode: Found single-line TOC entry for page {page_num_str}: '{line_text_stripped[:50]}...'")
                                    
                                    # Create hyperlink for the entire line
                                    rect = fitz.Rect(line["bbox"])
                                    expanded_rect = fitz.Rect(
                                        MARGIN_MM,
                                        rect.y0,
                                        page.rect.width - MARGIN_MM,
                                        rect.y1
                                    )
                                    
                                    page.insert_link({
                                        "kind": fitz.LINK_GOTO,
                                        "from": expanded_rect,
                                        "page": target_page_idx,
                                        "zoom": 0
                                    })
                                    
                                    logging.debug(f"Added automatic mode link from TOC page {toc_page_idx+1} to target page {target_page_idx+1}")
                                    break
            else:
                # Manual mode: sections are "14.1 Something", "14.3 Something"
                logging.info("Using manual mode hyperlink creation logic")
                
                # Pre-identify all section header lines for robust detection
                section_header_lines = []
                main_title_line = None
                
                # Scan through all pages in TOC
                for page_idx in range(min(num_toc_pages, 3)):  # Check first 3 pages max
                    page = doc[page_idx]
                    text_blocks = page.get_text("dict")["blocks"]
                    for block in text_blocks:
                        for line in block.get("lines", []):
                            line_text = "".join(span.get("text", "") for span in line.get("spans", []))
                            line_text_stripped = line_text.strip()
                            
                            # Check for main title
                            if (line_text_stripped.startswith("14. TABLES") or 
                                "TABLES, FIGURES AND GRAPHS" in line_text_stripped or
                                line_text_stripped == "Table of Contents"):
                                main_title_line = {
                                    'page': page_idx,
                                    'rect': fitz.Rect(line["bbox"]),
                                    'text': line_text_stripped
                                }
                                logging.info(f"Identified main title on page {page_idx+1}: '{line_text_stripped}'")
                                break
                            
                            # Check for manual mode section header patterns (14.1, 14.3, etc.)
                            if (line_text_stripped and
                                len(line_text_stripped.split()) >= 2 and 
                                any(line_text_stripped.startswith(f"{i}.{j}") for i in range(10, 20) for j in range(1, 10))):
                                section_header_lines.append({
                                    'page': page_idx,
                                    'rect': fitz.Rect(line["bbox"]),
                                    'text': line_text_stripped
                                })
                                logging.info(f"Identified manual mode section header line on page {page_idx+1}: '{line_text_stripped}'")
                
                for entry in toc_entries:
                    # Skip header entries - they don't get hyperlinks in the TOC
                    if entry.get('is_header', False):
                        logging.debug(f"Skipping link creation for header: {entry['text']}")
                        continue
                        
                    # Skip entries with no target page
                    if entry.get('target_page') is None:
                        logging.debug(f"Skipping link creation for entry with no target page: {entry['text']}")
                        continue
                        
                    toc_page_idx = entry['toc_page'] - 1  # Convert 1-based to 0-based
                    target_page_idx = entry['target_page'] - 1  # Convert 1-based to 0-based
                    
                    # Get the page
                    page = doc[toc_page_idx]
                    
                    # Enhanced matching for wrapped titles
                    is_multiline = entry.get('is_multiline', False)
                    page_num_str = entry['page_num_str']
                    first_words = entry.get('first_words', '')
                    
                    # Find the line(s) with this entry
                    text_blocks = page.get_text("dict")["blocks"]
                    entry_found = False
                    entry_rect = None
                    
                    for block in text_blocks:
                        for line in block.get("lines", []):
                            line_text = "".join(span.get("text", "") for span in line.get("spans", []))
                            
                            # Get the rectangle for this line
                            rect = fitz.Rect(line["bbox"])
                            
                            # Check if this line is the main title - never add hyperlinks to it
                            if main_title_line and main_title_line['page'] == toc_page_idx and main_title_line['rect'].intersects(rect):
                                continue
                            
                            # Check if this line is a section header
                            is_section_header = False
                            for header_line in section_header_lines:
                                if header_line['page'] == toc_page_idx and header_line['rect'].intersects(rect):
                                    is_section_header = True
                                    break
                            
                            # Skip section headers
                            if is_section_header:
                                continue
                            
                            # For multi-line entries, match by first words or page number
                            if is_multiline:
                                # Check if this line contains the beginning of our entry
                                if first_words and first_words in line_text:
                                    # This is the start of our multi-line entry
                                    entry_rect = fitz.Rect(line["bbox"])
                                    entry_found = True
                                    logging.debug(f"Found start of multi-line entry: '{line_text.strip()[:50]}...'")
                                # Or check if it's a line with page number at the end
                                elif page_num_str in line_text and line_text.strip().endswith(page_num_str):
                                    # This is the end line with page number
                                    if entry_rect:
                                        # Expand rect to include this line
                                        entry_rect = entry_rect | fitz.Rect(line["bbox"])
                                    else:
                                        entry_rect = fitz.Rect(line["bbox"])
                                    
                                    # Create hyperlink for the entire multi-line entry
                                    expanded_rect = fitz.Rect(
                                        MARGIN_MM,
                                        entry_rect.y0,
                                        page.rect.width - MARGIN_MM,
                                        entry_rect.y1
                                    )
                                    
                                    page.insert_link({
                                        "kind": fitz.LINK_GOTO,
                                        "from": expanded_rect,
                                        "page": target_page_idx,
                                        "zoom": 0
                                    })
                                    
                                    logging.info(f"Added multi-line link for manual mode entry ending with page {page_num_str}")
                                    entry_found = False  # Reset for next entry
                                    entry_rect = None
                                    break
                                # Continue building the rect if we're in a multi-line entry
                                elif entry_found and entry_rect:
                                    entry_rect = entry_rect | fitz.Rect(line["bbox"])
                            else:
                                # Single line entry - use existing logic
                                if page_num_str in line_text and line_text.strip().endswith(page_num_str):
                                    logging.info(f"Manual mode: Found single-line match for entry with page {page_num_str}: '{line_text.strip()[:50]}...'")
                                    
                                    # Create hyperlink for the entire line
                                    expanded_rect = fitz.Rect(
                                        MARGIN_MM,
                                        rect.y0,
                                        page.rect.width - MARGIN_MM,
                                        rect.y1
                                    )
                                    
                                    page.insert_link({
                                        "kind": fitz.LINK_GOTO,
                                        "from": expanded_rect,
                                        "page": target_page_idx,
                                        "zoom": 0
                                    })
                                    
                                    logging.debug(f"Added manual mode link from TOC page {toc_page_idx+1} to target page {target_page_idx+1}")
                                    break
        
        # Generate bookmarks
        final_bookmarks = []
        
        # Add main title as the first bookmark pointing to TOC page 1
        if main_title_line:
            main_title_text = main_title_line['text']
            # Add as level 1 bookmark (PyMuPDF requires first item to be level 1)
            final_bookmarks.append([1, main_title_text, 1])
            logging.info(f"Added main title as top-level bookmark: '{main_title_text}'")
        
        if not final_df.empty and page_map is not None:
            # Create a dictionary to keep track of TOC sections and their positions
            # This will help us create hierarchical bookmarks
            section_to_toc_page = {}
            
            # First, find TOC section header positions
            if toc_entries:
                for entry in toc_entries:
                    if entry.get('is_header', False):
                        # Extract section number from header text (expecting format like "14.1 Demographic Data" for manual mode
                        # or "1 Tables", "2 Figures", "3 Listings" for automatic mode)
                        text_parts = entry['text'].strip().split()
                        section_num = None
                        
                        # Check for manual mode format (e.g., "14.1", "14.3")
                        if text_parts and any(part.startswith(('14.1', '14.3')) for part in text_parts):
                            section_num = text_parts[0]  # Get '14.1' or '14.3' part
                        # Check for automatic mode format (e.g., "1 Tables", "2 Figures", "3 Listings")
                        elif (text_parts and len(text_parts) >= 2 and
                              text_parts[0].isdigit() and
                              any(word.lower() in ['tables', 'figures', 'listings'] for word in text_parts[1:])):
                            section_num = text_parts[0]  # Get '1', '2', or '3' part
                        
                        if section_num:
                            section_to_toc_page[section_num] = entry['toc_page']
                            logging.info(f"Found section header {section_num} on TOC page {entry['toc_page']}")
            
            # Group files by section number
            section_groups = {}
            
            # First pass - collect entries by section
            for index, row in final_df.iterrows():
                try:
                    section_number = row['section_number']
                    section_name = row.get('ICH_section_name', '')
                    
                    # Clean section name
                    section_name = clean_text(section_name)
                    
                    if section_number not in section_groups:
                        section_groups[section_number] = {
                            'title': f"{section_number} {section_name}",
                            'entries': []
                        }
                    
                    filepath_str = str(row['filepath'])
                    base_title = str(row['title'])
                    filename_stem = Path(filepath_str).stem
                    
                    # Clean the title text
                    base_title = clean_text(base_title)
                    bookmark_title = f"{base_title} ({filename_stem})"
                    
                    original_page_num = page_map.get(filepath_str)
                    if original_page_num is not None:
                        # Adjust page number by adding the number of TOC pages (1-based)
                        final_page_num = original_page_num + num_toc_pages
                        
                        # Add to the appropriate section group
                        section_groups[section_number]['entries'].append({
                            'title': bookmark_title,
                            'page': final_page_num,
                            'filename_stem': filename_stem  # Store for sorting
                        })
                except KeyError as e:
                    logging.warning(f"Skipping bookmark due to missing column in final_df: {e}")
            
            # Second pass - build hierarchical bookmarks
            # Sort the section numbers to match TOC order
            sorted_section_numbers = sorted(section_groups.keys())
            
            for section_number in sorted_section_numbers:
                group = section_groups[section_number]
                
                # Sort entries by filename to match TOC order
                group['entries'].sort(key=lambda x: x['filename_stem'])
                
                # Find the TOC page for this section (if available)
                toc_page = 1  # Default to first page of TOC
                # Convert section_number to the format used in section_to_toc_page
                section_key = section_number
                
                # For automatic mode, section_number will be "1", "2", "3" etc.
                # For manual mode, section_number will be "14.1", "14.3" etc.
                # No conversion needed - use section_number directly
                
                # Get TOC page for this section or use main TOC page
                if section_key in section_to_toc_page:
                    toc_page = section_to_toc_page[section_key]
                    logging.info(f"Section {section_key} bookmark will point to TOC page {toc_page}")
                
                # Add section header bookmark pointing to TOC
                section_bookmark = [2, group['title'], toc_page]  # Level 2 (under main title)
                final_bookmarks.append(section_bookmark)
                
                # Add document bookmarks under this section
                for entry in group['entries']:
                    document_bookmark = [3, entry['title'], entry['page']]  # Level 3 (under section)
                    final_bookmarks.append(document_bookmark)
            
            if final_bookmarks:
                doc.set_toc(final_bookmarks)
                logging.info(f"Generated {len(final_bookmarks)} hierarchical bookmarks")
        
        # Save the final PDF
        final_output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(final_output_path), garbage=4, deflate=True)
        doc.close()
        
        # Clean up temp file
        if temp_merged_path.exists():
            temp_merged_path.unlink()
            logging.debug(f"Removed temporary file {temp_merged_path}")
        
        logging.info(f"Successfully created final PDF: {final_output_path}")
        return final_output_path
    
    except Exception as e:
        logging.error(f"Error in prepend_toc_to_pdf: {e}", exc_info=True)
        return None

def clean_text(text):
    """Clean text by removing non-printable characters and normalizing whitespace.
    
    Args:
        text: The text to clean
        
    Returns:
        Cleaned text string
    """
    import re
    # Replace non-ASCII characters and control characters
    text = re.sub(r'[\x00-\x1F\x7F-\xFF\u200B-\u200F\u2028-\u202F\u2060-\u206F]', ' ', text)
    # Replace special Unicode characters often found in RTF
    text = text.replace('\u00a0', ' ')  # Non-breaking space
    # Clean up multiple spaces
    text = re.sub(r'\s+', ' ', text)
    # Replace € and similar markers
    text = re.sub(r'[€~]', ' ', text)
    return text.strip()

# def combine_pdfs(input_folder: Path, output_path: Path) -> Path | None:
#     """Combines all PDF files in a given folder into a single PDF.

#     Args:
#         input_folder: Path to the folder containing the individual PDF files.
#         output_path: The path where the combined PDF will be saved.

#     Returns:
#         The path to the combined PDF if successful, None otherwise.
#     """
#     logging.info(f"--- Combining PDFs from {input_folder.name} into {output_path.name} ---")
#     pdf_files = sorted(input_folder.glob("*.pdf")) # Sort to maintain order if needed

#     if not pdf_files:
#         logging.warning(f"No PDF files found in {input_folder}, skipping combination.")
#         return None

#     merger = PdfWriter()

#     try:
#         for pdf_file in pdf_files:
#             try:
#                 merger.append(str(pdf_file))
#                 logging.debug(f"Appended {pdf_file.name} to the merger.")
#             except Exception as append_err:
#                 logging.error(f"Failed to append {pdf_file.name}: {append_err}")
#                 # Optionally decide whether to continue or abort on individual file failure
#                 # return None # Abort if any file fails

#         output_path.parent.mkdir(parents=True, exist_ok=True) # Ensure output dir exists
#         merger.write(str(output_path))
#         merger.close()
#         logging.info(f"Successfully combined {len(pdf_files)} PDFs into {output_path}")
#         return output_path

#     except Exception as merge_err:
#         logging.error(f"Failed to combine PDFs: {merge_err}", exc_info=True)
#         merger.close() # Ensure the writer is closed even on error
#         return None 