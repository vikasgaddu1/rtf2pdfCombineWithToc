#!/usr/bin/env python3
import sys
import logging
import os # Import os for file operations
from pathlib import Path
import pandas as pd

# Import the function to build the initial title DataFrame
from src.rtf_parser import build_title_dataframe

# Import the processing functions moved to the new module
from src.data_processing import (
    load_filename_section_map,
    load_ich_categories_map,
    merge_and_validate,
    create_toc_structure,
    convert_all
)
# Import the PDF utility functions
from src.pdf_utils import (
    combine_pdfs,
    generate_toc_pdf,
    prepend_toc_to_pdf # Added new function
)

# —————————————————————————————————————————————————————————————————————————
# CONFIG & LOGGER
# —————————————————————————————————————————————————————————————————————————

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# —————————————————————————————————————————————————————————————————————————
# UTILITY FUNCTIONS
# —————————————————————————————————————————————————————————————————————————

def close_word_processes():
    """
    Attempt to close any open Word COM processes before starting.
    This helps prevent issues with Word instances that may be left open.
    """
    if sys.platform != 'win32':
        return
    
    try:
        import win32com.client
        import pythoncom
        
        # Try to get any existing Word Application
        try:
            word = win32com.client.GetObject("Word.Application")
            logging.info("Found open Word instances. Attempting to close...")
            word.Quit()
            word = None
            logging.info("Successfully closed Word instances.")
        except:
            # No Word instances found or couldn't get handle
            logging.info("No accessible Word instances found running.")
        
        # Force garbage collection to release COM objects
        import gc
        gc.collect()
        
    except ImportError:
        logging.warning("win32com not available - skipping Word process cleanup.")
    except Exception as e:
        logging.warning(f"Error closing Word processes: {e}")

# —————————————————————————————————————————————————————————————————————————
# MAIN
# —————————————————————————————————————————————————————————————————————————

def main():
    # Close any open Word processes before starting
    close_word_processes()
    
    # define all your paths here (or replace with argparse)
    BASE_DIR             = Path(__file__).parent
    input_folder         = BASE_DIR / "input"
    output_folder        = BASE_DIR / "output"
    output_pdf_folder    = output_folder / "_pdf" # Specific subdir for individual PDFs
    docs_folder          = BASE_DIR / "docs"
    file_section_xlsx    = docs_folder / "filename_section.xlsx"
    ich_categories_xlsx  = docs_folder / "iche3_categories.xlsx"
    
    # Temporary and Final files
    toc_intermediate_pdf_path = output_folder / "_toc_intermediate.pdf" # For the generated TOC itself
    combined_content_pdf_path = output_folder / "_combined_content.pdf" # Combined PDFs without TOC
    final_output_pdf_path   = output_folder / "final_document_with_toc.pdf" # Final output file

    logging.info("Starting PDF Combination Process...")

    # --- Step 1: Scan RTFs and Extract Titles ---
    logging.info(f"1. Scanning RTF files in: {input_folder}")
    titles_df = build_title_dataframe(input_folder)
    if titles_df.empty:
        logging.error("No RTF titles found; aborting.")
        sys.exit(1)
    logging.info(f"   Found {len(titles_df)} RTF files with potential titles.")

    # --- Step 2: Load Mapping Files ---
    logging.info(f"2. Loading mapping files from: {docs_folder}")
    try:
        filename_map = load_filename_section_map(file_section_xlsx)
        ich_map      = load_ich_categories_map(ich_categories_xlsx)
    except (FileNotFoundError, KeyError) as e:
        logging.error(f"Failed to load mapping files: {e}")
        sys.exit(1)
    logging.info("   Mapping files loaded successfully.")

    # --- Step 3: Merge Data and Validate ---
    logging.info("3. Merging and validating data...")
    final_df = merge_and_validate(titles_df, filename_map, ich_map)
    if final_df.empty:
        logging.error("No valid files remained after merge & validation; aborting.")
        sys.exit(1)
    # Ensure required columns exist for later steps
    # Check for 'title' (lowercase) instead of 'Title'
    required_cols = {'filepath', 'title', 'section_number', 'filename_stem'}
    if not required_cols.issubset(final_df.columns):
        missing_cols = required_cols - set(final_df.columns)
        logging.error(f"Critical columns missing from DataFrame after merge. Needed: {required_cols}, Missing: {missing_cols}")
        # Optional: Log existing columns for debugging
        # logging.debug(f"Existing columns: {list(final_df.columns)}")
        sys.exit(1)
    
    # Sort final_df to ensure consistent ordering throughout the process
    if 'section_number' in final_df.columns and 'filename_stem' in final_df.columns:
        logging.info("Sorting data by section_number and filename_stem...")
        final_df = final_df.sort_values(by=['section_number', 'filename_stem'])
        logging.info(f"   Sorted {len(final_df)} files for consistent ordering.")
    
    logging.info(f"   Validated {len(final_df)} files for processing.")

    # --- Step 4: Create TOC Data Structure ---
    logging.info("4. Creating Table of Contents data structure...")
    # Note: create_toc_structure also sorts the dataframe internally
    toc_data = create_toc_structure(final_df)
    if toc_data.empty:
        # This shouldn't happen if final_df was not empty, but check anyway
        logging.error("Failed to create TOC structure; aborting.")
        sys.exit(1)
    # Ensure required columns exist for later steps
    if not {'filepath', 'level', 'text', 'type'}.issubset(toc_data.columns):
        logging.error("Critical columns missing from TOC data. Needed: filepath, level, text, type")
        sys.exit(1)
    logging.info("   TOC data structure created.")

    # --- Step 5: Cleanup Output PDF Folder ---
    logging.info(f"5. Checking and clearing intermediate PDF output folder: {output_pdf_folder}")
    output_pdf_folder.mkdir(parents=True, exist_ok=True) # Ensure the directory exists
    deleted_count = 0
    for item in output_pdf_folder.iterdir():
        if item.is_file() and item.suffix.lower() == '.pdf':
            try:
                os.remove(item) # Use os.remove to delete files
                deleted_count += 1
                logging.debug(f"   Deleted existing file: {item.name}")
            except OSError as e:
                logging.warning(f"   Could not delete file {item.name}: {e}")
    if deleted_count > 0:
        logging.info(f"   Cleared {deleted_count} existing PDF files.")
    else:
        logging.info("   Intermediate PDF folder is empty or contains no PDFs.")

    # --- Step 6: Convert RTFs to Individual PDFs ---
    logging.info(f"6. Converting {len(final_df)} RTF files to PDF in {output_pdf_folder}...")
    # convert_all uses the order from final_df (already sorted by create_toc_structure's sort)
    ok, bad = convert_all(final_df, output_pdf_folder)
    logging.info(f"   Conversion Done: {ok} succeeded, {bad} failed.")
    if ok == 0:
        logging.error("No RTF files were successfully converted; aborting PDF generation.")
        sys.exit(1)

    # --- Step 7: Combine Individual PDFs & Create Bookmarks ---
    logging.info(f"7. Combining individual PDFs into '{combined_content_pdf_path.name}' with bookmarks...")
    # Pass final_df which now has the correct 'title' column
    combined_pdf_path, page_map = combine_pdfs(final_df, output_pdf_folder, combined_content_pdf_path)
    if not combined_pdf_path or page_map is None:
        logging.error("Failed to combine PDF files or generate page map; aborting.")
        sys.exit(1)
    logging.info(f"   Combined content PDF created at: {combined_pdf_path}")
    logging.info(f"   Generated page map for {len(page_map)} entries.")

    # --- Step 8: Generate Final TOC PDF with Links ---
    logging.info(f"8. Generating final TOC PDF ('{toc_intermediate_pdf_path.name}') with links...")
    final_toc_path, toc_page_count = generate_toc_pdf(toc_data, page_map, toc_intermediate_pdf_path)
    if not final_toc_path or toc_page_count is None:
        logging.error("Failed to generate final TOC PDF; aborting.")
        sys.exit(1)
    logging.info(f"   Final TOC PDF created at: {final_toc_path} ({toc_page_count} pages)")

    # --- Step 9: Prepend TOC to Content PDF ---
    logging.info(f"9. Prepending TOC to content PDF to create final document: '{final_output_pdf_path.name}'...")
    # Pass final_df and page_map to generate bookmarks
    final_doc_path = prepend_toc_to_pdf(final_toc_path, combined_pdf_path, final_output_pdf_path, final_df, page_map)
    if not final_doc_path:
        logging.error("Failed to prepend TOC and create final document; aborting.")
        sys.exit(1)
    logging.info(f"   Successfully created final document: {final_doc_path}")

    # --- Step 10: Cleanup of Intermediate Files ---
    try:
        logging.info("10. Cleaning up intermediate files...")
        if toc_intermediate_pdf_path.exists():
            os.remove(toc_intermediate_pdf_path)
            logging.info(f"   Removed intermediate TOC: {toc_intermediate_pdf_path.name}")
        if combined_content_pdf_path.exists():
            os.remove(combined_content_pdf_path)
            logging.info(f"   Removed intermediate combined content: {combined_content_pdf_path.name}")
        # Keep the _pdf folder with individual PDFs? Or delete?
        # For now, let's keep it for debugging.
        # import shutil
        # if output_pdf_folder.exists():
        #     shutil.rmtree(output_pdf_folder)
        #     logging.info(f"   Removed intermediate PDF folder: {output_pdf_folder}")
    except OSError as e:
        logging.warning(f"   Warning: Could not clean up intermediate file: {e}")

    logging.info("--- PDF Combination Process Completed Successfully! ---")


if __name__ == "__main__":
    main()
