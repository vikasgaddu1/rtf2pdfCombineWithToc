#!/usr/bin/env python3
import sys
import logging
import os # Import os for file operations
from pathlib import Path
import pandas as pd

# Import the GUI configuration
from src.gui_config import GUIConfig

# Import the function to build the initial title DataFrame
from src.rtf_parser import build_title_dataframe

# Import the processing functions moved to the new module
from src.data_processing import (
    load_filename_section_map,
    load_ich_categories_map,
    merge_and_validate,
    create_toc_structure,
    convert_all,
    create_automatic_sections,
    save_mismatch_report_to_file
)
# Import the PDF utility functions
from src.pdf_utils import (
    combine_pdfs,
    generate_toc_pdf,
    prepend_toc_to_pdf # Added new function
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

def main(config: GUIConfig = None, progress_callback=None):
    """
    Main function to process RTF files and create a combined PDF with TOC.
    
    Args:
        config: GUIConfig object containing all settings
        progress_callback: Optional callback function to update progress (0-100)
                         Can be called with (value) or (value, file_progress)
    """
    # Close any open Word processes before starting
    close_word_processes()
    
    # Use provided config or create default for command line usage
    if config is None:
        # Default configuration for command line usage
        config = GUIConfig(
            input_folder=Path.cwd() / "input",
            output_folder=Path.cwd() / "output"
        )
    
    # Get paths from configuration
    input_folder = config.input_folder
    output_folder = config.output_folder
    output_pdf_folder = config.get_output_pdf_folder()
    docs_folder = config.get_docs_folder()
    
    # Set up file paths based on configuration
    if config.use_section_file:
        file_section_xlsx = config.section_file_path
        if not file_section_xlsx or not file_section_xlsx.exists():
            logging.error(f"Section mapping file not found: {file_section_xlsx}")
            sys.exit(1)
    else:
        file_section_xlsx = None
        
    ich_categories_xlsx = config.get_ich_categories_path()
    
    # Temporary and Final files
    toc_intermediate_pdf_path = config.get_intermediate_toc_path()
    combined_content_pdf_path = config.get_intermediate_combined_path()
    final_output_pdf_path = config.get_final_output_path()

    logging.info("Starting PDF Combination Process...")
    if progress_callback:
        progress_callback(0)

    # --- Step 1: Scan RTFs and Extract Titles ---
    logging.info(f"1. Scanning RTF files in: {input_folder}")
    titles_df = build_title_dataframe(input_folder)
    if titles_df.empty:
        logging.error("No RTF titles found; aborting.")
        sys.exit(1)
    logging.info(f"   Found {len(titles_df)} RTF files with potential titles.")
    if progress_callback:
        progress_callback(10)

    # --- Step 2: Load Mapping Files or Create Automatic Sections ---
    mismatch_report = None  # Initialize for both modes
    
    if config.use_section_file:
        logging.info(f"2. Loading mapping files from: {docs_folder}")
        try:
            filename_map = load_filename_section_map(file_section_xlsx)
            ich_map = load_ich_categories_map(ich_categories_xlsx)
            # Merge and validate using the mapping files - now returns report too
            final_df, mismatch_report = merge_and_validate(titles_df, filename_map, ich_map)
            
            # Save mismatch report to file if there are any mismatches
            if mismatch_report and mismatch_report['total_mismatched'] > 0:
                report_path = output_folder / "file_mismatch_report.txt"
                save_mismatch_report_to_file(mismatch_report, report_path)
                
        except (FileNotFoundError, KeyError) as e:
            logging.error(f"Failed to load mapping files: {e}")
            sys.exit(1)
        logging.info("   Mapping files loaded successfully.")
    else:
        logging.info("2. Creating automatic sections based on filename prefixes...")
        final_df = create_automatic_sections(titles_df)
        if final_df.empty:
            logging.error("No valid files remained after automatic section assignment; aborting.")
            sys.exit(1)
        logging.info("   Automatic sections created successfully.")
        # No mismatch report in automatic mode as all input files are used
    if progress_callback:
        progress_callback(20)

    # Ensure required columns exist for later steps
    required_cols = {'filepath', 'title', 'section_number', 'filename_stem'}
    if not required_cols.issubset(final_df.columns):
        missing_cols = required_cols - set(final_df.columns)
        logging.error(f"Critical columns missing from DataFrame after merge. Needed: {required_cols}, Missing: {missing_cols}")
        sys.exit(1)
    
    # Sort final_df to ensure consistent ordering throughout the process
    if 'section_number' in final_df.columns and 'filename_stem' in final_df.columns:
        logging.info("Sorting data by section_number and filename_stem...")
        final_df = final_df.sort_values(by=['section_number', 'filename_stem'])
        logging.info(f"   Sorted {len(final_df)} files for consistent ordering.")
    
    logging.info(f"   Validated {len(final_df)} files for processing.")
    if progress_callback:
        progress_callback(30)

    # --- Step 3: Create TOC Data Structure ---
    logging.info("3. Creating Table of Contents data structure...")
    toc_data = create_toc_structure(final_df)
    if toc_data.empty:
        logging.error("Failed to create TOC structure; aborting.")
        sys.exit(1)
    logging.info("   TOC data structure created.")
    if progress_callback:
        progress_callback(40)

    # --- Step 4: Cleanup Output PDF Folder ---
    logging.info(f"4. Checking and clearing intermediate PDF output folder: {output_pdf_folder}")
    output_pdf_folder.mkdir(parents=True, exist_ok=True)
    deleted_count = 0
    for item in output_pdf_folder.iterdir():
        if item.is_file() and item.suffix.lower() == '.pdf':
            try:
                os.remove(item)
                deleted_count += 1
                logging.debug(f"   Deleted existing file: {item.name}")
            except OSError as e:
                logging.warning(f"   Could not delete file {item.name}: {e}")
    if deleted_count > 0:
        logging.info(f"   Cleared {deleted_count} existing PDF files.")
    else:
        logging.info("   Intermediate PDF folder is empty or contains no PDFs.")
    if progress_callback:
        progress_callback(50)

    # --- Step 5: Convert RTFs to Individual PDFs ---
    logging.info(f"5. Converting {len(final_df)} RTF files to PDF in {output_pdf_folder}...")
    
    # Modify convert_all to accept progress callback
    def convert_progress_callback(file_index, total_files):
        if progress_callback:
            progress_callback(50, file_index / total_files)
    
    ok, bad = convert_all(final_df, output_pdf_folder, progress_callback=convert_progress_callback)
    logging.info(f"   Conversion Done: {ok} succeeded, {bad} failed.")
    if ok == 0:
        logging.error("No RTF files were successfully converted; aborting PDF generation.")
        sys.exit(1)
    if progress_callback:
        progress_callback(60)

    # --- Step 6: Combine Individual PDFs & Create Bookmarks ---
    logging.info(f"6. Combining individual PDFs into '{combined_content_pdf_path.name}' with bookmarks...")
    combined_pdf_path, page_map = combine_pdfs(final_df, output_pdf_folder, combined_content_pdf_path)
    if not combined_pdf_path or page_map is None:
        logging.error("Failed to combine PDF files or generate page map; aborting.")
        sys.exit(1)
    logging.info(f"   Combined content PDF created at: {combined_pdf_path}")
    logging.info(f"   Generated page map for {len(page_map)} entries.")
    if progress_callback:
        progress_callback(70)

    # --- Step 7: Generate Final TOC PDF with Links ---
    logging.info(f"7. Generating final TOC PDF ('{toc_intermediate_pdf_path.name}') with links...")
    final_toc_path, toc_page_count = generate_toc_pdf(toc_data, page_map, toc_intermediate_pdf_path, config)
    if not final_toc_path or toc_page_count is None:
        logging.error("Failed to generate final TOC PDF; aborting.")
        sys.exit(1)
    logging.info(f"   Final TOC PDF created at: {final_toc_path} ({toc_page_count} pages)")
    if progress_callback:
        progress_callback(80)

    # --- Step 8: Prepend TOC to Content PDF ---
    logging.info(f"8. Prepending TOC to content PDF to create final document: '{final_output_pdf_path.name}'...")
    final_doc_path = prepend_toc_to_pdf(final_toc_path, combined_pdf_path, final_output_pdf_path, final_df, page_map)
    if not final_doc_path:
        logging.error("Failed to prepend TOC and create final document; aborting.")
        sys.exit(1)
    logging.info(f"   Successfully created final document: {final_doc_path}")
    if progress_callback:
        progress_callback(90)

    # --- Step 9: Cleanup of Intermediate Files ---
    try:
        logging.info("9. Cleaning up intermediate files...")
        if toc_intermediate_pdf_path.exists():
            os.remove(toc_intermediate_pdf_path)
            logging.info(f"   Removed intermediate TOC: {toc_intermediate_pdf_path.name}")
        if combined_content_pdf_path.exists():
            os.remove(combined_content_pdf_path)
            logging.info(f"   Removed intermediate combined content: {combined_content_pdf_path.name}")
    except OSError as e:
        logging.warning(f"   Warning: Could not clean up intermediate file: {e}")

    logging.info("--- PDF Combination Process Completed Successfully! ---")
    if progress_callback:
        progress_callback(100)


if __name__ == "__main__":
    # For command line usage, use default configuration
    main()
