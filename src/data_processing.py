#!/usr/bin/env python3
import logging
from pathlib import Path
from typing import Tuple

import pandas as pd

# Import the converter function needed by convert_all
from src.rtf_converter import convert_rtf_to_pdf

# Configure logging (can be configured globally in main if preferred)
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# —————————————————————————————————————————————————————————————————————————
# I/O FUNCTIONS
# —————————————————————————————————————————————————————————————————————————

def load_filename_section_map(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, dtype={"section_number": str})
    for col in ("filename", "section_number"):
        if col not in df.columns:
            raise KeyError(f"'{col}' not in {xlsx_path.name}")
    df["filename"] = df["filename"].str.strip().str.lower()
    return df[["filename", "section_number"]]


def load_ich_categories_map(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, dtype={"section_number": str})
    for col in ("section_number", "ICH_section_name"):
        if col not in df.columns:
            raise KeyError(f"'{col}' not in {xlsx_path.name}")
    return df[["section_number", "ICH_section_name"]]


# —————————————————————————————————————————————————————————————————————————
# DATA MERGE & VALIDATION
# —————————————————————————————————————————————————————————————————————————

def merge_and_validate(
    titles: pd.DataFrame,
    filename_map: pd.DataFrame,
    ich_map: pd.DataFrame
) -> tuple[pd.DataFrame, dict]:
    """
    Merge and validate data from RTF files with section mapping files.
    
    Args:
        titles: DataFrame containing RTF file information
        filename_map: DataFrame mapping filenames to section numbers
        ich_map: DataFrame mapping section numbers to ICH section names
        
    Returns:
        Tuple of (validated_dataframe, mismatch_report)
        where mismatch_report contains information about files that exist
        in only one source (section file or input folder)
    """
    
    # Generate mismatch report before merging
    mismatch_report = generate_mismatch_report(titles, filename_map)
    
    # 1) attach section_number by filename
    df = titles.merge(
        filename_map,
        left_on="filename_stem",
        right_on="filename",
        how="inner",
        validate="one_to_one"
    )

    # 2) attach ICH_section_name by section_number
    df = df.merge(
        ich_map,
        on="section_number",
        how="inner",
        validate="many_to_one"
    )

    # 3) basic filename-based rules
    def is_valid(row):
        fn, sec = row.filename_stem, row.section_number
        if fn.startswith(("t", "f")) and not sec.startswith("14"):
            return False
        if fn.startswith("l") and not sec.startswith("16"):
            return False
        return True

    df["valid"] = df.apply(is_valid, axis=1)
    bad = df[~df.valid]
    if not bad.empty:
        for _, r in bad.iterrows():
            logging.warning(f"Excluding {r.filename_stem}: section {r.section_number!r} invalid")
    
    return df[df.valid].drop(columns=["filename", "valid"]), mismatch_report


def generate_mismatch_report(titles_df: pd.DataFrame, filename_map: pd.DataFrame) -> dict:
    """
    Generate a report of files that exist in only one source.
    
    Args:
        titles_df: DataFrame containing actual RTF files from input folder
        filename_map: DataFrame containing files listed in section mapping file
        
    Returns:
        Dictionary containing mismatch information with keys:
        - 'files_in_input_only': list of files in input folder but not in section file
        - 'files_in_section_only': list of files in section file but not in input folder
        - 'matched_files': list of files present in both sources
        - 'total_input_files': count of files in input folder
        - 'total_section_files': count of files in section file
    """
    
    # Get sets of filenames (normalized to lowercase for comparison)
    input_files = set(titles_df['filename_stem'].str.lower())
    section_files = set(filename_map['filename'].str.lower())
    
    # Find mismatches
    files_in_input_only = sorted(input_files - section_files)
    files_in_section_only = sorted(section_files - input_files)
    matched_files = sorted(input_files & section_files)
    
    # Create report
    report = {
        'files_in_input_only': files_in_input_only,
        'files_in_section_only': files_in_section_only,
        'matched_files': matched_files,
        'total_input_files': len(input_files),
        'total_section_files': len(section_files),
        'total_matched': len(matched_files),
        'total_mismatched': len(files_in_input_only) + len(files_in_section_only)
    }
    
    # Log the report
    logging.info("--- Manual Mode File Mismatch Report ---")
    logging.info(f"Total files in input folder: {report['total_input_files']}")
    logging.info(f"Total files in section file: {report['total_section_files']}")
    logging.info(f"Files successfully matched: {report['total_matched']}")
    logging.info(f"Total mismatched files: {report['total_mismatched']}")
    
    if files_in_input_only:
        logging.warning(f"Files in input folder but NOT in section file ({len(files_in_input_only)}):")
        for filename in files_in_input_only:
            logging.warning(f"  - {filename}")
    
    if files_in_section_only:
        logging.warning(f"Files in section file but NOT in input folder ({len(files_in_section_only)}):")
        for filename in files_in_section_only:
            logging.warning(f"  - {filename}")
    
    if not files_in_input_only and not files_in_section_only:
        logging.info("✓ All files are perfectly matched between input folder and section file!")
    
    return report


def save_mismatch_report_to_file(report: dict, output_path: Path) -> None:
    """
    Save the mismatch report to a text file.
    
    Args:
        report: Dictionary containing mismatch report data
        output_path: Path where the report file should be saved
    """
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("FILE MISMATCH REPORT - MANUAL MODE\n")
            f.write("=" * 50 + "\n\n")
            
            f.write("SUMMARY:\n")
            f.write(f"- Total files in input folder: {report['total_input_files']}\n")
            f.write(f"- Total files in section file: {report['total_section_files']}\n")
            f.write(f"- Files successfully matched: {report['total_matched']}\n")
            f.write(f"- Total mismatched files: {report['total_mismatched']}\n\n")
            
            if report['files_in_input_only']:
                f.write(f"FILES IN INPUT FOLDER BUT NOT IN SECTION FILE ({len(report['files_in_input_only'])}):\n")
                f.write("-" * 60 + "\n")
                for filename in report['files_in_input_only']:
                    f.write(f"  • {filename}\n")
                f.write("\n")
                f.write("ACTION NEEDED: Add these files to your section mapping file or remove them from the input folder.\n\n")
            
            if report['files_in_section_only']:
                f.write(f"FILES IN SECTION FILE BUT NOT IN INPUT FOLDER ({len(report['files_in_section_only'])}):\n")
                f.write("-" * 60 + "\n")
                for filename in report['files_in_section_only']:
                    f.write(f"  • {filename}\n")
                f.write("\n")
                f.write("ACTION NEEDED: Either add these RTF files to your input folder or remove them from the section mapping file.\n\n")
            
            if report['matched_files']:
                f.write(f"SUCCESSFULLY MATCHED FILES ({len(report['matched_files'])}):\n")
                f.write("-" * 60 + "\n")
                for filename in report['matched_files']:
                    f.write(f"  ✓ {filename}\n")
                f.write("\n")
            
            if not report['files_in_input_only'] and not report['files_in_section_only']:
                f.write("✓ PERFECT MATCH: All files are correctly mapped between input folder and section file!\n")
            
            f.write("\nGenerated on: " + pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        
        logging.info(f"Mismatch report saved to: {output_path}")
        
    except Exception as e:
        logging.error(f"Failed to save mismatch report to {output_path}: {e}")


# —————————————————————————————————————————————————————————————————————————
# PROCESSING LOOP
# —————————————————————————————————————————————————————————————————————————

def convert_all(final_df: pd.DataFrame, output_pdf_folder: Path, progress_callback=None) -> tuple[int, int]:
    """
    Convert all RTF files to PDFs using Word COM automation.
    
    Args:
        final_df: DataFrame containing file information
        output_pdf_folder: Path to output folder for PDFs
        progress_callback: Optional callback function to report progress
                         Called with (file_index, total_files)
    
    Returns:
        Tuple of (successful_conversions, failed_conversions)
    """
    if final_df.empty:
        logging.warning("No files to convert.")
        return 0, 0
        
    total_files = len(final_df)
    successful = 0
    failed = 0
    
    for index, row in final_df.iterrows():
        try:
            file_path = row['filepath']
            title = row['title']
            
            # Convert RTF to PDF
            pdf_path = output_pdf_folder / f"{Path(file_path).stem}.pdf"
            if convert_rtf_to_pdf(str(file_path), str(pdf_path), title):
                successful += 1
                logging.info(f"Successfully converted {file_path.name}")
            else:
                failed += 1
                logging.error(f"Failed to convert {file_path.name}")
                
            # Report progress
            if progress_callback:
                progress_callback(index + 1, total_files)
                
        except Exception as e:
            failed += 1
            logging.error(f"Error converting {file_path.name}: {e}")
            
    return successful, failed


# —————————————————————————————————————————————————————————————————————————
# TOC DATA STRUCTURE GENERATION
# —————————————————————————————————————————————————————————————————————————

def create_toc_structure(final_df: pd.DataFrame) -> pd.DataFrame:
    """Sorts the merged/validated data and creates a TOC structure DataFrame.

    Args:
        final_df: DataFrame containing merged and validated data with columns
                  like 'section_number', 'filename_stem', 'ICH_section_name',
                  'title', 'filepath'.

    Returns:
        A DataFrame formatted for TOC generation with columns:
        'level' (int): 1 for header, 2 for entry.
        'text' (str): Text to display in TOC.
        'type' (str): 'header' or 'entry'.
        'filepath' (Path | None): Original RTF path for entries, None for headers.
        'filename_stem' (str | None): Filename stem for entries, None for headers.
    """
    logging.info("Sorting data and preparing TOC structure...")
    # Sort by section, then filename stem within the section
    df_sorted = final_df.sort_values(by=['section_number', 'filename_stem'])

    toc_rows = []
    last_section = None
    for index, row in df_sorted.iterrows():
        current_section = row['section_number']
        ich_name = row['ICH_section_name']
        doc_title = row['title'] if pd.notna(row['title']) else row['filename_stem'] # Fallback title
        filepath_val = row['filepath'] # Use the correct column name from final_df
        filename_stem = row['filename_stem']

        # If this is the first row of a new section, add the section header
        if current_section != last_section:
            section_header_text = f"{current_section}  {ich_name}"
            toc_rows.append({
                'level': 1,  # Level 1 for section headers
                'text': section_header_text,
                'type': 'header',
                'filepath': None, # Headers don't have a source file
                'filename_stem': None # Keep column consistent
            })
            last_section = current_section
            logging.debug(f"Added TOC header: {section_header_text}")

        # Add the document entry row
        toc_rows.append({
            'level': 2, # Level 2 for document entries
            'text': doc_title,
            'type': 'entry',
            'filepath': filepath_val, # Use 'filepath' key consistently
            'filename_stem': filename_stem # Add filename stem
        })
        logging.debug(f"Added TOC entry: {doc_title}")

    toc_data = pd.DataFrame(toc_rows)
    logging.info(f"Created TOC data structure with {len(toc_data)} entries.")
    return toc_data


def create_automatic_sections(titles_df: pd.DataFrame) -> pd.DataFrame:
    """Creates automatic sections based on filename prefixes.
    
    Args:
        titles_df: DataFrame containing file information with 'filename_stem' column
        
    Returns:
        DataFrame with added section information
    """
    # Create a copy to avoid modifying the original
    df = titles_df.copy()
    
    # Define section mappings
    section_mappings = {
        't': '1.Tables',
        'f': '2.Figures',
        'l': '3.Listings'
    }
    
    # Add section_number and ICH_section_name columns
    df['section_number'] = None
    df['ICH_section_name'] = None
    
    # Assign sections based on filename prefix
    for prefix, section in section_mappings.items():
        mask = df['filename_stem'].str.lower().str.startswith(prefix)
        df.loc[mask, 'section_number'] = section.split('.')[0]  # Get the number part
        df.loc[mask, 'ICH_section_name'] = section.split('.')[1]  # Get the name part
    
    # Filter out files that don't match any prefix
    df = df[df['section_number'].notna()]
    
    if df.empty:
        logging.warning("No files matched the automatic section prefixes (t, f, l)")
    else:
        logging.info(f"Automatically assigned sections to {len(df)} files")
        
    return df 