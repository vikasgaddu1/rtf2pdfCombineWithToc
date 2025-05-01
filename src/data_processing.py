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
) -> pd.DataFrame:
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
    return df[df.valid].drop(columns=["filename", "valid"])


# —————————————————————————————————————————————————————————————————————————
# PROCESSING LOOP
# —————————————————————————————————————————————————————————————————————————

def convert_all(
    df: pd.DataFrame,
    output_folder: Path
) -> Tuple[int, int]:
    output_folder.mkdir(parents=True, exist_ok=True)
    success = fail = 0
    
    # Ensure df is sorted by section_number and filename_stem for consistent ordering
    if 'section_number' in df.columns and 'filename_stem' in df.columns:
        df = df.sort_values(by=['section_number', 'filename_stem'])
        logging.info(f"Sorted {len(df)} files by section_number and filename_stem for conversion")

    for _, row in df.iterrows():
        src: Path = row.filepath
        title: str = row.title if pd.notna(row.title) else None
        dest = output_folder / f"{src.stem}.pdf"

        if title is None:
            logging.info(f"No title for {src.name!r}, skipping bookmark")
        else:
            logging.info(f"Using title {title!r} for {src.name!r}")

        if convert_rtf_to_pdf(str(src), str(dest), title=title):
            success += 1
        else:
            fail += 1

    return success, fail


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