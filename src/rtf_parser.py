import logging
import os
from pathlib import Path
import pandas as pd
from striprtf.striprtf import rtf_to_text

# Configure logging (consistent with other modules)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def _extract_title_from_single_rtf(rtf_path: Path, max_bytes: int = 10000) -> str | None:
    """Internal helper to extract title from a single RTF file."""
    base_name = rtf_path.name # For logging
    try:
        # First try opening with binary mode to check RTF header
        with open(rtf_path, 'rb') as file:
            # Read just enough for the header check
            header = file.read(6) # b'{\\\\rtf' is 6 bytes

            # Check if we have the RTF header
            if not header.startswith(b'{\\rtf'):
                logging.warning(f"File does not appear to start with RTF header: {base_name}")
                # Proceed anyway, but log warning

            # Reset file position and read limited content
            file.seek(0)
            rtf_binary = file.read(max_bytes)

        # Convert binary content to string using a forgiving encoding
        # latin-1 (ISO-8859-1) can handle any byte value
        rtf_content = rtf_binary.decode('latin-1', errors='ignore')

        # Convert RTF to plain text using striprtf
        plain_text = rtf_to_text(rtf_content)

        if not plain_text:
             logging.warning(f"No text content extracted (within first {max_bytes} bytes) from {base_name}")
             return None

        # Get the first non-empty line as title
        lines = plain_text.split('\n')
        title = next((line.strip() for line in lines if line.strip()), None) # Return None if no title
        if title:
            # remove trailing | from title if present
            title = title.rstrip('|').strip()
            if title: # Check if title is not empty after stripping
                logging.debug(f"Extracted title '{title}' from {base_name}")
                return title
            else:
                 logging.warning(f"Extracted title was empty or only '|' for {base_name}")
                 return None # Treat empty title as no title found
        else:
            logging.warning(f"No non-empty lines found (within first {max_bytes} bytes) to use as title in {base_name}")
            return None

    except FileNotFoundError:
        # This shouldn't happen if called from build_title_dataframe which finds the file first
        logging.error(f"RTF file not found during title extraction: {rtf_path}")
        return None
    except Exception as e:
        logging.error(f"Error processing {base_name} for title: {e}")
        return None


def build_title_dataframe(input_dir: Path, max_bytes: int = 10000) -> pd.DataFrame:
    """
    Scans an input directory for RTF files, extracts the title from each,
    and returns a pandas DataFrame mapping absolute file paths to titles.

    Args:
        input_dir: Path object representing the directory containing RTF files.
        max_bytes: Maximum number of bytes to read per file for title extraction.

    Returns:
        A pandas DataFrame with columns 'filepath' (Path object), 'filename_stem' (str), 
        and 'title' (str | None).
    """
    rtf_files = list(input_dir.glob('*.rtf'))
    if not rtf_files:
        logging.warning(f"No RTF files found in {input_dir}")
        return pd.DataFrame({'filepath': [], 'filename_stem': [], 'title': []}) # Return empty DataFrame with new column

    logging.info(f"Found {len(rtf_files)} RTF files in {input_dir}. Extracting titles...")

    data = []
    for rtf_path in rtf_files:
        title = _extract_title_from_single_rtf(rtf_path, max_bytes)
        filename_stem = rtf_path.stem # Get filename without extension
        # Store absolute path for consistency
        data.append({
            'filepath': rtf_path.resolve(), 
            'filename_stem': filename_stem,
            'title': title
        })
        # Log progress periodically if needed (e.g., every 50 files)

    df = pd.DataFrame(data)
    logging.info(f"Finished extracting titles. Found titles for {df['title'].notna().sum()} out of {len(df)} files.")
    return df
