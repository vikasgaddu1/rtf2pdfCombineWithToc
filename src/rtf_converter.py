import sys
import logging
import gc
from pathlib import Path

# Optional imports
try:
    import fitz  # PyMuPDF for bookmarks
except ImportError:
    fitz = None

# Only import com client if on Windows
if sys.platform == 'win32':
    try:
        import win32com.client
    except ImportError:
        win32com = None
else:
    win32com = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

WD_FORMAT_PDF = 17  # Word constant


def _add_bookmark(pdf_path: Path, title: str) -> bool:
    """Open the PDF at pdf_path, add a top‐level bookmark, and overwrite it."""
    if not fitz:
        logging.warning("PyMuPDF not installed; skipping bookmark.")
        return False

    tmp_path = pdf_path.with_suffix(pdf_path.suffix + ".tmp")
    try:
        with fitz.open(pdf_path) as src, fitz.open() as dst:
            dst.insert_pdf(src)
            dst.set_toc([(1, title, 1)])
            dst.save(tmp_path, garbage=4, deflate=True)

        tmp_path.replace(pdf_path)
        logging.info(f"Bookmarked PDF saved: {pdf_path.name}")
        return True

    except Exception as err:
        logging.error(f"Failed to add bookmark: {err}")
        tmp_path.unlink(missing_ok=True)
        return False


def convert_rtf_to_pdf(rtf_path: str, pdf_path: str, title: str = None) -> bool:
    """
    Convert an RTF to PDF via Word COM; optionally add a bookmark.
    Returns True if conversion succeeded (bookmark failures don't fail conversion).
    """
    rtf = Path(rtf_path)
    pdf = Path(pdf_path)

    if sys.platform != 'win32':
        logging.error("RTF→PDF conversion only supported on Windows.")
        return False

    if not win32com:
        logging.error("pywin32 is required for COM automation.")
        return False

    # Ensure output directory exists
    pdf.parent.mkdir(parents=True, exist_ok=True)

    word = None
    doc = None

    try:
        logging.info(f"Converting {rtf.name} → {pdf.name}")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        # Ensure absolute paths are passed to Word
        rtf_abs = str(rtf.resolve())
        pdf_abs = str(pdf.resolve())

        doc = word.Documents.Open(rtf_abs, ReadOnly=True)
        doc.SaveAs(pdf_abs, FileFormat=WD_FORMAT_PDF)
        logging.info("PDF conversion succeeded.")

        # Add bookmark if requested
        # <<< Commenting out bookmark addition during RTF conversion >>>
        # if title:
        #     _add_bookmark(pdf, title)

        return True

    except Exception as e:
        logging.error(f"Conversion error: {e}")
        return False

    finally:
        # Cleanly close Word objects
        closed_doc = False
        quit_word = False
        try:
            if doc:
                logging.debug(f"Attempting to close document for {rtf.name}")
                doc.Close(False)
                closed_doc = True
                logging.debug(f"Document closed for {rtf.name}")
        except Exception as doc_close_err:
            logging.warning(f"Error closing document for {rtf.name}: {doc_close_err}")
        finally:
            # Always try to quit Word, even if doc close failed
            try:
                if word:
                    logging.debug(f"Attempting to quit Word for {rtf.name}")
                    word.Quit()
                    quit_word = True
                    logging.debug(f"Word Quit command issued for {rtf.name}")
            except Exception as word_quit_err:
                logging.warning(f"Error quitting Word for {rtf.name}: {word_quit_err}")
            finally:
                # Release COM objects and collect garbage
                doc = None
                word = None
                # Explicitly call com release (optional, DispatchEx might handle it)
                # if sys.platform == 'win32' and win32com:
                #    win32com.client.pythoncom.CoUninitialize()
                #    win32com.client.pythoncom.CoInitialize()
                logging.debug(f"Running garbage collection after {rtf.name}")
                gc.collect()
                logging.debug(f"Garbage collection finished after {rtf.name}")
