#!/usr/bin/env python3
import argparse
import logging
from pathlib import Path
from src.gui_config import GUIConfig

def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert RTF files to a single PDF with table of contents and bookmarks."
    )
    
    # Input/Output paths
    parser.add_argument(
        "--input-folder",
        type=Path,
        default=Path.cwd() / "input",
        help="Folder containing RTF files (default: ./input)"
    )
    parser.add_argument(
        "--output-folder",
        type=Path,
        default=Path.cwd() / "output",
        help="Folder for output files (default: ./output)"
    )
    parser.add_argument(
        "--output-filename",
        type=str,
        default="final_document_with_toc.pdf",
        help="Name of the final combined PDF (default: final_document_with_toc.pdf)"
    )
    
    # Section file options
    parser.add_argument(
        "--use-section-file",
        action="store_true",
        help="Use a section mapping Excel file instead of automatic section organization"
    )
    parser.add_argument(
        "--section-file",
        type=Path,
        help="Path to the section mapping Excel file (required if --use-section-file is set)"
    )
    
    # PDF settings
    parser.add_argument(
        "--page-width",
        type=float,
        default=210.0,
        help="Page width in millimeters (default: 210.0)"
    )
    parser.add_argument(
        "--margin",
        type=float,
        default=15.0,
        help="Page margin in millimeters (default: 15.0)"
    )
    parser.add_argument(
        "--font-size",
        type=float,
        default=8.0,
        help="Base font size (default: 8.0)"
    )
    parser.add_argument(
        "--header-font-size",
        type=float,
        default=10.0,
        help="Header font size (default: 10.0)"
    )
    
    # Logging options
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="Set the logging level (default: INFO)"
    )
    
    return parser.parse_args()

def validate_args(args):
    """Validate command line arguments."""
    # Check input folder
    if not args.input_folder.exists():
        raise ValueError(f"Input folder does not exist: {args.input_folder}")
    
    # Check section file if enabled
    if args.use_section_file:
        if not args.section_file:
            raise ValueError("--section-file is required when --use-section-file is set")
        if not args.section_file.exists():
            raise ValueError(f"Section file does not exist: {args.section_file}")
    
    # Validate numeric values
    if args.page_width <= 0:
        raise ValueError("Page width must be positive")
    if args.margin < 0:
        raise ValueError("Margin cannot be negative")
    if args.font_size <= 0:
        raise ValueError("Font size must be positive")
    if args.header_font_size <= 0:
        raise ValueError("Header font size must be positive")

def create_config_from_args(args):
    """Create a GUIConfig object from command line arguments."""
    config = GUIConfig(
        input_folder=args.input_folder,
        output_folder=args.output_folder,
        final_output=args.output_filename,
        use_section_file=args.use_section_file,
        section_file_path=args.section_file if args.use_section_file else None,
        section_file_name=args.section_file.name if args.use_section_file and args.section_file else None,
        page_width_mm=args.page_width,
        margin_mm=args.margin,
        font_size=args.font_size,
        header_font_size=args.header_font_size,
        log_level=args.log_level
    )
    return config

def main():
    """Main entry point for CLI version."""
    # Parse and validate arguments
    args = parse_args()
    try:
        validate_args(args)
    except ValueError as e:
        logging.error(str(e))
        return 1
    
    # Create configuration
    config = create_config_from_args(args)
    
    # Import here to avoid circular imports
    from main import main as process_main
    
    # Run the main process
    try:
        process_main(config=config)
        return 0
    except Exception as e:
        logging.error(f"Error during processing: {e}")
        return 1

if __name__ == "__main__":
    main() 