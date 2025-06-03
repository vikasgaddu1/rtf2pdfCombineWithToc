"""
Simple configuration holder for GUI settings.
No file persistence - all settings come from GUI for each run.
"""

from pathlib import Path
from dataclasses import dataclass
from typing import Optional

@dataclass
class GUIConfig:
    """Holds all configuration from GUI without persistence."""
    
    # Paths
    input_folder: Path
    output_folder: Path
    output_pdf_folder: str = "_pdf"
    docs_folder: str = "docs"
    
    # Files
    final_output: str = "final_document_with_toc.pdf"
    ich_categories: str = "iche3_categories.xlsx"
    intermediate_toc: str = "_toc_intermediate.pdf"
    intermediate_combined: str = "_combined_content.pdf"
    
    # Section file settings
    use_section_file: bool = False
    section_file_path: Optional[Path] = None
    section_file_name: Optional[str] = None
    
    # PDF settings
    page_width_mm: float = 210.0
    margin_mm: float = 15.0
    font_size: float = 8.0
    header_font_size: float = 10.0
    
    # Logging
    log_level: str = "INFO"
    log_format: str = "%(asctime)s - %(levelname)s - %(message)s"
    
    def get_output_pdf_folder(self) -> Path:
        """Get the full path to the PDF output folder."""
        return self.output_folder / self.output_pdf_folder
    
    def get_docs_folder(self) -> Path:
        """Get the full path to the docs folder."""
        return Path(__file__).parent.parent / self.docs_folder
    
    def get_ich_categories_path(self) -> Path:
        """Get the full path to the ICH categories file."""
        return self.get_docs_folder() / self.ich_categories
    
    def get_section_file_path(self) -> Optional[Path]:
        """Get the full path to the section file if enabled."""
        if self.use_section_file and self.section_file_path:
            return self.section_file_path
        return None
    
    def get_intermediate_toc_path(self) -> Path:
        """Get the full path to the intermediate TOC file."""
        return self.output_folder / self.intermediate_toc
    
    def get_intermediate_combined_path(self) -> Path:
        """Get the full path to the intermediate combined file."""
        return self.output_folder / self.intermediate_combined
    
    def get_final_output_path(self) -> Path:
        """Get the full path to the final output file."""
        return self.output_folder / self.final_output 