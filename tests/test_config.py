#!/usr/bin/env python3
"""
Test script to verify configuration loading and section file settings.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.config_manager import ConfigManager

def test_config():
    print("Testing Configuration Manager...")
    
    # Initialize config manager
    config = ConfigManager()
    
    # Test loading configuration
    print(f"Section file enabled: {config.is_section_file_enabled()}")
    print(f"Section file name: {config.get_file_path('section_file')['name']}")
    print(f"Input folder: {config.get_path('input_folder')}")
    print(f"Output folder: {config.get_path('output_folder')}")
    print(f"Final output: {config.get_file_path('final_output')}")
    
    # Test updating configuration
    print("\nTesting configuration update...")
    new_config = {
        'files': {
            'section_file': {
                'enabled': True,
                'name': "test_section.xlsx"
            }
        }
    }
    
    config.update_config(new_config)
    print(f"After update - Section file enabled: {config.is_section_file_enabled()}")
    print(f"After update - Section file name: {config.get_file_path('section_file')['name']}")
    
    # Test reverting
    print("\nTesting revert to automatic mode...")
    revert_config = {
        'files': {
            'section_file': {
                'enabled': False,
                'name': "filename_section.xlsx"
            }
        }
    }
    
    config.update_config(revert_config)
    print(f"After revert - Section file enabled: {config.is_section_file_enabled()}")
    print(f"After revert - Section file name: {config.get_file_path('section_file')['name']}")

if __name__ == "__main__":
    test_config() 