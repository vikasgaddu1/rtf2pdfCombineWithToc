#!/usr/bin/env python3
"""
Test script to demonstrate the section file configuration issue.
This shows that the GUI updates the config in memory but doesn't save it to file.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.config_manager import ConfigManager

def test_section_file_issue():
    print("Testing Section File Configuration Issue")
    print("=" * 50)
    
    # Initialize config manager
    config = ConfigManager()
    
    # Show initial state
    print(f"Initial section file enabled: {config.is_section_file_enabled()}")
    print(f"Initial section file name: {config.get_file_path('section_file')['name']}")
    
    # Simulate what the GUI does - update config in memory
    print("\nSimulating GUI update (enabling section file)...")
    new_config = {
        'files': {
            'section_file': {
                'enabled': True,
                'name': "my_sections.xlsx"
            }
        }
    }
    config.update_config(new_config)
    
    # Check if update worked in memory
    print(f"After update - section file enabled: {config.is_section_file_enabled()}")
    print(f"After update - section file name: {config.get_file_path('section_file')['name']}")
    
    # Now create a new config instance (simulating what main.py does)
    print("\nCreating new ConfigManager instance (simulating main.py)...")
    config2 = ConfigManager()
    
    print(f"New instance - section file enabled: {config2.is_section_file_enabled()}")
    print(f"New instance - section file name: {config2.get_file_path('section_file')['name']}")
    
    print("\nISSUE: The configuration update was not persisted to file!")
    print("The GUI updates the config in memory but doesn't save it.")
    print("When main.py creates a new ConfigManager, it reads the original file.")

if __name__ == "__main__":
    test_section_file_issue() 