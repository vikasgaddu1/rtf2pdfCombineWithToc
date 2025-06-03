#!/usr/bin/env python3
"""
Test script to verify the section file configuration fix.
This shows that the configuration is now properly saved and persisted.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.config_manager import ConfigManager

def test_section_file_fix():
    print("Testing Section File Configuration Fix")
    print("=" * 50)
    
    # Initialize config manager
    config = ConfigManager()
    
    # Show initial state
    print(f"Initial section file enabled: {config.is_section_file_enabled()}")
    print(f"Initial section file name: {config.get_file_path('section_file')['name']}")
    
    # Store original values to restore later
    original_enabled = config.is_section_file_enabled()
    original_name = config.get_file_path('section_file')['name']
    
    # Simulate what the GUI does - update config and save
    print("\nSimulating GUI update with save (enabling section file)...")
    new_config = {
        'files': {
            'section_file': {
                'enabled': True,
                'name': "test_sections.xlsx"
            }
        }
    }
    config.update_config(new_config)
    
    # Now save the config (this is what was missing!)
    save_result = config.save_config()
    print(f"Configuration saved: {save_result}")
    
    # Check if update worked in memory
    print(f"After update - section file enabled: {config.is_section_file_enabled()}")
    print(f"After update - section file name: {config.get_file_path('section_file')['name']}")
    
    # Now create a new config instance (simulating what main.py does)
    print("\nCreating new ConfigManager instance (simulating main.py)...")
    config2 = ConfigManager()
    
    print(f"New instance - section file enabled: {config2.is_section_file_enabled()}")
    print(f"New instance - section file name: {config2.get_file_path('section_file')['name']}")
    
    if config2.is_section_file_enabled() == True and config2.get_file_path('section_file')['name'] == "test_sections.xlsx":
        print("\nSUCCESS: The configuration was properly persisted!")
    else:
        print("\nFAILED: The configuration was not persisted.")
    
    # Restore original configuration
    print("\nRestoring original configuration...")
    restore_config = {
        'files': {
            'section_file': {
                'enabled': original_enabled,
                'name': original_name
            }
        }
    }
    config2.update_config(restore_config)
    config2.save_config()
    print("Original configuration restored.")

if __name__ == "__main__":
    test_section_file_fix() 