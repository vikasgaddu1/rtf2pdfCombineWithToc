#!/usr/bin/env python3
import sys
import os
from pathlib import Path

# Add the parent directory to the Python path
sys.path.append(str(Path(__file__).parent))

# Import and run the GUI
from src.gui import main

if __name__ == "__main__":
    main() 
    