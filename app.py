"""
Halton Cost Sheet Generator - Streamlit Cloud Entry Point
Main application entry point for deployment on Streamlit Cloud.
"""

import sys
import os

# Add src directory to Python path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import and run the main application
from src.app import main

if __name__ == "__main__":
    main() 