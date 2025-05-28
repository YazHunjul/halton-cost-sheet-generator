"""
Halton Cost Sheet Generator - Streamlit Cloud Entry Point
Main application entry point for deployment on Streamlit Cloud.
"""

import sys
import os

# Add src directory to Python path for imports
src_dir = os.path.join(os.path.dirname(__file__), 'src')
sys.path.insert(0, src_dir)

# Import the main function using importlib to avoid naming conflicts
import importlib.util

# Load the src/app.py module directly
app_path = os.path.join(src_dir, 'app.py')
spec = importlib.util.spec_from_file_location("main_app", app_path)
main_app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(main_app)

if __name__ == "__main__":
    main_app.main() 