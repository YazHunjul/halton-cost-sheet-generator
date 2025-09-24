#!/usr/bin/env python3
"""
Debug script to reproduce the Word generation error.
"""
import sys
sys.path.append('src')

from utils.excel import save_to_excel
from utils.word import generate_quotation_document
import os
import traceback

def debug_word_error():
    """Debug the Word generation error."""
    
    # Create test project data
    project_data = {
        "project_name": "Debug Word Error",
        "project_number": "P12345",
        "customer": "Test Customer",
        "date": "01/01/2025",
        "project_location": "Test Location",
        "delivery_location": "Test Delivery",
        "estimator": "Joe Salloum",
        "revision": "A",
        "project_type": "Commercial Kitchen",
        "levels": [
            {
                "level_number": 1,
                "level_name": "Level 1",
                "areas": [
                    {
                        "name": "Main Kitchen",
                        "canopies": [
                            {
                                "reference_number": "1.01",
                                "model": "KV",
                                "configuration": "Wall",
                                "wall_cladding": {"type": "None"},
                                "options": {"fire_suppression": False}
                            }
                        ],
                        "options": {"uvc": False, "sdu": True, "recoair": False}
                    }
                ]
            }
        ]
    }
    
    print("=== Debugging Word Generation Error ===")
    
    try:
        # Generate Excel file first
        print("1. Generating Excel file...")
        excel_path = save_to_excel(project_data)
        print(f"   Excel file created: {excel_path}")
        
        # Try to generate Word document
        print("\n2. Attempting to generate Word document...")
        word_path = generate_quotation_document(project_data, excel_path)
        print(f"   Word document created: {word_path}")
        
        # Clean up
        if os.path.exists(excel_path):
            os.remove(excel_path)
        if os.path.exists(word_path):
            os.remove(word_path)
            
        print("\n✅ Word generation successful!")
        
    except Exception as e:
        print(f"\n❌ Error occurred: {str(e)}")
        print("\nFull traceback:")
        traceback.print_exc()
        
        # Clean up on error
        if 'excel_path' in locals() and os.path.exists(excel_path):
            os.remove(excel_path)

if __name__ == "__main__":
    debug_word_error() 