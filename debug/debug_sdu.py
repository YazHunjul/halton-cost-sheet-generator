#!/usr/bin/env python3
"""
Debug SDU functionality to find the exact source of the merged cell error.
"""
import sys
sys.path.append('src')

from utils.excel import load_template_workbook
from openpyxl import load_workbook

def debug_sdu_metadata():
    """Debug the SDU metadata writing to find the merged cell issue."""
    
    print("Loading template workbook...")
    wb = load_template_workbook()
    
    # Get the first SDU sheet
    sdu_sheets = [sheet for sheet in wb.sheetnames if 'SDU' in sheet and 'CANOPY' not in sheet and 'FIRE' not in sheet]
    if not sdu_sheets:
        print("No SDU sheets found!")
        return
    
    print(f"Found SDU sheets: {sdu_sheets}")
    sdu_sheet = wb[sdu_sheets[0]]
    
    print("Testing individual cell writes...")
    
    # Test each cell individually
    test_cells = ['C12', 'D37', 'E37', 'D3', 'D5', 'D7', 'H3', 'H5', 'H7', 'O7']
    
    for cell in test_cells:
        try:
            print(f"Testing cell {cell}...")
            current_value = sdu_sheet[cell].value
            print(f"  Current value: {current_value}")
            print(f"  Cell type: {type(sdu_sheet[cell])}")
            
            # Try to write a test value
            sdu_sheet[cell] = "TEST"
            print(f"  ✅ Successfully wrote to {cell}")
            
        except Exception as e:
            print(f"  ❌ Error writing to {cell}: {str(e)}")
            print(f"  Cell type: {type(sdu_sheet[cell])}")
            
            # Check if it's a merged cell
            is_merged = False
            for merged_range in sdu_sheet.merged_cells.ranges:
                if cell in merged_range:
                    print(f"  📋 Cell {cell} is part of merged range: {merged_range}")
                    is_merged = True
                    break
            
            if not is_merged:
                print(f"  📋 Cell {cell} is not merged")
    
    wb.close()

if __name__ == "__main__":
    debug_sdu_metadata() 