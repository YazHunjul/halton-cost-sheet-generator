#!/usr/bin/env python3
"""
Debug script to check fire suppression data in Excel files.
"""

import sys
import os
import json

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

def debug_fire_suppression_data(excel_path):
    """Debug fire suppression data from an Excel file."""
    print(f"=== Debugging Fire Suppression Data from {excel_path} ===")
    
    try:
        from utils.excel import read_excel_project_data
        from utils.word import prepare_template_context
        
        # Read project data from Excel
        print("\n1. Reading Excel data...")
        project_data = read_excel_project_data(excel_path)
        
        # Check canopies for fire suppression data
        print("\n2. Checking canopies for fire suppression tank quantities...")
        total_canopies = 0
        canopies_with_fire_suppression = 0
        
        for level in project_data.get('levels', []):
            for area in level.get('areas', []):
                for canopy in area.get('canopies', []):
                    total_canopies += 1
                    tank_qty = canopy.get('fire_suppression_tank_quantity', 0)
                    
                    print(f"  Canopy {canopy.get('reference_number', 'Unknown')}:")
                    print(f"    Tank Quantity: {tank_qty}")
                    print(f"    Model: {canopy.get('model', 'Unknown')}")
                    print(f"    Location: {level.get('level_name', 'Unknown')} - {area.get('name', 'Unknown')}")
                    
                    if tank_qty > 0:
                        canopies_with_fire_suppression += 1
                        print(f"    ✅ HAS FIRE SUPPRESSION")
                    else:
                        print(f"    ❌ NO FIRE SUPPRESSION")
                    print()
        
        print(f"Summary:")
        print(f"  Total canopies: {total_canopies}")
        print(f"  Canopies with fire suppression: {canopies_with_fire_suppression}")
        
        # Process through Word context
        print("\n3. Processing through Word context...")
        context = prepare_template_context(project_data)
        fire_suppression_items = context.get('fire_suppression_items', [])
        
        print(f"Fire suppression items in context: {len(fire_suppression_items)}")
        
        if fire_suppression_items:
            print("\nFire suppression items found:")
            for item in fire_suppression_items:
                print(f"  - Item {item.get('item_number')}: {item.get('tank_quantity')} tanks at {item.get('level_area_combined')}")
        else:
            print("\n❌ No fire suppression items found in context")
            print("\nPossible reasons:")
            print("1. No canopies have tank quantities > 0")
            print("2. Excel cells C18, C35, C52, etc. are empty")
            print("3. Tank quantity values are not in expected format (e.g., '1 TANK', '2 TANK')")
        
        # Save debug data
        debug_data = {
            'project_data': project_data,
            'context': context,
            'fire_suppression_items': fire_suppression_items
        }
        
        with open('debug_fire_suppression_data.json', 'w') as f:
            json.dump(debug_data, f, indent=2)
        
        print(f"\n✅ Debug data saved to debug_fire_suppression_data.json")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python debug_fire_suppression_data.py <excel_file_path>")
        print("Example: python debug_fire_suppression_data.py 'output/PROJECT-123 Cost Sheet 25052024.xlsx'")
    else:
        excel_path = sys.argv[1]
        debug_fire_suppression_data(excel_path) 