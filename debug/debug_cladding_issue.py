#!/usr/bin/env python3
"""
Debug script to trace cladding data flow and identify why cladding is not showing in Word documents.
"""

import json
import sys
import os

# Add src directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from utils.excel import read_excel_project_data
from utils.word import prepare_template_context

def debug_cladding_data_flow(excel_path: str):
    """
    Debug the complete cladding data flow from Excel to Word template context.
    
    Args:
        excel_path (str): Path to Excel file with cladding data
    """
    print("🔍 DEBUGGING CLADDING DATA FLOW")
    print("=" * 50)
    
    # Step 1: Read Excel data
    print("\n1️⃣ READING EXCEL DATA")
    print("-" * 30)
    
    try:
        project_data = read_excel_project_data(excel_path)
        print(f"✅ Successfully read Excel file: {excel_path}")
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return
    
    # Step 2: Examine raw project data for cladding
    print("\n2️⃣ RAW PROJECT DATA - CLADDING SEARCH")
    print("-" * 40)
    
    cladding_found_in_raw = False
    for level_idx, level in enumerate(project_data.get('levels', [])):
        level_name = level.get('level_name', f'Level {level_idx}')
        print(f"\n📁 Level: {level_name}")
        
        for area_idx, area in enumerate(level.get('areas', [])):
            area_name = area.get('name', f'Area {area_idx}')
            print(f"  📂 Area: {area_name}")
            
            for canopy_idx, canopy in enumerate(area.get('canopies', [])):
                ref_num = canopy.get('reference_number', f'Canopy {canopy_idx}')
                print(f"    🔹 Canopy: {ref_num}")
                
                # Check cladding price
                cladding_price = canopy.get('cladding_price', 0)
                print(f"      • cladding_price: {cladding_price}")
                
                # Check wall_cladding structure
                wall_cladding = canopy.get('wall_cladding', {})
                print(f"      • wall_cladding type: {wall_cladding.get('type', 'NOT_FOUND')}")
                print(f"      • wall_cladding width: {wall_cladding.get('width', 'NOT_FOUND')}")
                print(f"      • wall_cladding height: {wall_cladding.get('height', 'NOT_FOUND')}")
                print(f"      • wall_cladding position: {wall_cladding.get('position', 'NOT_FOUND')}")
                
                # Check if cladding exists
                has_cladding_price = cladding_price > 0
                has_cladding_type = wall_cladding.get('type') not in ['None', None, '']
                
                if has_cladding_price or has_cladding_type:
                    cladding_found_in_raw = True
                    print(f"      ✅ CLADDING DETECTED!")
                    print(f"         - Has price: {has_cladding_price}")
                    print(f"         - Has type: {has_cladding_type}")
                else:
                    print(f"      ❌ No cladding detected")
    
    if not cladding_found_in_raw:
        print("\n❌ NO CLADDING DATA FOUND IN RAW PROJECT DATA")
        print("   This suggests the issue is in Excel reading, not Word generation.")
        return
    else:
        print(f"\n✅ CLADDING DATA FOUND IN RAW PROJECT DATA")
    
    # Step 3: Prepare template context
    print("\n3️⃣ PREPARING TEMPLATE CONTEXT")
    print("-" * 35)
    
    try:
        context = prepare_template_context(project_data)
        print("✅ Template context prepared successfully")
    except Exception as e:
        print(f"❌ Error preparing template context: {e}")
        return
    
    # Step 4: Check wall_cladding_items in context
    print("\n4️⃣ WALL CLADDING ITEMS IN CONTEXT")
    print("-" * 40)
    
    wall_cladding_items = context.get('wall_cladding_items', [])
    print(f"Number of wall_cladding_items: {len(wall_cladding_items)}")
    
    if wall_cladding_items:
        for idx, item in enumerate(wall_cladding_items):
            print(f"\n  Wall Cladding Item {idx + 1}:")
            print(f"    • item_number: {item.get('item_number', 'NOT_FOUND')}")
            print(f"    • description: {item.get('description', 'NOT_FOUND')}")
            print(f"    • width: {item.get('width', 'NOT_FOUND')}")
            print(f"    • height: {item.get('height', 'NOT_FOUND')}")
            print(f"    • dimensions: {item.get('dimensions', 'NOT_FOUND')}")
            print(f"    • position_description: {item.get('position_description', 'NOT_FOUND')}")
            print(f"    • level_name: {item.get('level_name', 'NOT_FOUND')}")
            print(f"    • area_name: {item.get('area_name', 'NOT_FOUND')}")
    else:
        print("  ❌ No wall_cladding_items found in context")
    
    # Step 5: Check pricing_totals for cladding
    print("\n5️⃣ PRICING TOTALS - CLADDING DATA")
    print("-" * 40)
    
    pricing_totals = context.get('pricing_totals', {})
    total_cladding_price = pricing_totals.get('total_cladding_price', 0)
    print(f"Total cladding price: £{total_cladding_price:,.2f}")
    
    areas = pricing_totals.get('areas', [])
    print(f"Number of areas in pricing_totals: {len(areas)}")
    
    for area in areas:
        area_name = area.get('level_area_combined', 'Unknown Area')
        cladding_total = area.get('cladding_total', 0)
        print(f"\n  📂 {area_name}:")
        print(f"    • cladding_total: £{cladding_total:,.2f}")
        
        canopies = area.get('canopies', [])
        print(f"    • Number of canopies: {len(canopies)}")
        
        for canopy in canopies:
            ref_num = canopy.get('reference_number', 'Unknown')
            cladding_price = canopy.get('cladding_price', 0)
            has_cladding = canopy.get('has_cladding', False)
            print(f"      🔹 {ref_num}: cladding_price=£{cladding_price:,.2f}, has_cladding={has_cladding}")
    
    # Step 6: Check enhanced levels for cladding
    print("\n6️⃣ ENHANCED LEVELS - CLADDING DATA")
    print("-" * 40)
    
    levels = context.get('levels', [])
    print(f"Number of levels in context: {len(levels)}")
    
    for level in levels:
        level_name = level.get('level_name', 'Unknown Level')
        print(f"\n  📁 {level_name}:")
        
        areas = level.get('areas', [])
        for area in areas:
            area_name = area.get('name', 'Unknown Area')
            print(f"    📂 {area_name}:")
            
            canopies = area.get('canopies', [])
            for canopy in canopies:
                ref_num = canopy.get('reference_number', 'Unknown')
                has_wall_cladding = canopy.get('has_wall_cladding', False)
                cladding_price = canopy.get('cladding_price', 0)
                wall_cladding = canopy.get('wall_cladding', {})
                
                print(f"      🔹 {ref_num}:")
                print(f"         • has_wall_cladding: {has_wall_cladding}")
                print(f"         • cladding_price: £{cladding_price:,.2f}")
                print(f"         • wall_cladding.type: {wall_cladding.get('type', 'NOT_FOUND')}")
                print(f"         • wall_cladding.width: {wall_cladding.get('width', 'NOT_FOUND')}")
                print(f"         • wall_cladding.height: {wall_cladding.get('height', 'NOT_FOUND')}")
                print(f"         • wall_cladding.position: {wall_cladding.get('position', 'NOT_FOUND')}")
    
    # Step 7: Save debug data to files
    print("\n7️⃣ SAVING DEBUG DATA")
    print("-" * 25)
    
    # Save raw project data
    with open('debug_cladding_raw_project.json', 'w') as f:
        json.dump(project_data, f, indent=2, default=str)
    print("✅ Saved raw project data to: debug_cladding_raw_project.json")
    
    # Save template context
    with open('debug_cladding_context.json', 'w') as f:
        json.dump(context, f, indent=2, default=str)
    print("✅ Saved template context to: debug_cladding_context.json")
    
    # Step 8: Summary and recommendations
    print("\n8️⃣ SUMMARY & RECOMMENDATIONS")
    print("-" * 35)
    
    if not wall_cladding_items:
        print("❌ ISSUE IDENTIFIED: wall_cladding_items is empty")
        print("\n🔧 POSSIBLE CAUSES:")
        print("   1. Cladding data not being read from Excel correctly")
        print("   2. Cladding detection logic in prepare_template_context() failing")
        print("   3. Wall cladding type is 'None' or empty")
        print("   4. Wall cladding dimensions are missing")
        
        print("\n💡 DEBUGGING STEPS:")
        print("   1. Check Excel file for cladding data in rows 19-24")
        print("   2. Verify dimensions format (e.g., '1000X2100')")
        print("   3. Verify position format (e.g., 'rear/left hand')")
        print("   4. Check if cladding price is in column N")
    
    if total_cladding_price == 0:
        print("❌ ISSUE IDENTIFIED: total_cladding_price is 0")
        print("\n🔧 POSSIBLE CAUSES:")
        print("   1. Cladding prices not being read from Excel")
        print("   2. Cladding prices not being summed correctly")
        print("   3. Canopy cladding_price field is 0 or missing")
    
    print(f"\n🎯 NEXT STEPS:")
    print("   1. Review the saved JSON files for detailed data")
    print("   2. Check Excel file structure for cladding data")
    print("   3. Verify Word template uses correct Jinja syntax")
    print("   4. Test with a simple cladding example")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 2:
        print("Usage: python debug_cladding_issue.py <excel_file_path>")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    debug_cladding_data_flow(excel_path) 