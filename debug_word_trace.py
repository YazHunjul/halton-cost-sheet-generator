#!/usr/bin/env python3
"""
Detailed trace script to follow data flow from input to Word template.
"""

import sys
import os
import json

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def trace_word_data_flow():
    """Trace the complete data flow to identify where values are lost."""
    print("=== Tracing Word Data Flow ===")
    
    # Step 1: Create input data
    test_data = {
        'customer': 'Test Customer',
        'company': 'Test Company',
        'address': 'Test Address',
        'project_name': 'Test Project',
        'location': 'Test Location',
        'project_number': 'TRACE-1223',
        'estimator': 'Test Estimator',
        'estimator_rank': 'Lead Estimator',
        'estimator_initials': 'TE',
        'date': '25/05/2024',
        'delivery_location': 'Test Delivery',
        'levels': [
            {
                'level_name': 'Level 1',
                'areas': [
                    {
                        'name': 'Main Kitchen',
                        'canopies': [
                            {
                                'reference_number': '1223',
                                'model': 'KVI',  # No 'F' - should have MUA and Supply as '-'
                                'configuration': 'Wall',
                                'length': '1000',
                                'width': '',  # Empty
                                'height': '555',
                                'sections': '',  # Empty
                                'lighting_type': '',  # Empty - should become '-'
                                'extract_volume': '',  # Empty - should become '-'
                                'extract_static': '',  # Empty - should become '-'
                                'mua_volume': '',  # Empty - should become '-' (because no F)
                                'supply_static': '',  # Empty - should become '-' (because no F)
                                'cws_capacity': '',
                                'hws_requirement': '',
                                'hw_storage': '',
                                'has_wash_capabilities': False,
                                'wall_cladding': {
                                    'type': 'None',
                                    'width': None,
                                    'height': None,
                                    'position': None
                                }
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    print("=== Step 1: Input Data ===")
    canopy = test_data['levels'][0]['areas'][0]['canopies'][0]
    print(f"Input canopy data:")
    print(f"  model: {repr(canopy.get('model'))}")
    print(f"  extract_static: {repr(canopy.get('extract_static'))}")
    print(f"  mua_volume: {repr(canopy.get('mua_volume'))}")
    print(f"  supply_static: {repr(canopy.get('supply_static'))}")
    print(f"  lighting_type: {repr(canopy.get('lighting_type'))}")
    
    try:
        from utils.word import handle_empty_value, transform_lighting_type
        
        print("\n=== Step 2: Testing Individual Functions ===")
        print(f"handle_empty_value(''): {repr(handle_empty_value(''))}")
        print(f"handle_empty_value('-'): {repr(handle_empty_value('-'))}")
        print(f"transform_lighting_type(''): {repr(transform_lighting_type(''))}")
        
        print("\n=== Step 3: Manual Business Logic Application ===")
        model = canopy.get('model', '').upper()
        print(f"Model: {repr(model)}")
        print(f"'F' in model: {'F' in model}")
        
        # Apply business rules manually
        if 'F' not in model:
            mua_volume = "-"
            supply_static = "-"
            print(f"No F in model -> mua_volume: {repr(mua_volume)}, supply_static: {repr(supply_static)}")
        else:
            mua_volume = handle_empty_value(canopy.get('mua_volume', ''))
            supply_static = handle_empty_value(canopy.get('supply_static', ''))
            print(f"Has F in model -> mua_volume: {repr(mua_volume)}, supply_static: {repr(supply_static)}")
        
        # Extract static logic
        if model in ['CMWF', 'CMWI']:
            extract_static = "-"
            print(f"CMWF/CMWI model -> extract_static: {repr(extract_static)}")
        else:
            extract_static = handle_empty_value(canopy.get('extract_static', ''))
            print(f"Other model -> extract_static: {repr(extract_static)}")
        
        # Lighting type
        lighting_type = transform_lighting_type(canopy.get('lighting_type', ''))
        print(f"lighting_type: {repr(lighting_type)}")
        
        print("\n=== Step 4: Apply handle_empty_value to processed values ===")
        # This is what the current code does after the business logic
        final_extract_static = handle_empty_value(extract_static)
        final_mua_volume = handle_empty_value(mua_volume)
        final_supply_static = handle_empty_value(supply_static)
        
        print(f"handle_empty_value(extract_static='{extract_static}'): {repr(final_extract_static)}")
        print(f"handle_empty_value(mua_volume='{mua_volume}'): {repr(final_mua_volume)}")
        print(f"handle_empty_value(supply_static='{supply_static}'): {repr(final_supply_static)}")
        
        print("\n=== Step 5: Full Template Context Processing ===")
        from utils.word import prepare_template_context
        
        context = prepare_template_context(test_data)
        
        # Check main canopies array
        canopies = context.get('canopies', [])
        if canopies:
            canopy_result = canopies[0]
            print(f"Main canopies[0] result:")
            print(f"  reference_number: {repr(canopy_result.get('reference_number'))}")
            print(f"  model: {repr(canopy_result.get('model'))}")
            print(f"  extract_static: {repr(canopy_result.get('extract_static'))}")
            print(f"  mua_volume: {repr(canopy_result.get('mua_volume'))}")
            print(f"  supply_static: {repr(canopy_result.get('supply_static'))}")
            print(f"  lighting_type: {repr(canopy_result.get('lighting_type'))}")
        
        # Check levels structure
        levels = context.get('levels', [])
        if levels and levels[0].get('areas') and levels[0]['areas'][0].get('canopies'):
            level_canopy = levels[0]['areas'][0]['canopies'][0]
            print(f"\nLevels[0].areas[0].canopies[0] result:")
            print(f"  reference_number: {repr(level_canopy.get('reference_number'))}")
            print(f"  model: {repr(level_canopy.get('model'))}")
            print(f"  extract_static: {repr(level_canopy.get('extract_static'))}")
            print(f"  mua_volume: {repr(level_canopy.get('mua_volume'))}")
            print(f"  supply_static: {repr(level_canopy.get('supply_static'))}")
            print(f"  lighting_type: {repr(level_canopy.get('lighting_type'))}")
        
        print("\n=== Step 6: Check for Data Consistency ===")
        if canopies and levels:
            main_canopy = canopies[0]
            level_canopy = levels[0]['areas'][0]['canopies'][0]
            
            fields_to_check = ['extract_static', 'mua_volume', 'supply_static', 'lighting_type']
            for field in fields_to_check:
                main_val = main_canopy.get(field)
                level_val = level_canopy.get(field)
                if main_val == level_val:
                    print(f"✅ {field}: Both structures have {repr(main_val)}")
                else:
                    print(f"❌ {field}: Main={repr(main_val)}, Level={repr(level_val)}")
        
        print("\n=== Step 7: Generate Word Document with Trace ===")
        from docxtpl import DocxTemplate
        
        template_path = "templates/word/Halton Quote Feb 2024.docx"
        if os.path.exists(template_path):
            # Create a simple context with just the problematic fields
            simple_context = {
                'client_name': 'TRACE TEST',
                'project_name': 'TRACE PROJECT',
                'canopies': [
                    {
                        'reference_number': '1223',
                        'model': 'KVI',
                        'extract_static': final_extract_static,
                        'mua_volume': final_mua_volume,
                        'supply_static': final_supply_static,
                        'lighting_type': lighting_type,
                        'extract_volume': '-',
                        'length': '1000',
                        'width': '-',
                        'height': '555',
                        'sections': '-',
                        'configuration': 'Wall'
                    }
                ],
                'levels': [
                    {
                        'level_name': 'Level 1',
                        'areas': [
                            {
                                'name': 'Main Kitchen',
                                'canopies': [
                                    {
                                        'reference_number': '1223',
                                        'model': 'KVI',
                                        'extract_static': final_extract_static,
                                        'mua_volume': final_mua_volume,
                                        'supply_static': final_supply_static,
                                        'lighting_type': lighting_type,
                                        'extract_volume': '-',
                                        'length': '1000',
                                        'width': '-',
                                        'height': '555',
                                        'sections': '-',
                                        'configuration': 'Wall'
                                    }
                                ]
                            }
                        ]
                    }
                ],
                'total_canopies': 1,
                'wall_cladding_items': []
            }
            
            print(f"Simple context canopy data:")
            print(f"  extract_static: {repr(simple_context['canopies'][0]['extract_static'])}")
            print(f"  mua_volume: {repr(simple_context['canopies'][0]['mua_volume'])}")
            print(f"  supply_static: {repr(simple_context['canopies'][0]['supply_static'])}")
            print(f"  lighting_type: {repr(simple_context['canopies'][0]['lighting_type'])}")
            
            # Save simple context to JSON
            with open('trace_simple_context.json', 'w') as f:
                json.dump(simple_context, f, indent=2)
            print(f"✅ Simple context saved to trace_simple_context.json")
            
            # Generate Word document
            doc = DocxTemplate(template_path)
            doc.render(simple_context)
            
            output_path = "output/trace_word_test.docx"
            os.makedirs("output", exist_ok=True)
            doc.save(output_path)
            print(f"✅ Trace Word document saved: {output_path}")
            
            print(f"\n=== Step 8: Instructions ===")
            print(f"1. Check trace_simple_context.json to see the exact data being passed")
            print(f"2. Open {output_path} to see what appears in the Word document")
            print(f"3. Compare the JSON values with what shows in the Word table")
            print(f"4. Look specifically for:")
            print(f"   - extract_static: Should show '{final_extract_static}'")
            print(f"   - mua_volume: Should show '{final_mua_volume}'")
            print(f"   - supply_static: Should show '{final_supply_static}'")
            print(f"   - lighting_type: Should show '{lighting_type}'")
        else:
            print(f"❌ Template not found: {template_path}")
        
    except Exception as e:
        print(f"❌ Error in trace: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    trace_word_data_flow()
    print("\n=== Trace completed ===") 