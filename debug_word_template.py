#!/usr/bin/env python3
"""
Debug script to examine Word template processing.
"""

import sys
import os
import json

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def debug_word_template():
    """Debug the Word template processing to see what data is being passed."""
    print("=== Debugging Word Template Processing ===")
    
    # Create simple test data with known values
    test_data = {
        'customer': 'Debug Customer',
        'company': 'Debug Company',
        'address': 'Debug Address',
        'project_name': 'Debug Project',
        'location': 'Debug Location',
        'project_number': 'DEBUG-001',
        'estimator': 'Debug Estimator',
        'estimator_rank': 'Lead Estimator',
        'estimator_initials': 'DE',
        'date': '25/05/2024',
        'delivery_location': 'Debug Delivery',
        'levels': [
            {
                'level_name': 'Debug Floor',
                'areas': [
                    {
                        'name': 'Debug Kitchen',
                        'canopies': [
                            {
                                'reference_number': 'DEBUG-1',
                                'model': 'UVF',
                                'configuration': 'Wall',
                                'length': '1000',
                                'width': '2',
                                'height': '555',
                                'sections': '1',
                                'lighting_type': 'LED STRIP L12 inc DALI',  # Should become "LED STRIP"
                                'extract_volume': '1.3',
                                'extract_static': '',  # Should become "-"
                                'mua_volume': '',  # Should become "-"
                                'supply_static': '',  # Should become "-"
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
    
    try:
        from utils.word import prepare_template_context
        
        # Get the processed context
        context = prepare_template_context(test_data)
        
        print("=== Template Context Debug ===")
        print(f"Total canopies: {len(context.get('canopies', []))}")
        
        for i, canopy in enumerate(context.get('canopies', [])):
            print(f"\nCanopy {i+1}:")
            print(f"  reference_number: {repr(canopy.get('reference_number'))}")
            print(f"  model: {repr(canopy.get('model'))}")
            print(f"  lighting_type: {repr(canopy.get('lighting_type'))}")
            print(f"  extract_volume: {repr(canopy.get('extract_volume'))}")
            print(f"  extract_static: {repr(canopy.get('extract_static'))}")
            print(f"  mua_volume: {repr(canopy.get('mua_volume'))}")
            print(f"  supply_static: {repr(canopy.get('supply_static'))}")
            
            # Check the actual type and value
            print(f"  lighting_type type: {type(canopy.get('lighting_type'))}")
            print(f"  extract_static type: {type(canopy.get('extract_static'))}")
            print(f"  mua_volume type: {type(canopy.get('mua_volume'))}")
            print(f"  supply_static type: {type(canopy.get('supply_static'))}")
            
            # Check if they equal the dash character
            print(f"  lighting_type == '-': {canopy.get('lighting_type') == '-'}")
            print(f"  extract_static == '-': {canopy.get('extract_static') == '-'}")
            print(f"  mua_volume == '-': {canopy.get('mua_volume') == '-'}")
            print(f"  supply_static == '-': {canopy.get('supply_static') == '-'}")
        
        # Save the context to a JSON file for inspection
        debug_context = {}
        for key, value in context.items():
            try:
                # Try to serialize the value
                json.dumps(value)
                debug_context[key] = value
            except (TypeError, ValueError):
                # If it can't be serialized, convert to string
                debug_context[key] = str(value)
        
        with open('debug_context.json', 'w') as f:
            json.dump(debug_context, f, indent=2, default=str)
        print(f"\n✅ Context saved to debug_context.json for inspection")
        
        # Now try to load the Word template and see what happens
        print("\n=== Testing Word Template Loading ===")
        from docxtpl import DocxTemplate
        
        template_path = "templates/word/Halton Quote Feb 2024.docx"
        if os.path.exists(template_path):
            print(f"✅ Template found: {template_path}")
            
            try:
                doc = DocxTemplate(template_path)
                print("✅ Template loaded successfully")
                
                # Try to render with our context
                doc.render(context)
                print("✅ Template rendered successfully")
                
                # Save the debug document
                debug_output = "debug_word_output.docx"
                doc.save(debug_output)
                print(f"✅ Debug document saved: {debug_output}")
                print(f"Full path: {os.path.abspath(debug_output)}")
                
            except Exception as e:
                print(f"❌ Error with template: {e}")
                import traceback
                traceback.print_exc()
        else:
            print(f"❌ Template not found: {template_path}")
        
    except Exception as e:
        print(f"❌ Error in debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_word_template()
    print("\n=== Debug completed ===") 