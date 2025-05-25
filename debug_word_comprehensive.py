#!/usr/bin/env python3
"""
Comprehensive debug script to identify Word template issues.
"""

import sys
import os
import json

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def debug_word_comprehensive():
    """Comprehensive debugging of Word template processing."""
    print("=== Comprehensive Word Template Debug ===")
    
    # Create test data that matches your image exactly
    test_data = {
        'customer': 'Test Customer',
        'company': 'Test Company',
        'address': 'Test Address',
        'project_name': 'Test Project',
        'location': 'Test Location',
        'project_number': 'DEBUG-1223',
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
    
    try:
        from utils.word import prepare_template_context
        from docxtpl import DocxTemplate
        
        # Step 1: Get the processed context
        print("=== Step 1: Processing Template Context ===")
        context = prepare_template_context(test_data)
        
        # Step 2: Save full context to JSON for inspection
        print("=== Step 2: Saving Context to JSON ===")
        debug_context = {}
        for key, value in context.items():
            try:
                json.dumps(value)
                debug_context[key] = value
            except (TypeError, ValueError):
                debug_context[key] = str(value)
        
        with open('debug_full_context.json', 'w') as f:
            json.dump(debug_context, f, indent=2, default=str)
        print(f"✅ Full context saved to debug_full_context.json")
        
        # Step 3: Extract and examine canopy data specifically
        print("=== Step 3: Examining Canopy Data ===")
        canopies = context.get('canopies', [])
        levels = context.get('levels', [])
        
        print(f"Total canopies in main array: {len(canopies)}")
        print(f"Total levels: {len(levels)}")
        
        # Check main canopies array
        if canopies:
            canopy = canopies[0]
            print(f"\nMain canopies[0] data:")
            print(f"  reference_number: {repr(canopy.get('reference_number'))}")
            print(f"  model: {repr(canopy.get('model'))}")
            print(f"  lighting_type: {repr(canopy.get('lighting_type'))}")
            print(f"  extract_volume: {repr(canopy.get('extract_volume'))}")
            print(f"  extract_static: {repr(canopy.get('extract_static'))}")
            print(f"  mua_volume: {repr(canopy.get('mua_volume'))}")
            print(f"  supply_static: {repr(canopy.get('supply_static'))}")
        
        # Check levels structure
        if levels:
            level = levels[0]
            areas = level.get('areas', [])
            if areas:
                area = areas[0]
                area_canopies = area.get('canopies', [])
                if area_canopies:
                    area_canopy = area_canopies[0]
                    print(f"\nLevels[0].areas[0].canopies[0] data:")
                    print(f"  reference_number: {repr(area_canopy.get('reference_number'))}")
                    print(f"  model: {repr(area_canopy.get('model'))}")
                    print(f"  lighting_type: {repr(area_canopy.get('lighting_type'))}")
                    print(f"  extract_volume: {repr(area_canopy.get('extract_volume'))}")
                    print(f"  extract_static: {repr(area_canopy.get('extract_static'))}")
                    print(f"  mua_volume: {repr(area_canopy.get('mua_volume'))}")
                    print(f"  supply_static: {repr(area_canopy.get('supply_static'))}")
        
        # Step 4: Test different placeholder values
        print("\n=== Step 4: Testing Different Placeholder Values ===")
        
        template_path = "templates/word/Halton Quote Feb 2024.docx"
        if not os.path.exists(template_path):
            print(f"❌ Template not found: {template_path}")
            return
        
        # Test 1: Original context with "-" values
        print("Test 1: Original context with '-' values")
        doc1 = DocxTemplate(template_path)
        doc1.render(context)
        output1 = "output/debug_dash_original.docx"
        os.makedirs("output", exist_ok=True)
        doc1.save(output1)
        print(f"✅ Saved: {output1}")
        
        # Test 2: Replace all "-" with "DASH_TEST"
        print("Test 2: Replace '-' with 'DASH_TEST'")
        context_dash_test = json.loads(json.dumps(debug_context, default=str))
        
        def replace_dashes(obj):
            if isinstance(obj, dict):
                return {k: replace_dashes(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [replace_dashes(item) for item in obj]
            elif obj == "-":
                return "DASH_TEST"
            else:
                return obj
        
        context_dash_test = replace_dashes(context_dash_test)
        
        doc2 = DocxTemplate(template_path)
        doc2.render(context_dash_test)
        output2 = "output/debug_dash_test.docx"
        doc2.save(output2)
        print(f"✅ Saved: {output2}")
        
        # Test 3: Replace all "-" with "—" (em dash)
        print("Test 3: Replace '-' with '—' (em dash)")
        context_em_dash = json.loads(json.dumps(debug_context, default=str))
        
        def replace_with_em_dash(obj):
            if isinstance(obj, dict):
                return {k: replace_with_em_dash(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [replace_with_em_dash(item) for item in obj]
            elif obj == "-":
                return "—"
            else:
                return obj
        
        context_em_dash = replace_with_em_dash(context_em_dash)
        
        doc3 = DocxTemplate(template_path)
        doc3.render(context_em_dash)
        output3 = "output/debug_em_dash.docx"
        doc3.save(output3)
        print(f"✅ Saved: {output3}")
        
        # Test 4: Use explicit values for the problematic fields
        print("Test 4: Explicit values for problematic fields")
        context_explicit = json.loads(json.dumps(debug_context, default=str))
        
        # Update canopies with explicit values
        for canopy in context_explicit.get('canopies', []):
            canopy['extract_static'] = 'EXTRACT_STATIC_TEST'
            canopy['mua_volume'] = 'MUA_VOLUME_TEST'
            canopy['supply_static'] = 'SUPPLY_STATIC_TEST'
            canopy['lighting_type'] = 'LIGHTING_TYPE_TEST'
        
        # Update levels structure too
        for level in context_explicit.get('levels', []):
            for area in level.get('areas', []):
                for canopy in area.get('canopies', []):
                    canopy['extract_static'] = 'EXTRACT_STATIC_TEST'
                    canopy['mua_volume'] = 'MUA_VOLUME_TEST'
                    canopy['supply_static'] = 'SUPPLY_STATIC_TEST'
                    canopy['lighting_type'] = 'LIGHTING_TYPE_TEST'
        
        doc4 = DocxTemplate(template_path)
        doc4.render(context_explicit)
        output4 = "output/debug_explicit_values.docx"
        doc4.save(output4)
        print(f"✅ Saved: {output4}")
        
        # Step 5: Summary and instructions
        print("\n=== Step 5: Summary and Instructions ===")
        print("Generated 4 test documents:")
        print(f"1. {output1} - Original with '-' values")
        print(f"2. {output2} - With 'DASH_TEST' instead of '-'")
        print(f"3. {output3} - With '—' (em dash) instead of '-'")
        print(f"4. {output4} - With explicit test values")
        print()
        print("Please check each document and report:")
        print("- Which values appear in the canopy table")
        print("- Which values are missing/empty")
        print("- Any patterns you notice")
        print()
        print("Also check debug_full_context.json to see the exact data being passed to the template.")
        
        # Save a summary of what we expect to see
        summary = {
            "expected_for_kvi_model": {
                "reference_number": "1223",
                "model": "KVI",
                "lighting_type": "-",
                "extract_volume": "-",
                "extract_static": "-",
                "mua_volume": "-",
                "supply_static": "-"
            },
            "business_rules": {
                "kvi_no_f": "MUA Volume and Supply Static should be '-'",
                "empty_values": "All empty values should become '-'",
                "extract_static": "Should come from G23 cell in Excel"
            }
        }
        
        with open('debug_expected_values.json', 'w') as f:
            json.dump(summary, f, indent=2)
        print("✅ Expected values saved to debug_expected_values.json")
        
    except Exception as e:
        print(f"❌ Error in comprehensive debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_word_comprehensive()
    print("\n=== Comprehensive debug completed ===") 