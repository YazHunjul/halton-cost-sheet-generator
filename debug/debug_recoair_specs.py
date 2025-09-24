#!/usr/bin/env python3
"""Debug RecoAir specifications lookup."""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from utils.excel import transform_recoair_model, get_recoair_specifications

print("=" * 80)
print("DEBUG RECOAIR SPECIFICATIONS LOOKUP")
print("=" * 80)

# Test the models we found in the Excel
test_models = [
    "RA3.0 STANDARD (Prem Controls)",  # Sheet 1, Row 20
    "RA1.0 STANDARD",                  # Sheet 2, Row 16
    "RA4.0 STANDARD (Prem Controls)"   # Sheet 3, Row 22
]

for original_model in test_models:
    print(f"\nTesting model: '{original_model}'")
    
    try:
        # Test transformation
        transformed = transform_recoair_model(original_model)
        print(f"  Transformed to: '{transformed}'")
        
        # Test specs lookup
        specs = get_recoair_specifications(transformed)
        print(f"  Specifications: {specs}")
        
        # Test accessing individual keys
        p_drop = specs['p_drop']
        motor = specs['motor'] 
        weight = specs['weight']
        print(f"  Individual values: p_drop={p_drop}, motor={motor}, weight={weight}")
        
        print(f"  ✅ SUCCESS - No exceptions")
        
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 80)