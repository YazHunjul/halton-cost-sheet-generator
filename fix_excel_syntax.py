#!/usr/bin/env python3
"""
Script to fix syntax errors in excel.py by removing duplicate code.
"""

import re

def fix_excel_syntax():
    """Fix syntax errors in excel.py."""
    
    # Read the file
    with open('src/utils/excel.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find the end of the create_uv_extra_over_calculations_sheet function
    # and remove everything after it that looks like duplicate code
    pattern = r'(def create_uv_extra_over_calculations_sheet.*?except Exception as e:\s*print\(f"Warning: Could not create UV Extra Over calculations sheet: \{str\(e\)\}"\)\s*pass)\s*# Extract UV Extra Over.*$'
    
    match = re.search(pattern, content, re.DOTALL)
    if match:
        # Keep only the function definition, remove the duplicate code
        fixed_content = content[:match.end(1)]
        
        # Write the fixed content back
        with open('src/utils/excel.py', 'w', encoding='utf-8') as f:
            f.write(fixed_content)
        
        print("✅ Fixed syntax errors in excel.py")
        return True
    else:
        print("❌ Could not find the duplicate code pattern")
        return False

if __name__ == "__main__":
    fix_excel_syntax() 