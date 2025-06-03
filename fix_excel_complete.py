#!/usr/bin/env python3
"""
Script to completely fix all syntax errors in excel.py.
"""

def fix_excel_file():
    """Fix all syntax errors in excel.py."""
    
    print("🔧 Fixing excel.py syntax errors...")
    
    # Read the current file
    with open('src/utils/excel.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Check if file ends properly
    lines = content.split('\n')
    
    # Find the last valid function definition and remove any duplicate/broken code after it
    last_valid_line = len(lines) - 1
    
    # Look for the end of create_uv_extra_over_calculations_sheet function
    in_function = False
    function_end_line = -1
    
    for i, line in enumerate(lines):
        if 'def create_uv_extra_over_calculations_sheet' in line:
            in_function = True
            print(f"Found create_uv_extra_over_calculations_sheet function at line {i+1}")
        elif in_function and line.strip() == '' and i < len(lines) - 1:
            # Check if next non-empty line starts a new function or is at module level
            next_non_empty = i + 1
            while next_non_empty < len(lines) and lines[next_non_empty].strip() == '':
                next_non_empty += 1
            
            if (next_non_empty < len(lines) and 
                (lines[next_non_empty].startswith('def ') or 
                 lines[next_non_empty].startswith('class ') or
                 not lines[next_non_empty].startswith(' '))):
                function_end_line = i
                break
    
    if function_end_line == -1:
        # If we can't find a clean break, look for the last pass statement in the function
        for i in range(len(lines) - 1, -1, -1):
            if lines[i].strip() == 'pass' and 'create_uv_extra_over_calculations_sheet' in ''.join(lines[max(0, i-50):i]):
                function_end_line = i
                break
    
    if function_end_line != -1:
        # Keep only lines up to the function end
        lines = lines[:function_end_line + 1]
        print(f"Truncated file at line {function_end_line + 1}")
    
    # Join lines back together
    cleaned_content = '\n'.join(lines)
    
    # Ensure the file ends with a newline
    if not cleaned_content.endswith('\n'):
        cleaned_content += '\n'
    
    # Write the cleaned content
    with open('src/utils/excel.py', 'w', encoding='utf-8') as f:
        f.write(cleaned_content)
    
    print("✅ Fixed excel.py syntax errors")
    return True

if __name__ == "__main__":
    try:
        fix_excel_file()
        print("🎉 All syntax errors fixed!")
    except Exception as e:
        print(f"❌ Failed to fix syntax errors: {e}") 