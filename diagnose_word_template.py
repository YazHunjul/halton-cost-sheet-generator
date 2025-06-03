#!/usr/bin/env python3
"""
Script to diagnose Word template syntax errors.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from docxtpl import DocxTemplate
from jinja2 import Template
import zipfile

def check_template_syntax(template_path):
    """Check a Word template for Jinja2 syntax errors."""
    
    print(f"\n🔍 Checking template: {template_path}")
    
    try:
        # Try to load the template
        template = DocxTemplate(template_path)
        
        # Extract the XML content to check for syntax errors
        with zipfile.ZipFile(template_path, 'r') as zip_file:
            # Check the main document part
            if 'word/document.xml' in zip_file.namelist():
                content = zip_file.read('word/document.xml').decode('utf-8')
                
                # Try to create a Jinja2 template from the content
                try:
                    jinja_template = Template(content)
                    print(f"   ✅ Template syntax is valid")
                    return True
                except Exception as e:
                    print(f"   ❌ Template syntax error: {str(e)}")
                    
                    # Try to find the problematic line
                    lines = content.split('\n')
                    print(f"   📝 Template has {len(lines)} lines")
                    
                    # Look for common Jinja2 patterns around line 625
                    for i, line in enumerate(lines[620:630], start=621):
                        if any(char in line for char in ['{', '}', '(', ')']):
                            print(f"   Line {i}: {line[:100]}...")
                    
                    return False
        
    except Exception as e:
        print(f"   ❌ Could not load template: {str(e)}")
        return False

def main():
    """Check all Word templates for syntax errors."""
    
    print("=== Word Template Syntax Diagnosis ===")
    
    templates = [
        "templates/word/Halton Quote Feb 2024.docx",
        "templates/word/Halton RECO Quotation Jan 2025 (2).docx"
    ]
    
    for template_path in templates:
        if os.path.exists(template_path):
            check_template_syntax(template_path)
        else:
            print(f"\n❌ Template not found: {template_path}")
    
    print("\n=== Recommendation ===")
    print("If a template has syntax errors, you'll need to:")
    print("1. Open the Word template file")
    print("2. Find the Jinja2 expression around line 625")  
    print("3. Fix the syntax (likely a missing parenthesis or extra brace)")
    print("4. Save the template")

if __name__ == "__main__":
    main() 