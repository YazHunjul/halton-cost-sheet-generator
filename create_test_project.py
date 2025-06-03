#!/usr/bin/env python3

"""
Create a test project with UV Extra Over options to test area matching
"""

import sys
import os
sys.path.append('src')

from utils.excel import save_to_excel

def create_test_project():
    print("🔧 Creating test project with UV Extra Over options...")
    
    # Create a simple test project
    project_data = {
        'company_name': 'Test Company',
        'contact_name': 'Test Contact',
        'project_name': 'UV Test Project',
        'reference_number': 'UV001',
        'date_of_quote': '03/06/2025',
        'estimator': 'Test Estimator',
        'sales_contact': 'Test Sales',
        'organization': 'External',
        'delivery_location': 'Brighton',
        'levels': [
            {
                'level_name': 'Level 1',
                'level_number': 1,
                'areas': [
                    {
                        'name': 'Main Kitchen',
                        'options': {
                            'uv_extra_over': True  # Enable UV Extra Over for this area
                        },
                        'canopies': [
                            {
                                'reference_number': 'C001',
                                'model': 'UVI',  # Use UV model instead of KVT
                                'canopy_price': 1000,
                                'fire_suppression_price': 0,
                                'cladding_price': 0,
                                'fire_suppression_tank_quantity': 0,
                                'has_cladding': False,
                                'has_uv': True,  # Enable UV option
                                'uv_option': 'Yes'  # Set UV option
                            }
                        ]
                    }
                ]
            },
            {
                'level_name': 'Level 2', 
                'level_number': 2,
                'areas': [
                    {
                        'name': 'Bakery',
                        'options': {
                            'uv_extra_over': True  # Enable UV Extra Over for this area
                        },
                        'canopies': [
                            {
                                'reference_number': 'C002',
                                'model': 'UVF',  # Use UV model instead of KVT
                                'canopy_price': 800,
                                'fire_suppression_price': 0,
                                'cladding_price': 0,
                                'fire_suppression_tank_quantity': 0,
                                'has_cladding': False,
                                'has_uv': True,  # Enable UV option
                                'uv_option': 'Yes'  # Set UV option
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    try:
        # Generate the Excel file
        excel_path = save_to_excel(project_data)
        print(f"✅ Created Excel file: {excel_path}")
        return excel_path
        
    except Exception as e:
        print(f"❌ Error creating Excel file: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    create_test_project() 