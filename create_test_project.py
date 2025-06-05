#!/usr/bin/env python3

"""
Create a test project with pricing data to test new N column pricing storage
"""

import sys
import os
sys.path.append('src')

from utils.excel import save_to_excel

def create_test_project():
    print("🔧 Creating test project with pricing data to test N column storage...")
    
    # Create a test project with comprehensive pricing data
    project_data = {
        'company_name': 'Test Company',
        'contact_name': 'Test Contact',
        'project_name': 'Pricing Test Project',
        'reference_number': 'P001',
        'date_of_quote': '06/06/2025',
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
                        'delivery_installation_price': 500.00,  # Test delivery pricing in N182
                        'commissioning_price': 300.00,  # Test commissioning pricing in N183
                        'options': {
                            'fire_suppression': True,
                            'uvc': False,
                            'sdu': False,
                            'recoair': False
                        },
                        'canopies': [
                            {
                                'reference_number': 'C001',
                                'model': 'KVT',
                                'width': 2000,  # Test width in E14
                                'length': 1500,  # Test length in F14  
                                'height': 500,   # Test height in G14
                                'sections': 3,   # Test sections in H14
                                'canopy_price': 1200.00,  # Test canopy pricing in N12
                                'fire_suppression_price': 800.00,  # Test fire suppression pricing in N13
                                'cladding_price': 150.00,  # Test cladding pricing in N14
                                'fire_suppression_tank_quantity': 1,
                                'options': {
                                    'fire_suppression': True
                                }
                            },
                            {
                                'reference_number': 'C002',
                                'model': 'KVI',
                                'width': 2500,  # Test width in E31
                                'length': 2000,  # Test length in F31
                                'height': 600,   # Test height in G31
                                'sections': 4,   # Test sections in H31
                                'canopy_price': 1500.00,  # Test canopy pricing in N29
                                'fire_suppression_price': 900.00,  # Test fire suppression pricing in N30
                                'cladding_price': 200.00,  # Test cladding pricing in N31
                                'fire_suppression_tank_quantity': 1,
                                'options': {
                                    'fire_suppression': True
                                }
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
                        'delivery_installation_price': 300.00,  # Test delivery pricing in N182
                        'commissioning_price': 200.00,  # Test commissioning pricing in N183
                        'options': {
                            'fire_suppression': False,
                            'uvc': True,
                            'sdu': True,
                            'recoair': False
                        },
                        'canopies': [
                            {
                                'reference_number': 'C003',
                                'model': 'UVF',
                                'width': 1800,  # Test width in E14 (second sheet)
                                'length': 1200,  # Test length in F14 (second sheet)
                                'height': 400,   # Test height in G14 (second sheet)
                                'sections': 2,   # Test sections in H14 (second sheet)
                                'canopy_price': 800.00,  # Test canopy pricing in N12 (second sheet)
                                'fire_suppression_price': 0.00,
                                'cladding_price': 100.00,  # Test cladding pricing in N14 (second sheet)
                                'fire_suppression_tank_quantity': 0,
                                'options': {
                                    'fire_suppression': False
                                }
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
        
        # Test reading the pricing data back
        from utils.excel import read_excel_project_data
        read_data = read_excel_project_data(excel_path)
        
        print("\n📊 Testing pricing data storage and retrieval:")
        
        # Check Level 1 - Main Kitchen pricing
        level1_areas = read_data.get('levels', [{}])[0].get('areas', [])
        if level1_areas:
            main_kitchen = level1_areas[0]
            print(f"Level 1 - Main Kitchen:")
            print(f"  Delivery/Installation: £{main_kitchen.get('delivery_installation_price', 0)}")
            print(f"  Commissioning: £{main_kitchen.get('commissioning_price', 0)}")
            
            canopies = main_kitchen.get('canopies', [])
            for i, canopy in enumerate(canopies):
                print(f"  Canopy {i+1} ({canopy.get('reference_number', 'Unknown')}):")
                print(f"    Canopy Price: £{canopy.get('canopy_price', 0)}")
                print(f"    Fire Suppression Price: £{canopy.get('fire_suppression_price', 0)}")
                print(f"    Cladding Price: £{canopy.get('cladding_price', 0)}")
        
        # Check Level 2 - Bakery pricing
        if len(read_data.get('levels', [])) > 1:
            level2_areas = read_data['levels'][1].get('areas', [])
            if level2_areas:
                bakery = level2_areas[0]
                print(f"\nLevel 2 - Bakery:")
                print(f"  Delivery/Installation: £{bakery.get('delivery_installation_price', 0)}")
                print(f"  Commissioning: £{bakery.get('commissioning_price', 0)}")
                
                canopies = bakery.get('canopies', [])
                for i, canopy in enumerate(canopies):
                    print(f"  Canopy {i+1} ({canopy.get('reference_number', 'Unknown')}):")
                    print(f"    Canopy Price: £{canopy.get('canopy_price', 0)}")
                    print(f"    Fire Suppression Price: £{canopy.get('fire_suppression_price', 0)}")
                    print(f"    Cladding Price: £{canopy.get('cladding_price', 0)}")
        
        print(f"\n✅ Pricing test completed successfully!")
        return excel_path
        
    except Exception as e:
        print(f"❌ Error creating or testing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    create_test_project() 