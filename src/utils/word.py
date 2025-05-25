"""
Word document generation utilities for HVAC quotation system.
Handles creation of quotation documents from Excel data using Jinja templating.
"""
from typing import Dict, List
import os
from datetime import datetime
from docxtpl import DocxTemplate
from config.business_data import SALES_CONTACTS, ESTIMATORS
import streamlit as st
# Template path for Word documents
WORD_TEMPLATE_PATH = "templates/word/Halton Quote Feb 2024.docx"

def get_sales_contact_info(estimator_name: str) -> Dict[str, str]:
    """
    Get sales contact information based on estimator name.
    
    Args:
        estimator_name (str): Name of the estimator
        
    Returns:
        Dict: Contact information including name and phone
    """
    # Try to match estimator to sales contact
    for contact_name, phone in SALES_CONTACTS.items():
        if estimator_name and any(name.lower() in estimator_name.lower() for name in contact_name.split()):
            return {
                'name': contact_name,
                'phone': phone
            }
    
    # Default to first sales contact if no match
    first_contact = list(SALES_CONTACTS.items())[0]
    return {
        'name': first_contact[0],
        'phone': first_contact[1]
    }

def format_halton_reference(project_number: str, date: str) -> str:
    """
    Format the Halton reference number.
    
    Args:
        project_number (str): Project number
        date (str): Project date
        
    Returns:
        str: Formatted Halton reference
    """
    try:
        if isinstance(date, str) and '/' in date:
            # Extract year from date (assume format DD/MM/YYYY)
            year = date.split('/')[-1][-2:]  # Get last 2 digits of year
        else:
            year = str(datetime.now().year)[-2:]
        
        # Format as project_number/month/year
        month = datetime.now().strftime("%m")
        return f"{project_number}/{month}/{year}"
    except:
        return f"{project_number}/XX/XX"

def format_date_for_display(date_str: str) -> str:
    """
    Format date for display in the document.
    
    Args:
        date_str (str): Date string from Excel
        
    Returns:
        str: Formatted date string
    """
    try:
        if isinstance(date_str, str) and '/' in date_str:
            # Convert DD/MM/YYYY to DD Month YYYY
            parts = date_str.split('/')
            if len(parts) == 3:
                day, month, year = parts
                month_names = [
                    'January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'
                ]
                month_name = month_names[int(month) - 1] if 1 <= int(month) <= 12 else month
                return f"{day} {month_name} {year}"
        return date_str or datetime.now().strftime('%d %B %Y')
    except:
        return datetime.now().strftime('%d %B %Y')

def transform_lighting_type(lighting_type: str) -> str:
    """
    Transform lighting type for Word document display.
    
    Args:
        lighting_type (str): Raw lighting type from Excel
        
    Returns:
        str: Transformed lighting type
    """
    # Handle None, empty string, or whitespace-only strings
    if not lighting_type or str(lighting_type).strip() == "":
        return "-"
    
    lighting_str = str(lighting_type).upper().strip()
    
    # Handle "LIGHT SELECTION" as empty value
    if lighting_str == "LIGHT SELECTION":
        return "-"
    
    # Check for LED strip variations (any string containing "LED STRIP")
    if "LED STRIP" in lighting_str:
        return "LED STRIP"
    # Check for spots variations
    elif "SPOTS" in lighting_str or "SPOT" in lighting_str:
        return "LED SPOTS"
    else:
        # For any other value or unrecognized lighting type, return "-"
        return "-"

def handle_empty_value(value) -> str:
    """
    Convert empty values to '-' for Word document display.
    
    Args:
        value: Any value that might be empty
        
    Returns:
        str: The value or '-' if empty
    """
    # Handle None values
    if value is None:
        return "-"
    
    # Convert to string and check for empty or whitespace-only strings
    str_value = str(value).strip()
    if str_value == "" or str_value.lower() == "none":
        return "-"
    
    return str_value

def format_extract_static(value) -> str:
    """
    Format extract static value by removing 'Pa' and returning clean number.
    
    Args:
        value: Extract static value that might contain 'Pa'
        
    Returns:
        str: Formatted extract static value without 'Pa'
    """
    if not value:
        return "-"
    
    # Convert to string and clean up
    str_value = str(value).strip()
    
    # Handle empty or dash values
    if str_value == "" or str_value == "-":
        return "-"
    
    # Remove 'Pa' (case insensitive) and any extra whitespace
    cleaned_value = str_value.replace("Pa", "").replace("pa", "").replace("PA", "").strip()
    
    # If nothing left after removing Pa, return dash
    if not cleaned_value:
        return "-"
    
    return cleaned_value

def format_mua_volume(value) -> str:
    """
    Format MUA volume by rounding to 1 decimal place.
    
    Args:
        value: MUA volume value to format
        
    Returns:
        str: Formatted MUA volume rounded to 1 decimal place
    """
    if not value:
        return "-"
    
    # Convert to string and clean up
    str_value = str(value).strip()
    
    # Handle empty or dash values
    if str_value == "" or str_value == "-":
        return "-"
    
    try:
        # Try to convert to float and round to 1 decimal place
        float_value = float(str_value)
        return f"{float_value:.1f}"
    except (ValueError, TypeError):
        # If conversion fails, return the original value
        return str_value

def prepare_template_context(project_data: Dict) -> Dict:
    """
    Prepare the context dictionary for Jinja template rendering.
    
    Args:
        project_data (Dict): Project data from Excel
        
    Returns:
        Dict: Context dictionary for template rendering
    """
    # Get full estimator name (not initials)
    estimator = project_data.get('estimator', '')
    estimator_rank = project_data.get('estimator_rank', 'Estimator')
    
    # Get sales contact info
    sales_contact = get_sales_contact_info(estimator)
    
    # Prepare all canopies data with level-area combinations
    all_canopies = []
    enhanced_levels = []
    wall_cladding_items = []  # Collect all wall cladding data
    fire_suppression_items = []  # Collect all fire suppression data
    
    # Check if fire suppression sheets exist by looking for any areas with fire suppression data
    has_fire_suppression_sheets = False
    
    for level in project_data.get('levels', []):
        level_name = level.get('level_name', '')
        enhanced_areas = []
        
        for area in level.get('areas', []):
            area_name = area.get('name', '')
            level_area_combined = f"{level_name} - {area_name}"
            
            # Check if any canopy in this area has fire suppression data (even if tank quantity is 0)
            area_has_fire_suppression = any(
                canopy.get('fire_suppression_tank_quantity', 0) >= 0 and 
                'fire_suppression_tank_quantity' in canopy
                for canopy in area.get('canopies', [])
            )
            
            if area_has_fire_suppression:
                has_fire_suppression_sheets = True
            
            # Process canopies for this area and create transformed versions
            transformed_canopies = []
            
            # Add to all canopies with enhanced info
            for canopy in area.get('canopies', []):
                # Get model for business logic
                model = canopy.get('model', '').upper()
                
                # Apply business rules for volume and static data
                # If canopy doesn't have 'F' in its name, set MUA volume and supply static to '-'
                if 'F' not in model:
                    mua_volume = "-"
                    supply_static = "-"
                else:
                    mua_volume = format_mua_volume(canopy.get('mua_volume', ''))
                    supply_static = format_extract_static(canopy.get('supply_static', ''))
                
                # For extract static: if it's CMWF/CMWI, set it as '-'
                if model in ['CMWF', 'CMWI']:
                    extract_static = "-"
                else:
                    extract_static = format_extract_static(canopy.get('extract_static', ''))
                
                # Check for wall cladding on this canopy
                wall_cladding = canopy.get('wall_cladding', {})
                if wall_cladding.get('type') != 'None' and wall_cladding.get('type'):
                    # Handle position as list or string
                    position = wall_cladding.get('position', [])
                    if isinstance(position, list):
                        position_list = position if position else []
                    else:
                        position_list = [position] if position else []
                    
                    # Create proper description based on number of positions
                    if len(position_list) == 0:
                        description = "Cladding to walls"
                    elif len(position_list) == 1:
                        description = f"Cladding to {position_list[0]} walls"
                    elif len(position_list) == 2:
                        description = f"Cladding to {position_list[0]} and {position_list[1]} walls"
                    else:
                        # For 3 or more positions: "item1, item2 and item3 walls"
                        description = f"Cladding to {', '.join(position_list[:-1])} and {position_list[-1]} walls"
                    
                    # Join positions for other uses
                    position_str = "/".join(position_list) if position_list else ""
                    
                    wall_cladding_item = {
                        'item_number': canopy.get('reference_number', ''),  # Use canopy reference number
                        'description': description,
                        'width': wall_cladding.get('width', 0),
                        'height': wall_cladding.get('height', 0),
                        'dimensions': f"{wall_cladding.get('width', 0)}X{wall_cladding.get('height', 0)}",
                        'position_description': position_str,
                        'canopy_ref': canopy.get('reference_number', ''),
                        'level_name': level_name,
                        'area_name': area_name,
                        'level_area_combined': level_area_combined
                    }
                    wall_cladding_items.append(wall_cladding_item)
                
                # Check for fire suppression on this canopy
                # If fire suppression sheets exist, include all canopies from those areas
                if area_has_fire_suppression:
                    tank_quantity = canopy.get('fire_suppression_tank_quantity', 0)
                    # Include fire suppression item even if tank quantity is 0 (sheet exists but not filled)
                    fire_suppression_item = {
                        'item_number': canopy.get('reference_number', ''),
                        'system_description': 'Ansul R 102 System',
                        'manual_release': '1no station',
                        'tank_quantity': tank_quantity if tank_quantity > 0 else 'TBD',  # Show TBD if not specified
                        'canopy_ref': canopy.get('reference_number', ''),
                        'level_name': level_name,
                        'area_name': area_name,
                        'level_area_combined': level_area_combined
                    }
                    fire_suppression_items.append(fire_suppression_item)
                
                # Create transformed canopy data
                transformed_canopy = {
                    'reference_number': handle_empty_value(canopy.get('reference_number', '')),
                    'model': handle_empty_value(canopy.get('model', '')),
                    'configuration': handle_empty_value(canopy.get('configuration', '')),
                    'length': handle_empty_value(canopy.get('length', '')),
                    'width': handle_empty_value(canopy.get('width', '')),
                    'height': handle_empty_value(canopy.get('height', '')),
                    'sections': handle_empty_value(canopy.get('sections', '')),
                    'lighting_type': transform_lighting_type(canopy.get('lighting_type', '')),
                    'extract_volume': format_mua_volume(canopy.get('extract_volume', '')),
                    'extract_static': extract_static,
                    'mua_volume': mua_volume, 
                    'supply_static': supply_static,
                    
                    # CWS/HWS data for wash canopies
                    'cws_capacity': handle_empty_value(canopy.get('cws_capacity', '')),
                    'hws_requirement': handle_empty_value(canopy.get('hws_requirement', '')),
                    'hw_storage': handle_empty_value(canopy.get('hw_storage', '')),
                    'has_wash_capabilities': canopy.get('has_wash_capabilities', False),
                    
                    # Fire suppression data
                    'fire_suppression_tank_quantity': canopy.get('fire_suppression_tank_quantity', 0),
                    
                    # Wall cladding data for this canopy
                    'has_wall_cladding': wall_cladding.get('type') != 'None' and wall_cladding.get('type'),
                    'wall_cladding': wall_cladding,
                }
                # Add to transformed canopies for this area
                transformed_canopies.append(transformed_canopy)
                
                # Create canopy info for the main canopies array (with additional location info)
                canopy_info = {
                    'level': level_name,
                    'area': area_name,
                    'level_area_combined': level_area_combined,
                    'location': f"{level_name} - {area_name}",  # Keep this for backwards compatibility
                    **transformed_canopy  # Include all the transformed data
                }
                all_canopies.append(canopy_info)
            
            # Enhanced area with transformed canopy data
            enhanced_area = {
                'name': area_name,
                'level_area_name': level_area_combined,
                'canopies': transformed_canopies  # Use transformed data instead of raw data
            }
            enhanced_areas.append(enhanced_area)
        
        # Enhanced level with combined names in areas
        enhanced_level = {
            'level_name': level_name,
            'areas': enhanced_areas
        }
        enhanced_levels.append(enhanced_level)
    
    # Prepare the context
    context = {
        # Basic project information
        'client_name': handle_empty_value(project_data.get('customer', '')),
        'company': handle_empty_value(project_data.get('company', '')),
        'address': handle_empty_value(project_data.get('address', '')),
        'project_name': handle_empty_value(project_data.get('project_name', '')),
        'location': handle_empty_value(project_data.get('location', '')),
        'project_number': handle_empty_value(project_data.get('project_number', '')),
        'estimator': estimator,  # Full name
        'estimator_rank': estimator_rank,  # Lead Estimator, Estimator, etc.
        'estimator_initials': handle_empty_value(project_data.get('estimator_initials', '')),  # For any places that need initials
        
        # Formatted data
        'date': format_date_for_display(project_data.get('date', '')),
        'halton_ref': format_halton_reference(project_data.get('project_number', ''), project_data.get('date', '')),
        
        # Sales contact information
        'sales_contact_name': sales_contact['name'],
        'sales_contact_phone': sales_contact['phone'],
        
        # Additional business data
        'delivery_location': handle_empty_value(project_data.get('delivery_location', '')),
        
        # Project structure (enhanced with combined names)
        'levels': enhanced_levels,
        'canopies': all_canopies,
        'total_canopies': len(all_canopies),
        
        # Derived information
        'dear_line': f"{project_data.get('customer', '')}," if project_data.get('customer') else "Sir/Madam,",
        'subject_line': f"{project_data.get('project_name', '')}, {project_data.get('location', '')}",
        
        # Estimator with rank for signatures
        'estimator_with_rank': f"{estimator}\n{estimator_rank}" if estimator and estimator_rank else estimator,
        
        # Current date for any additional formatting needs
        'current_date': datetime.now().strftime('%d %B %Y'),
        'current_year': datetime.now().year,
        
        # Wall cladding data
        'wall_cladding_items': wall_cladding_items,
        
        # Fire suppression data
        'fire_suppression_items': fire_suppression_items
    }
    
    return context

def generate_quotation_document(project_data: Dict) -> str:
    """
    Generate a Word quotation document from project data using Jinja templating.
    
    Args:
        project_data (Dict): Project data extracted from Excel
        
    Returns:
        str: Path to the generated Word document
    """
    try:
        # Check if template exists
        if not os.path.exists(WORD_TEMPLATE_PATH):
            raise Exception(f"Template file not found at {WORD_TEMPLATE_PATH}")
        
        # Load the template
        doc = DocxTemplate(WORD_TEMPLATE_PATH)
        
        # Prepare the context for template rendering
        context = prepare_template_context(project_data)
        
        # Render the template with the context
        doc.render(context)
        
        # Save the document
        output_path = f"output/quotation_{project_data.get('project_number', 'unknown')}.docx"
        os.makedirs("output", exist_ok=True)
        doc.save(output_path)
        
        return output_path
        
    except Exception as e:
        raise Exception(f"Failed to generate Word document: {str(e)}") 