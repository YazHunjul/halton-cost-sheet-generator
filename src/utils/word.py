"""
Word document generation utilities for Halton quotation system.
Handles creation of quotation documents from Excel data using Jinja templating.
"""
from typing import Dict, List, Tuple, Union
import os
import zipfile
from datetime import datetime
from docxtpl import DocxTemplate
from config.business_data import SALES_CONTACTS, ESTIMATORS
from config.constants import is_feature_enabled
import streamlit as st
# Template path for Word documents
WORD_TEMPLATE_PATH = "templates/word/Halton Quote Feb 2024.docx"
RECOAIR_TEMPLATE_PATH = "templates/word/Halton RECO Quotation Jan 2025 (2).docx"



def get_fire_suppression_system_description(system_type: str) -> str:
    """
    Determine the fire suppression system description based on the system type from C16.
    
    Args:
        system_type (str): The system type value from C16 (NOBEL, AMAREX, or other)
        
    Returns:
        str: The appropriate system description
    """
    if not system_type:
        return 'Ansul R102 system. Supplied, installed & commissioned.'
    
    system_type_upper = str(system_type).upper().strip()
    
    if 'NOBEL' in system_type_upper:
        return 'NOBEL System. Supplied, installed & commissioned.'
    elif 'AMAREX' in system_type_upper:
        return 'AMAREX System. Supplied, installed & commissioned.'
    else:
        # For "FIRE SUPPRESSION", "1 TANK SYSTEM", or any other value, default to Ansul R102
        return 'Ansul R102 system. Supplied, installed & commissioned.'

def get_sales_contact_info(estimator_name: str, project_data: Dict = None) -> Dict[str, str]:
    """
    Get sales contact information based on project data or estimator name.
    
    Args:
        estimator_name (str): Name of the estimator (kept for backward compatibility)
        project_data (Dict, optional): Project data containing sales_contact
        
    Returns:
        Dict: Contact information including name and phone
    """
    # First, try to use sales_contact from project_data if available
    if project_data and project_data.get('sales_contact'):
        sales_contact_name = project_data['sales_contact']
        if sales_contact_name in SALES_CONTACTS:
            return {
                'name': sales_contact_name,
                'phone': SALES_CONTACTS[sales_contact_name]
            }
    
    # Fallback: try to match estimator to sales contact (old logic, likely won't match)
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
        if "DALI" in lighting_str:
            # Return specific LED STRIP type with DALI
            if "L6" in lighting_str:
                return "LED STRIP L6 Inc DALI"
            elif "L12" in lighting_str:
                return "LED STRIP L12 Inc DALI"
            elif "L18" in lighting_str:
                return "LED STRIP L18 Inc DALI"
            else:
                return "LED STRIP Inc DALI"
        elif "EM" in lighting_str:
            # Return specific LED STRIP type with EM
            if "L6" in lighting_str:
                return "LED STRIP L6EM"
            elif "L12" in lighting_str:
                return "LED STRIP L12EM"
            elif "L18" in lighting_str:
                return "LED STRIP L18EM"
            else:
                return "LED STRIP EM"
        else:
            # Return specific LED STRIP type without DALI/EM
            if "LM6" in lighting_str or "LM-6" in lighting_str:
                return "LM6"
            elif "LM12" in lighting_str or "LM-12" in lighting_str:
                return "LM12"
            elif "LM18" in lighting_str or "LM-18" in lighting_str:
                return "LM18"
            else:
                return "LED STRIP"
    
    # Check for spots variations
    elif "SPOTS" in lighting_str or "SPOT" in lighting_str:
        if "DALI" in lighting_str:
            if "SMALL" in lighting_str:
                return "Small LED Spots Inc DALI"
            elif "LARGE" in lighting_str:
                return "Large LED Spots Inc DALI"
            else:
                return "LED SPOTS Inc DALI"
        else:
            return "LED SPOTS"
    
    # Check for HCL variations
    elif lighting_str.startswith("HCL"):
        if "600" in lighting_str:
            return "HCL600 DALI"
        elif "1200" in lighting_str:
            return "HCL1200 DALI"
        elif "1800" in lighting_str:
            return "HCL1800 DALI"
        else:
            return "HCL DALI"
    
    # Check for EL variations
    elif lighting_str.startswith("EL"):
        if "215" in lighting_str:
            return "EL215"
        elif "218" in lighting_str:
            return "EL218"
        else:
            return "EL"
    
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
    Format MUA volume keeping the original value as-is without rounding.
    
    Args:
        value: MUA volume value to format
        
    Returns:
        str: Original MUA volume value or "-" if empty
    """
    if not value:
        return "-"
    
    # Convert to string and clean up
    str_value = str(value).strip()
    
    # Handle empty or dash values
    if str_value == "" or str_value == "-":
        return "-"
    
    # Return the original value as-is without any rounding
    return str_value

def get_combined_initials(sales_contact_name: str, estimator_name: str) -> str:
    """
    Generate combined initials from sales contact and estimator names.
    Format: Sales Contact Initials / Estimator Initials
    
    Args:
        sales_contact_name (str): Full name of sales contact
        estimator_name (str): Full name of estimator
        
    Returns:
        str: Combined initials (e.g., "YH/JS" for "Yazan Hunjul" + "Joe Salloum")
    """
    from utils.excel import get_initials
    
    sales_initials = get_initials(sales_contact_name) if sales_contact_name else ""
    estimator_initials = get_initials(estimator_name) if estimator_name else ""
    
    # Combine with '/' separator
    if sales_initials and estimator_initials:
        return f"{sales_initials}/{estimator_initials}"
    elif sales_initials:
        return sales_initials
    elif estimator_initials:
        return estimator_initials
    else:
        return ""

def get_customer_first_name(customer_name: str) -> str:
    """
    Extract the first name from a customer name.
    
    Args:
        customer_name (str): Full customer name
        
    Returns:
        str: First name of the customer
    """
    if not customer_name:
        return ""
    
    # Split by spaces and take the first part
    name_parts = customer_name.strip().split()
    if name_parts:
        return name_parts[0]
    else:
        return ""

def generate_reference_variable(project_number: str, sales_contact_name: str, estimator_name: str) -> str:
    """
    Generate reference variable in format: projectnumber/salesinitials/estimatorintials
    
    Args:
        project_number (str): Project number
        sales_contact_name (str): Full name of sales contact
        estimator_name (str): Full name of estimator
        
    Returns:
        str: Reference variable (e.g., "P12345/YH/JS")
    """
    from utils.excel import get_initials
    
    sales_initials = get_initials(sales_contact_name) if sales_contact_name else ""
    estimator_initials = get_initials(estimator_name) if estimator_name else ""
    
    # Format: projectnumber/salesinitials/estimatorintials
    return f"{project_number}/{sales_initials}/{estimator_initials}"

def generate_quote_title(revision: str) -> str:
    """
    Generate quote title based on revision.
    
    Args:
        revision (str): Project revision (A, B, C, etc.) or empty for initial version
        
    Returns:
        str: Quote title ("QUOTATION" for no revision or blank, "QUOTATION - REVISION A" for A and beyond)
    """
    # If no revision or empty/blank revision, just return "QUOTATION"
    if not revision or str(revision).strip() == "":
        return "QUOTATION"
    
    revision = str(revision).strip().upper()
    
    # For blank/empty revision (initial version), just return "QUOTATION"
    if revision == "":
        return "QUOTATION"
    
    # For any actual revision letter (A, B, C, etc.), include it in the title
    return f"QUOTATION - REVISION {revision}"

def normalize_area_object(area: Dict) -> Dict:
    """
    Ensure area object has all required attributes for template compatibility.
    
    Args:
        area (Dict): Area object to normalize
        
    Returns:
        Dict: Normalized area object with all required attributes
    """
    normalized = area.copy()
    
    # Ensure options exists and is a dictionary
    if 'options' not in normalized:
        normalized['options'] = {}
    elif normalized['options'] is None:
        normalized['options'] = {}
    
    # Ensure all expected option keys exist
    if not isinstance(normalized['options'], dict):
        normalized['options'] = {}
    
    option_defaults = {'uvc': False, 'sdu': False, 'recoair': False}
    for key, default_value in option_defaults.items():
        if key not in normalized['options']:
            normalized['options'][key] = default_value
    
    return normalized

def prepare_template_context(project_data: Dict, excel_file_path: str = None) -> Dict:
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
    sales_contact = get_sales_contact_info(estimator, project_data)
    
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
            # Normalize area object to ensure it has proper structure
            area = normalize_area_object(area)
            area_name = area.get('name', '')
            level_area_combined = f"{level_name} - {area_name}"
            
            # Check if any canopy in this area actually has fire suppression (tank quantity > 0 or price > 0)
            area_has_fire_suppression = any(
                (canopy.get('fire_suppression_tank_quantity', 0) > 0 or 
                 canopy.get('fire_suppression_price', 0) > 0)
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
                # CXW models are extract-only, so they never have supply static or MUA volume
                if model == "CXW":
                    mua_volume = "-"
                    supply_static = "-"
                # If canopy doesn't have 'F' in its name, set MUA volume and supply static to '-'
                elif 'F' not in model:
                    mua_volume = "-"
                    supply_static = "-"
                else:
                    mua_volume = format_mua_volume(canopy.get('mua_volume', ''))
                    
                    # Handle supply static for F models (excluding CXW)
                    raw_supply_static = canopy.get('supply_static', '')
                    
                    # Check if we should use the existing value or apply default
                    should_use_default = False
                    
                    if not raw_supply_static:
                        should_use_default = True
                    elif str(raw_supply_static).strip() == "":
                        should_use_default = True
                    else:
                        # Check if it's a very small value that should be treated as empty
                        try:
                            numeric_value = float(str(raw_supply_static).strip())
                            # Treat values less than 1 as empty (likely Excel calculation artifacts)
                            if numeric_value < 1:
                                should_use_default = True
                        except (ValueError, TypeError):
                            # If it's not a number, use the existing value
                            pass
                    
                    if should_use_default:
                        # Apply default value for F models
                        supply_static = "45"
                    else:
                        # Use existing value from Excel
                        supply_static = format_extract_static(raw_supply_static)
                
                # Extract static logic based on model type
                if model in ['CMWF', 'CMWI']:
                    # CMWF/CMWI models always show '-'
                    extract_static = "-"
                elif model == "CXW":
                    # CXW models always show '45'
                    extract_static = "45"
                else:
                    # KV, UV, and all other models get value from Excel
                    raw_extract_static = canopy.get('extract_static', '')
                    extract_static = format_extract_static(raw_extract_static)
                
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
                # Only include canopies that actually have fire suppression
                tank_quantity = canopy.get('fire_suppression_tank_quantity', 0)
                fire_suppression_price = canopy.get('fire_suppression_price', 0)
                
                # Only add to fire suppression items if this specific canopy has fire suppression
                if tank_quantity > 0 or fire_suppression_price > 0:
                    fire_suppression_item = {
                        'item_number': canopy.get('reference_number', ''),
                        'system_description': get_fire_suppression_system_description(canopy.get('fire_suppression_system_type', '')),
                        'fire_suppression_system_type': canopy.get('fire_suppression_system_type', ''),  # Add raw system type
                        'manual_release': '1no station',
                        'tank_quantity': tank_quantity if tank_quantity > 0 else 'TBD',  # Show TBD if not specified
                        'price': fire_suppression_price,  # Fire suppression price (includes base + commissioning share + delivery share)
                        'canopy_ref': canopy.get('reference_number', ''),
                        'level_name': level_name,
                        'area_name': area_name,
                        'level_area_combined': level_area_combined
                    }
                    fire_suppression_items.append(fire_suppression_item)
                
                # Create transformed canopy data
                fs_system_type_raw = canopy.get('fire_suppression_system_type', '')
                fs_system_desc = get_fire_suppression_system_description(fs_system_type_raw)
                
                # Handle sections - double the value if configuration is ISLAND
                raw_sections = canopy.get('sections', '')
                configuration = canopy.get('configuration', '').upper().strip()
                
                if raw_sections and raw_sections != '' and configuration == 'ISLAND':
                    try:
                        # Try to convert to number, double it, then convert back to string
                        sections_num = float(raw_sections)
                        doubled_sections = sections_num * 2
                        # If it's a whole number, display without decimal
                        if doubled_sections.is_integer():
                            display_sections = str(int(doubled_sections))
                        else:
                            display_sections = str(doubled_sections)
                    except (ValueError, TypeError):
                        # If conversion fails, just use the original value
                        display_sections = handle_empty_value(raw_sections)
                else:
                    # For WALL configuration or any other configuration, use original value
                    display_sections = handle_empty_value(raw_sections)
                
                transformed_canopy = {
                    'reference_number': handle_empty_value(canopy.get('reference_number', '')),
                    'model': handle_empty_value(canopy.get('model', '')),
                    'configuration': handle_empty_value(canopy.get('configuration', '')),
                    'length': handle_empty_value(canopy.get('length', '')),
                    'width': handle_empty_value(canopy.get('width', '')),
                    'height': handle_empty_value(canopy.get('height', '')),
                    'sections': display_sections,  # Use the processed sections value
                    'sections_raw': handle_empty_value(raw_sections),  # Keep original value for reference
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
                    'fire_suppression_system_type': handle_empty_value(canopy.get('fire_suppression_system_type', '')),
                    'fire_suppression_system_description': fs_system_desc,
                    
                    # Pricing data
                    'canopy_price': canopy.get('canopy_price', 0),
                    'fire_suppression_price': canopy.get('fire_suppression_price', 0),
                    'cladding_price': canopy.get('cladding_price', 0),
                    
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
                'canopies': transformed_canopies,  # Use transformed data instead of raw data
                
                # Area-level options
                'options': area.get('options', {}),
                
                # Area-level pricing data
                'delivery_installation_price': area.get('delivery_installation_price', 0),
                'commissioning_price': area.get('commissioning_price', 0),
                
                # Area-level option pricing
                'uvc_price': area.get('uvc_price', 0),
                'sdu_price': area.get('sdu_price', 0),
                'recoair_price': area.get('recoair_price', 0),
                
                # RecoAir units data (detailed unit specifications)
                'recoair_units': area.get('recoair_units', [])
            }
            enhanced_areas.append(enhanced_area)
        
        # Enhanced level with combined names in areas
        enhanced_level = {
            'level_name': level_name,
            'areas': enhanced_areas
        }
        enhanced_levels.append(enhanced_level)
    
    # Generate combined initials (Sales Contact / Estimator)
    combined_initials = get_combined_initials(sales_contact['name'], estimator)
    
    # Generate reference variable (projectnumber/salesinitials/estimatorintials)
    reference_variable = generate_reference_variable(
        project_data.get('project_number', ''), 
        sales_contact['name'], 
        estimator
    )
    
    # Generate quote title based on revision
    quote_title = generate_quote_title(project_data.get('revision', ''))
    
    # Extract customer first name
    customer_first_name = get_customer_first_name(project_data.get('customer', ''))
    
    # Collect RecoAir pricing data (areas and job totals)
    recoair_pricing_data = collect_recoair_pricing_schedule_data(project_data)
    
    # Collect SDU data for areas with SDU systems
    sdu_data = collect_sdu_data(project_data, excel_file_path)
    
    # Analyze project for global flags
    has_canopies, has_recoair, is_recoair_only, has_uv = analyze_project_areas(project_data)
    
    # Calculate pricing totals once
    pricing_totals = calculate_pricing_totals(project_data, excel_file_path)
    
    # Prepare the context
    context = {
        # Basic project information
        'client_name': project_data.get('customer', ''),  # Don't use handle_empty_value for customer name
        'customer_first_name': customer_first_name,  # Don't use handle_empty_value for customer first name
        'company': handle_empty_value(project_data.get('company', '')),
        'address': handle_empty_value(project_data.get('address', '')),
        'project_name': handle_empty_value(project_data.get('project_name', '')),
        'location': handle_empty_value(project_data.get('project_location') or project_data.get('location', '')),
        'project_number': handle_empty_value(project_data.get('project_number', '')),
        'estimator': estimator,  # Full name
        'estimator_rank': estimator_rank,  # Lead Estimator, Estimator, etc.
        'estimator_initials': combined_initials,  # Combined Sales Contact / Estimator initials
        'reference_variable': reference_variable,  # Project reference (projectnumber/salesinitials/estimatorintials)
        'quote_title': quote_title,  # Quote title based on revision (QUOTATION or QUOTATION - REVISION X)
        'revision': project_data.get('revision', ''),  # Project revision - keep blank for initial version
        
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
        'subject_line': f"{project_data.get('project_name', '')}, {project_data.get('project_location') or project_data.get('location', '')}",
        
        # Estimator with rank for signatures
        'estimator_with_rank': f"{estimator}\n{estimator_rank}" if estimator and estimator_rank else estimator,
        
        # Current date for any additional formatting needs
        'current_date': datetime.now().strftime('%d %B %Y'),
        'current_year': datetime.now().year,
        
        # Wall cladding data
        'wall_cladding_items': wall_cladding_items,
        
        # Fire suppression data
        'fire_suppression_items': fire_suppression_items,
        
        # Scope of works data
        'scope_of_works': generate_scope_of_works(project_data),
        
        # Pricing data
        'pricing_totals': pricing_totals,
        'recoair_pricing_schedules': recoair_pricing_data['areas'],  # RecoAir area-by-area pricing schedules
        'recoair_job_totals': recoair_pricing_data['job_totals'],  # RecoAir job totals
        'format_currency': format_currency,  # Make currency formatter available in templates
        'format_current': format_currency,  # Alias for format_currency (for template compatibility)
        
        # Individual pricing totals for template compatibility
        'total_canopy_price': pricing_totals.get('total_canopy_price', 0),
        'total_fire_suppression_price': pricing_totals.get('total_fire_suppression_price', 0),
        'total_cladding_price': pricing_totals.get('total_cladding_price', 0),
        'total_delivery_installation': pricing_totals.get('total_delivery_installation', 0),
        'total_commissioning': pricing_totals.get('total_commissioning', 0),
        'total_uvc_price': pricing_totals.get('total_uvc_price', 0),
        'total_sdu_price': pricing_totals.get('total_sdu_price', 0),
        'total_recoair_price': pricing_totals.get('total_recoair_price', 0),
        'project_total': pricing_totals.get('project_total', 0),
        
        # RecoAir-specific data (for RecoAir templates)
        'recoair_areas': [area for level in enhanced_levels for area in level.get('areas', []) if area.get('options', {}).get('recoair', False)],
        'total_recoair_units': sum(len(area.get('recoair_units', [])) for level in enhanced_levels for area in level.get('areas', [])),
        
        # SDU-specific data
        'sdu_areas': sdu_data,
        'has_sdu': len(sdu_data) > 0,
        'total_sdu_areas': len(sdu_data),
        
        # Global project flags
        'has_canopies': has_canopies,
        'has_recoair': has_recoair,
        'is_recoair_only': is_recoair_only,
        'has_uv': has_uv,
        
        # Feature flags for conditional display of systems
        'show_kitchen_extract_system': is_feature_enabled('kitchen_extract_system'),
        'show_kitchen_makeup_air_system': is_feature_enabled('kitchen_makeup_air_system'),
        'show_marvel_system': is_feature_enabled('marvel_system'),
        'show_cyclocell_cassette_ceiling': is_feature_enabled('cyclocell_cassette_ceiling'),
        'show_reactaway_unit': is_feature_enabled('reactaway_unit'),
        'show_dishwasher_extract': is_feature_enabled('dishwasher_extract'),
        'show_gas_interlocking': is_feature_enabled('gas_interlocking'),
        'show_pollustop_unit': is_feature_enabled('pollustop_unit'),
        
        # Fallback variables for template compatibility
        'level': enhanced_levels[0] if enhanced_levels else {'level_name': '', 'areas': []},  # First level as fallback
        'area': enhanced_levels[0].get('areas', [{}])[0] if enhanced_levels and enhanced_levels[0].get('areas') else {'name': '', 'recoair_units': [], 'options': {}},
        
        # RecoAir unit fallback variables
        'model': '',  # Fallback for RecoAir model
        'extract_volume': 0,  # Fallback for extract volume
        'width': 0,  # Fallback for width
        'length': 0,  # Fallback for length
        'height': 0,  # Fallback for height
        'recoair_location': 'INTERNAL',  # Fallback for RecoAir unit location
        'unit_price': 0,  # Fallback for unit price
        'quantity': 0,  # Fallback for quantity
        'delivery_installation_price': 0,  # Fallback for delivery price
        'total_uv_extra_over_cost': pricing_totals.get('total_uv_extra_over_cost', 0),  # Use calculated total UV Extra Over costs
        'has_any_uv_extra_over': pricing_totals.get('has_any_uv_extra_over', False),  # Use calculated UV Extra Over flag
        'extra_overs': area.get('options', {}).get('uv_extra_over', False),  # Easy flag for templates
    }
    
    return context

def analyze_project_areas(project_data: Dict) -> Tuple[bool, bool, bool, bool]:
    """
    Analyze project areas to determine what types of systems are present.
    
    Args:
        project_data (Dict): Project data with levels and areas
        
    Returns:
        Tuple[bool, bool, bool, bool]: (has_canopies, has_recoair, is_recoair_only, has_uv)
    """
    has_canopies = False
    has_recoair = False
    has_uv = False
    
    for level in project_data.get('levels', []):
        for area in level.get('areas', []):
            # Check if area has canopies (check length, not just existence)
            canopies = area.get('canopies', [])
            if len(canopies) > 0:
                has_canopies = True
            
            # Check for UV canopy models (UVI, UVF, etc.) across all canopies in the project
            for canopy in canopies:
                model = canopy.get('model', '').upper().strip()
                if model.startswith('UV'):  # UV models like UVI, UVF, etc.
                    has_uv = True
            
            # Check if area has RecoAir option
            if area.get('options', {}).get('recoair', False):
                has_recoair = True
    
    # Determine if project is RecoAir-only
    is_recoair_only = has_recoair and not has_canopies
    
    return has_canopies, has_recoair, is_recoair_only, has_uv

def generate_single_document(project_data: Dict, template_path: str, output_filename: str, excel_file_path: str = None) -> str:
    """
    Generate a single Word document from project data using specified template.
    
    Args:
        project_data (Dict): Project data extracted from Excel
        template_path (str): Path to the Word template
        output_filename (str): Name for the output file
        
    Returns:
        str: Path to the generated Word document
    """
    try:
        # Check if template exists
        if not os.path.exists(template_path):
            raise Exception(f"Template file not found at {template_path}")
        
        # Load the template
        doc = DocxTemplate(template_path)
        
        # Prepare the context for template rendering
        context = prepare_template_context(project_data, excel_file_path)
        
        # Render the template with the context
        doc.render(context)
        
        # Save the document
        output_path = f"output/{output_filename}"
        os.makedirs("output", exist_ok=True)
        doc.save(output_path)
        
        return output_path
        
    except Exception as e:
        raise Exception(f"Failed to generate Word document: {str(e)}")

def format_date_for_filename(date_str: str) -> str:
    """
    Format date for filename (remove slashes and make it filename-safe).
    
    Args:
        date_str (str): Date string from project data
        
    Returns:
        str: Formatted date string for filename
    """
    if date_str:
        # Convert DD/MM/YYYY to DDMMYYYY or similar format
        return date_str.replace('/', '').replace('-', '')
    else:
        # Use current date if no date provided
        return datetime.now().strftime("%d%m%Y")

def generate_quotation_document(project_data: Dict, excel_file_path: str = None) -> Union[str, str]:
    """
    Generate Word quotation document(s) from project data using Jinja templating.
    Returns either a single document path or a zip file path containing multiple documents.
    
    Args:
        project_data (Dict): Project data extracted from Excel
        
    Returns:
        str: Path to the generated Word document or zip file
    """
    try:
        # Analyze project to determine what documents to generate
        has_canopies, has_recoair, is_recoair_only, has_uv = analyze_project_areas(project_data)
        
        project_number = project_data.get('project_number', 'unknown')
        date_str = format_date_for_filename(project_data.get('date', ''))
        
        # Case 1: RecoAir-only project - generate only RecoAir quotation
        # Format: "Project Number RecoAir Quotation Date"
        if is_recoair_only:
            output_filename = f"{project_number} RecoAir Quotation {date_str}.docx"
            return generate_single_document(project_data, RECOAIR_TEMPLATE_PATH, output_filename, excel_file_path)
        
        # Case 2: Mixed project or canopy-only project
        documents_to_generate = []
        
        # Always generate main quotation if there are canopies
        # Format: "Project Number Quotation Date"
        if has_canopies:
            main_filename = f"{project_number} Quotation {date_str}.docx"
            documents_to_generate.append((WORD_TEMPLATE_PATH, main_filename, "Main Quotation"))
        
        # Generate RecoAir quotation if there are RecoAir areas
        # Format: "Project Number RecoAir Quotation Date"
        if has_recoair:
            recoair_filename = f"{project_number} RecoAir Quotation {date_str}.docx"
            documents_to_generate.append((RECOAIR_TEMPLATE_PATH, recoair_filename, "RecoAir Quotation"))
        
        # If only one document to generate, return it directly
        if len(documents_to_generate) == 1:
            template_path, filename, _ = documents_to_generate[0]
            return generate_single_document(project_data, template_path, filename, excel_file_path)
        
        # If multiple documents, generate all and create zip file
        if len(documents_to_generate) > 1:
            generated_files = []
            
            # Generate each document
            for template_path, filename, description in documents_to_generate:
                try:
                    file_path = generate_single_document(project_data, template_path, filename, excel_file_path)
                    generated_files.append((file_path, filename))
                except Exception as e:
                    print(f"Warning: Failed to generate {description}: {str(e)}")
                    continue
            
            # Create zip file if we have multiple documents
            if len(generated_files) > 1:
                zip_filename = f"{project_number} Quotations {date_str}.zip"
                zip_path = f"output/{zip_filename}"
                
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    # Add Word documents only
                    for file_path, filename in generated_files:
                        zipf.write(file_path, filename)
                
                return zip_path
            
            # If only one document was successfully generated, return it
            elif len(generated_files) == 1:
                return generated_files[0][0]
        
        # Fallback: generate main quotation only
        main_filename = f"{project_number} Quotation {date_str}.docx"
        return generate_single_document(project_data, WORD_TEMPLATE_PATH, main_filename, excel_file_path)
        
    except Exception as e:
        raise Exception(f"Failed to generate Word document(s): {str(e)}")

def collect_recoair_pricing_schedule_data(project_data: Dict) -> Dict:
    """
    Collect RecoAir pricing schedule data per area for Word document display.
    
    Args:
        project_data (Dict): Project data with levels and areas
        
    Returns:
        Dict: Dictionary containing area pricing schedules and job totals
    """
    pricing_schedules = []
    
    # Job totals tracking
    job_totals = {
        'total_units_price': 0,
        'total_delivery_price': 0,
        'total_commissioning_price': 0,
        'total_flat_pack_price': 0,
        'job_total': 0,
        'total_areas': 0,
        'total_units': 0
    }
    
    for level in project_data.get('levels', []):
        level_name = level.get('level_name', '')
        
        for area in level.get('areas', []):
            area_name = area.get('name', '')
            level_area_combined = f"{level_name} - {area_name}"
            
            # Only process areas that have RecoAir systems
            if not area.get('options', {}).get('recoair', False):
                continue
            
            # Get RecoAir units for this area
            recoair_units = area.get('recoair_units', [])
            
            if not recoair_units:
                continue
            
            # Collect RecoAir unit items for this area
            recoair_items = []
            recoair_units_full = []  # Full unit specifications
            area_units_total = 0
            total_delivery_price = 0
            
            for unit in recoair_units:
                reference_number = unit.get('item_reference', '')
                model = unit.get('model', '')
                base_unit_price = unit.get('base_unit_price', 0) or 0  # Use base price from N12
                delivery_price = unit.get('delivery_installation_price', 0) or 0
                
                if reference_number and model:  # Only include units with reference and model
                    # Basic pricing item (for backward compatibility)
                    recoair_item = {
                        'reference_number': reference_number,
                        'model': model,
                        'price': base_unit_price,  # Use base price from N12
                        'delivery_price': delivery_price
                    }
                    recoair_items.append(recoair_item)
                    
                    # Full unit specification (includes all technical details)
                    recoair_unit_full = {
                        'reference_number': reference_number,
                        'model': model,
                        'price': base_unit_price,  # Use base price from N12
                        'delivery_price': delivery_price,
                        # Technical specifications
                        'length': unit.get('length', 0),
                        'width': unit.get('width', 0),
                        'height': unit.get('height', 0),
                        'extract_volume': unit.get('extract_volume', 0),
                        'p_drop': unit.get('p_drop', 0),
                        'motor': unit.get('motor', 0),
                        'weight': unit.get('weight', 0),
                        'location': unit.get('location', 'INTERNAL'),
                        # Additional data
                        'quantity': unit.get('quantity', 1),
                        'model_original': unit.get('model_original', model),
                        'extract_volume_raw': unit.get('extract_volume_raw', ''),
                        'base_unit_price': base_unit_price,  # Keep base price for reference
                        'n29_addition': unit.get('n29_addition', 0)
                    }
                    recoair_units_full.append(recoair_unit_full)
                    
                    area_units_total += base_unit_price  # Use base price from N12
                    total_delivery_price += delivery_price
            
            # Get commissioning price (should be read from N46 in Excel)
            # For now, use area commissioning price or default to 0
            commissioning_price = area.get('recoair_commissioning_price', 0) or area.get('commissioning_price', 0)
            
            # Get flat pack data if available
            flat_pack_data = area.get('recoair_flat_pack', {})
            flat_pack_price = flat_pack_data.get('price', 0) if flat_pack_data.get('has_flat_pack', False) else 0
            
            # Calculate different subtotals for different purposes
            recoair_subtotal = area_units_total + total_delivery_price + commissioning_price  # RecoAir units + delivery + commissioning (excluding flat pack)
            area_total_with_flat_pack = recoair_subtotal + flat_pack_price  # Everything including flat pack
            
            # Only create pricing schedule if there are RecoAir units in this area
            if recoair_items:
                area_pricing = {
                    'level_name': level_name,
                    'area_name': area_name,
                    'level_area_combined': level_area_combined,
                    'recoair_items': recoair_items,  # Basic pricing items (for backward compatibility)
                    'units': recoair_units_full,  # Full unit specifications with technical details
                    'units_total': area_units_total,
                    'delivery_installation_price': total_delivery_price,
                    'commissioning_price': commissioning_price,
                    'flat_pack_price': flat_pack_price,
                    'flat_pack_description': flat_pack_data.get('description', ''),
                    'flat_pack_item_reference': flat_pack_data.get('item_reference', ''),
                    'has_flat_pack': flat_pack_data.get('has_flat_pack', False),
                    'area_subtotal': recoair_subtotal,  # RecoAir units + delivery + commissioning (excluding flat pack)
                    'area_total_with_flat_pack': area_total_with_flat_pack,  # Everything including flat pack
                    'unit_count': len(recoair_items)
                }
                pricing_schedules.append(area_pricing)
                
                # Add to job totals (without rounding)
                job_totals['total_units_price'] += area_units_total
                job_totals['total_delivery_price'] += total_delivery_price
                job_totals['total_commissioning_price'] += commissioning_price
                job_totals['total_flat_pack_price'] += flat_pack_price
                job_totals['total_areas'] += 1
                job_totals['total_units'] += len(recoair_items)
    
    # Calculate overall job total
    job_totals['job_total'] = (
        job_totals['total_units_price'] + 
        job_totals['total_delivery_price'] + 
        job_totals['total_commissioning_price'] + 
        job_totals['total_flat_pack_price']
    )
    
    return {
        'areas': pricing_schedules,
        'job_totals': job_totals
    }

def calculate_pricing_totals(project_data: Dict, excel_file_path: str = None) -> Dict:
    """
    Calculate comprehensive pricing totals for the project.
    
    Args:
        project_data (Dict): Project data with pricing information
        excel_file_path (str, optional): Path to Excel file to extract detailed SDU data
        
    Returns:
        Dict: Pricing totals and breakdowns
    """
    totals = {
        'total_canopies': 0,
        'total_canopy_price': 0,
        'total_fire_suppression_price': 0,
        'total_cladding_price': 0,
        'total_delivery_installation': 0,
        'total_commissioning': 0,
        'total_uvc_price': 0,
        'total_sdu_price': 0,
        'total_sdu_subtotal': 0,  # New: total of all SDU subtotals
        'total_recoair_price': 0,
        'total_uv_extra_over_cost': 0,  # New: total of all UV Extra Over costs
        'has_any_uv_extra_over': False,  # New: flag indicating if any area has UV Extra Over
        'areas': [],
        'project_total': 0
    }
    
    # Collect SDU data once for all areas
    sdu_data_by_area = {}
    sdu_areas = collect_sdu_data(project_data, excel_file_path)
    for sdu_area in sdu_areas:
        key = sdu_area['level_area_combined']
        sdu_data_by_area[key] = sdu_area
    
    for level in project_data.get('levels', []):
        for area in level.get('areas', []):
            # Skip areas that only have RecoAir systems (no canopies) - they belong in separate RecoAir quotation
            canopies = area.get('canopies', [])
            has_canopies = len(canopies) > 0
            has_recoair_only = area.get('options', {}).get('recoair', False) and not has_canopies
            
            if has_recoair_only:
                continue  # Skip RecoAir-only areas from main quotation
            
            level_area_combined = f"{level.get('level_name', '')} - {area.get('name', '')}"
            
            area_totals = {
                'level_name': level.get('level_name', ''),
                'area_name': area.get('name', ''),
                'level_area_combined': level_area_combined,
                'canopy_count': len(canopies),
                'has_canopies': len(canopies) > 0,  # Boolean flag for easy template checking
                'canopy_total': 0,
                'fire_suppression_total': 0,
                'cladding_total': 0,
                'delivery_installation_price': area.get('delivery_installation_price', 0),
                'commissioning_price': area.get('commissioning_price', 0),
                'uvc_price': area.get('uvc_price', 0),
                'sdu_price': area.get('sdu_price', 0),
                'recoair_price': area.get('recoair_price', 0),
                'canopy_schedule_subtotal': 0,  # Canopy total + delivery + commissioning (excluding fire suppression)
                'area_subtotal': 0,
                'canopies': [],
                
                # Add UV Extra Over support
                'has_uv_extra_over': area.get('options', {}).get('uv_extra_over', False),
                'uv_extra_over_cost': area.get('uv_extra_over_cost', 0),
                'extra_overs': area.get('options', {}).get('uv_extra_over', False),  # Easy flag for templates
                
                # Add SDU data if this area has SDU
                'has_sdu': area.get('options', {}).get('sdu', False),
                'sdu': sdu_data_by_area.get(level_area_combined, {})
            }
            
            # Calculate canopy and fire suppression totals for this area
            for canopy in canopies:
                canopy_price = canopy.get('canopy_price', 0) or 0
                fire_suppression_price = canopy.get('fire_suppression_price', 0) or 0
                cladding_price = canopy.get('cladding_price', 0) or 0
                
                area_totals['canopy_total'] += canopy_price
                area_totals['fire_suppression_total'] += fire_suppression_price
                area_totals['cladding_total'] += cladding_price
                
                # Add individual canopy pricing info
                area_totals['canopies'].append({
                    'reference_number': canopy.get('reference_number', ''),
                    'model': canopy.get('model', ''),
                    'canopy_price': canopy_price,
                    'fire_suppression_price': fire_suppression_price,
                    'cladding_price': cladding_price,
                    'fire_suppression_tank_quantity': canopy.get('fire_suppression_tank_quantity', 0),
                    'has_cladding': canopy.get('cladding_price', 0) > 0 or canopy.get('wall_cladding', {}).get('type') not in ['None', None]
                })
                
                # Add to project totals
                totals['total_canopies'] += 1
                totals['total_canopy_price'] += canopy_price
                totals['total_fire_suppression_price'] += fire_suppression_price
                totals['total_cladding_price'] += cladding_price
            
            # Calculate SDU subtotal from detailed pricing if available
            sdu_subtotal = 0
            sdu_data = sdu_data_by_area.get(level_area_combined, {})
            if sdu_data and sdu_data.get('pricing', {}):
                sdu_pricing = sdu_data['pricing']
                sdu_subtotal = (
                    sdu_pricing.get('final_carcass_price', 0) +
                    sdu_pricing.get('final_electrical_price', 0) +
                    (sdu_pricing.get('live_site_test_price', 0) if sdu_pricing.get('has_live_test', False) else 0)
                )
            
            # Add SDU subtotal to area totals
            area_totals['sdu_subtotal'] = sdu_subtotal
            
            # Calculate area subtotal (excluding RecoAir price - RecoAir has separate quotation)
            # Use SDU subtotal if available, otherwise fall back to sdu_price
            sdu_total_for_area = sdu_subtotal if sdu_subtotal > 0 else area_totals['sdu_price']
            
            area_totals['area_subtotal'] = (
                area_totals['canopy_total'] + 
                area_totals['fire_suppression_total'] + 
                area_totals['cladding_total'] +
                area_totals['delivery_installation_price'] + 
                area_totals['commissioning_price'] +
                area_totals['uvc_price'] +
                sdu_total_for_area
                # Note: recoair_price excluded - RecoAir systems have separate quotation
            )
            
            # Calculate canopy schedule subtotal (canopy total + delivery + commissioning, excluding fire suppression)
            area_totals['canopy_schedule_subtotal'] = (
                area_totals['canopy_total'] + 
                area_totals['delivery_installation_price'] + 
                area_totals['commissioning_price']
            )
            
            # Add to project totals
            totals['total_delivery_installation'] += area_totals['delivery_installation_price']
            totals['total_commissioning'] += area_totals['commissioning_price']
            totals['total_uvc_price'] += area_totals['uvc_price']
            totals['total_sdu_price'] += area_totals['sdu_price']
            totals['total_sdu_subtotal'] += sdu_subtotal  # Add SDU subtotal to project totals
            totals['total_recoair_price'] += area_totals['recoair_price']
            
            # Update UV Extra Over cost and flag
            if area.get('options', {}).get('uv_extra_over', False):
                totals['total_uv_extra_over_cost'] += area.get('uv_extra_over_cost', 0)
                totals['has_any_uv_extra_over'] = True
            
            totals['areas'].append(area_totals)
    
    # Calculate project total (excluding RecoAir price - RecoAir has separate quotation)
    # Use SDU subtotal if available, otherwise fall back to SDU price
    sdu_total_for_project = totals['total_sdu_subtotal'] if totals['total_sdu_subtotal'] > 0 else totals['total_sdu_price']
    
    totals['project_total'] = (
        totals['total_canopy_price'] + 
        totals['total_fire_suppression_price'] + 
        totals['total_cladding_price'] +
        totals['total_delivery_installation'] + 
        totals['total_commissioning'] +
        totals['total_uvc_price'] +
        sdu_total_for_project
        # Note: total_recoair_price excluded - RecoAir systems have separate quotation
    )
    
    return totals

def format_currency(amount) -> str:
    """
    Format currency amount for display with ceiling rounding.
    All amounts are rounded UP to the nearest whole pound.
    
    Args:
        amount: Currency amount to format
        
    Returns:
        str: Formatted currency string (always ends in .00)
    """
    if not amount:
        return "£0.00"
    
    try:
        import math
        # Round UP to the nearest whole number (ceiling)
        # Only round if this is a final total (amount ends in .99 or similar)
        float_amount = float(amount)
        decimal_part = float_amount - int(float_amount)
        
        if decimal_part > 0:
            # This is a final total that needs rounding
            rounded_amount = math.ceil(float_amount)
        else:
            # This is a component price that shouldn't be rounded
            rounded_amount = float_amount
            
        return f"£{rounded_amount:,.2f}"
    except (ValueError, TypeError):
        return "£0.00"

def generate_scope_of_works(project_data: Dict) -> List[Dict]:
    """
    Generate scope of works based on canopy models, counts, and lighting types.
    
    Args:
        project_data (Dict): Project data with canopy information
        
    Returns:
        List[Dict]: List of scope items with counts and descriptions
    """
    # Count canopies by model and lighting type across the entire project
    model_lighting_counts = {}
    areas_with_cladding = set()
    
    for level in project_data.get('levels', []):
        for area in level.get('areas', []):
            area_has_cladding = False
            
            for canopy in area.get('canopies', []):
                model = canopy.get('model', '').upper().strip()
                lighting_type = canopy.get('lighting_type', '').upper().strip()
                
                # Check if this canopy has wall cladding
                wall_cladding = canopy.get('wall_cladding', {})
                if wall_cladding.get('type') not in ['None', None, ''] and wall_cladding.get('type'):
                    area_has_cladding = True
                
                if model:
                    # Normalize lighting type - only include if it's a real lighting selection
                    if lighting_type and lighting_type not in ['-', 'NONE', 'LIGHT SELECTION', '']:
                        # Simplify lighting type names
                        if 'LED STRIP' in lighting_type:
                            lighting_normalized = 'LED STRIP'
                        elif 'SPOT' in lighting_type:
                            lighting_normalized = 'LED SPOTS'
                        else:
                            lighting_normalized = lighting_type
                    else:
                        lighting_normalized = None  # No lighting
                    
                    # Create key combining model and lighting
                    key = (model, lighting_normalized)
                    model_lighting_counts[key] = model_lighting_counts.get(key, 0) + 1
            
            # Track areas with cladding
            if area_has_cladding:
                level_name = level.get('level_name', '')
                area_name = area.get('name', '')
                areas_with_cladding.add(f"{level_name} - {area_name}")
    
    # Generate scope descriptions based on model types and lighting
    scope_items = []
    
    for (model, lighting), count in model_lighting_counts.items():
        count_str = f"{count}no"
        
        # Base description based on model type
        if model.startswith('CMW'):
            if 'F' in model:
                base_description = f"{count_str} Extract/Supply Canopy c/w Capture Jet Tech and Water Wash Function"
            else:
                base_description = f"{count_str} Extract Canopy c/w Capture Jet Tech and Water Wash Function"
        elif model.startswith('UV'):
            if 'F' in model:
                base_description = f"{count_str} Extract/Supply Canopies c/w Capture Jet Tech and UV-c Filtration"
            else:
                base_description = f"{count_str} Extract Canopies c/w Capture Jet Tech and UV-c Filtration"
        elif model.startswith('CXW'):
            base_description = f"{count_str} Condense Canopies c/w Extract"
        else:
            # Standard canopies (KV, etc)
            if 'F' in model:
                base_description = f"{count_str} Extract/Supply Canopies c/w Capture Jet Tech"
            else:
                base_description = f"{count_str} Extract Canopies c/w Capture Jet Tech"
        
        # Add lighting type if present
        if lighting:
            description = f"{base_description} with {lighting}"
        else:
            description = base_description
        
        scope_items.append({
            'model': model,
            'lighting': lighting,
            'count': count,
            'count_str': count_str,
            'description': description
        })
    
    # Add wall cladding areas if any exist
    if areas_with_cladding:
        cladding_count = len(areas_with_cladding)
        cladding_count_str = f"{cladding_count}no"
        
        scope_items.append({
            'model': 'CLADDING',
            'lighting': None,
            'count': cladding_count,
            'count_str': cladding_count_str,
            'description': f"{cladding_count_str} Areas with Stainless Steel Cladding",
            'areas': list(areas_with_cladding)  # Include list of areas for reference
        })
    
    # Sort by model name for consistent ordering (put cladding at the end)
    scope_items.sort(key=lambda x: (x['model'] == 'CLADDING', x['model'], x.get('lighting') or ''))
    
    return scope_items

def collect_sdu_data(project_data: Dict, excel_file_path: str = None) -> List[Dict]:
    """
    Collect SDU data from project areas that have SDU systems.
    Reads electrical services data from Excel sheets if available.
    
    Args:
        project_data (Dict): Project data with levels and areas
        excel_file_path (str, optional): Path to Excel file to extract detailed SDU data
        
    Returns:
        List[Dict]: List of SDU data for each area with SDU systems
    """
    sdu_areas = []
    
    try:
        # Import here to avoid circular imports
        from utils.excel import extract_sdu_electrical_services
        from openpyxl import load_workbook
        
        # Load Excel workbook if path is provided
        wb = None
        if excel_file_path and os.path.exists(excel_file_path):
            try:
                wb = load_workbook(excel_file_path, data_only=True)
            except Exception as e:
                print(f"Warning: Could not load Excel file for SDU data: {str(e)}")
        
        for level in project_data.get('levels', []):
            level_name = level.get('level_name', '')
            
            for area in level.get('areas', []):
                area_name = area.get('name', '')
                
                # Check if this area has SDU system
                if area.get('options', {}).get('sdu', False):
                    # Create basic SDU data structure
                    sdu_data = {
                        'level_name': level_name,
                        'area_name': area_name,
                        'level_area_combined': f"{level_name} - {area_name}",
                        'sdu_price': area.get('sdu_price', 0),
                        
                        # Default electrical services (will be updated if Excel data is available)
                        'electrical_services': {
                            'distribution_board': 0,
                            'single_phase_switched_spur': 0,
                            'three_phase_socket_outlet': 0,
                            'switched_socket_outlet': 0,
                            'emergency_knock_off': 0,
                            'ring_main_inc_2no_sso': 0
                        },
                        
                        # Gas services (now implemented)
                        'gas_services': {
                            'gas_manifold': 0,
                            'gas_connection_15mm': 0,
                            'gas_connection_20mm': 0,
                            'gas_connection_25mm': 0,
                            'gas_connection_32mm': 0,
                            'gas_solenoid_valve': 0
                        },
                        
                        # Water services (now implemented)
                        'water_services': {
                            'cws_manifold_22mm': 0,
                            'cws_manifold_15mm': 0,
                            'hws_manifold': 0,
                            'water_connection_15mm': 0,
                            'water_connection_22mm': 0,
                            'water_connection_28mm': 0
                        },
                        
                        # Pricing information (now implemented)
                        'pricing': {
                            'carcass_only_price': 0,
                            'electrical_mechanical_price': 0,
                            'live_site_test_price': 0,
                            'delivery_price': 0,
                            'final_carcass_price': 0,
                            'final_electrical_price': 0,
                            'has_live_test': False
                        }
                    }
                    
                    # Try to extract electrical services data from Excel if workbook is available
                    if wb:
                        # Find the corresponding SDU sheet
                        sdu_sheet_name = None
                        for sheet_name in wb.sheetnames:
                            if 'SDU - ' in sheet_name and level_name in sheet_name:
                                # Check if this sheet corresponds to this area by checking the title in B1
                                try:
                                    sheet = wb[sheet_name]
                                    title = sheet['B1'].value
                                    if title and area_name in title:
                                        sdu_sheet_name = sheet_name
                                        break
                                except:
                                    continue
                        
                        if sdu_sheet_name:
                            try:
                                sdu_sheet = wb[sdu_sheet_name]
                                services_data = extract_sdu_electrical_services(sdu_sheet)
                                # Update electrical, gas, water services, and pricing
                                sdu_data['electrical_services'] = services_data.get('electrical_services', sdu_data['electrical_services'])
                                sdu_data['gas_services'] = services_data.get('gas_services', sdu_data['gas_services'])
                                sdu_data['water_services'] = services_data.get('water_services', sdu_data['water_services'])
                                sdu_data['pricing'] = services_data.get('pricing', sdu_data['pricing'])
                                print(f"Extracted services for {level_name} - {area_name}:")
                                print(f"  Electrical: {services_data.get('electrical_services', {})}")
                                print(f"  Gas: {services_data.get('gas_services', {})}")
                                print(f"  Water: {services_data.get('water_services', {})}")
                                print(f"  Pricing: {services_data.get('pricing', {})}")
                            except Exception as e:
                                print(f"Warning: Could not extract services from {sdu_sheet_name}: {str(e)}")
                    
                    sdu_areas.append(sdu_data)
        
        # Close workbook if we opened it
        if wb:
            wb.close()
    
    except Exception as e:
        print(f"Warning: Could not collect SDU data: {str(e)}")
    
    return sdu_areas 