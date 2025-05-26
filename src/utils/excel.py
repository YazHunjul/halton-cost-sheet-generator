"""
Excel generation utilities for Halton quotation system.
Handles creation and manipulation of Excel workbooks based on templates.
"""
from typing import Dict, List, Union, Optional
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from config.business_data import VALID_CANOPY_MODELS

# Constants for Excel operations
TEMPLATE_PATH = "templates/excel/Halton Cost Sheet Jan 2025.xlsx"
BASE_SHEET_NAME = "CANOPY"  # The template sheet to copy from
RECOAIR_SHEET_NAME = "RECOAIR"
EDGE_BOX_SHEET_NAME = "EDGE BOX"
FIRE_SUPPRESSION_SHEET_NAME = "FIRE SUPPRESSION"  # Template sheet name
LISTS_SHEET_NAME = "Lists"

# Output sheet name mapping
OUTPUT_SHEET_NAMES = {
    FIRE_SUPPRESSION_SHEET_NAME: "FIRE SUPP"  # Map template name to output name
}

# Cell mappings for different data points (CANOPY, FIRE SUPP, JOB TOTAL, etc.)
CELL_MAPPINGS = {
    "project_number": "C3",  # Job No
    "customer": "C5",        # Customer
    "estimator": "C7",       # Sales Manager / Estimator Initials
    "project_name": "G3",    # Project Name
    "project_location": "G5",        # Project Location (was "location")
    "date": "G7",           # Date
    "revision": "O7",       # Revision
}

# Row spacing for canopy entries
CANOPY_ROW_SPACING = 17

# Starting row for canopy data
CANOPY_START_ROW = 14  # First canopy starts at row 14

# Tab color mapping for different levels
TAB_COLORS = [
    "FF92D050",  # Light green
    "FF00B0F0",  # Light blue
    "FFFF9900",  # Orange
    "FFFF00FF",  # Pink
    "FF7030A0",  # Purple
    "FFFF0000",  # Red
    "FF00FF00",  # Green
    "FF0070C0",  # Blue
    "FFFFC000",  # Gold
    "FF00FFFF",  # Cyan
]

def load_template_workbook() -> Workbook:
    """
    Load the master Excel template workbook.
    
    Returns:
        Workbook: The loaded template workbook
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template file not found at {TEMPLATE_PATH}")
    
    try:
        return load_workbook(TEMPLATE_PATH)
    except Exception as e:
        raise Exception(f"Failed to load template workbook: {str(e)}")

def copy_template_sheet(wb: Workbook, sheet_name: str, new_name: str) -> Worksheet:
    """
    Create a copy of a template sheet with a new name.
    
    Args:
        wb (Workbook): The workbook to modify
        sheet_name (str): Name of the template sheet to copy
        new_name (str): Name for the new sheet
    
    Returns:
        Worksheet: The newly created worksheet
    """
    try:
        template_sheet = wb[sheet_name]
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = new_name
        return new_sheet
    except KeyError:
        raise Exception(f"Template sheet '{sheet_name}' not found in workbook. Available sheets: {wb.sheetnames}")
    except Exception as e:
        raise Exception(f"Failed to copy template sheet: {str(e)}")

def get_sheet_type(project_type: str, canopy_data: Optional[Dict] = None) -> List[str]:
    """
    Determine which template sheets to use based on project and canopy type.
    
    Args:
        project_type (str): Type of project (Canopy or RecoAir)
        canopy_data (Dict, optional): Canopy configuration data
    
    Returns:
        List[str]: List of template sheet names to use
    """
    sheets = []
    
    if project_type == "RecoAir Project":
        sheets.append(RECOAIR_SHEET_NAME)
    elif project_type == "Canopy Project":
        # Check if canopy has specific requirements for EDGE BOX
        if canopy_data and canopy_data.get("configuration") == "Edge":
            sheets.append(EDGE_BOX_SHEET_NAME)
        else:
            sheets.append(BASE_SHEET_NAME)
        
        # Add Fire Suppression sheet if the option is enabled
        if canopy_data and canopy_data.get("options", {}).get("fire_suppression"):
            sheets.append(FIRE_SUPPRESSION_SHEET_NAME)
    else:
        raise ValueError(f"Unknown project type: {project_type}")
    
    return sheets

def get_output_sheet_name(template_sheet_name: str) -> str:
    """
    Get the output sheet name for a given template sheet name.
    
    Args:
        template_sheet_name (str): The name of the template sheet
    
    Returns:
        str: The name to use in the output file
    """
    return OUTPUT_SHEET_NAMES.get(template_sheet_name, template_sheet_name)

def get_initials(name_str: str) -> str:
    """
    Convert name string into initials. Handles multiple names separated by 'and' or '/'.
    Example: "Yazan Hunjul / Joe Salloum" -> "YH/JS"
    
    Args:
        name_str (str): Name string potentially containing multiple names
        
    Returns:
        str: Initials with slash separator
    """
    if not name_str:
        return ""
    
    # Split by common separators (and, /, &)
    names = [n.strip() for n in name_str.replace(" and ", "/").replace("&", "/").split("/")]
    
    # Get initials for each name
    initials_list = []
    for name in names:
        parts = name.strip().split()
        if parts:
            initials = ''.join(part[0].upper() for part in parts if part)
            initials_list.append(initials)
    
    # Join with slash
    return '/'.join(initials_list)

def read_wall_cladding_from_canopy(sheet: Worksheet, base_row: int) -> Dict:
    """
    Read wall cladding data from canopy row in Excel.
    
    Args:
        sheet (Worksheet): The worksheet to read from
        base_row (int): Base row for this canopy (model row)
        
    Returns:
        Dict: Wall cladding data
    """
    try:
        cladding_indicator_row = base_row + 5  # Row for wall cladding indicator (row 19 for first canopy)
        cladding_data_row = base_row + 6  # Row for wall cladding data (row 20 for first canopy)
        
        # Check for "2M² (HFL)" indicator in column C
        cladding_indicator = sheet[f"C{cladding_indicator_row}"].value or ""
        has_cladding_indicator = "2M²" in str(cladding_indicator).upper() or "HFL" in str(cladding_indicator).upper()
        
        # Read wall cladding data from columns P, Q, S
        width = sheet[f"P{cladding_data_row}"].value or None
        height = sheet[f"Q{cladding_data_row}"].value or None
        position_str = sheet[f"S{cladding_data_row}"].value or None
        
        # Check if any wall cladding data exists (either indicator or actual data)
        if has_cladding_indicator or width or height or position_str:
            # Convert position string to list (handle "and" separator)
            if position_str and str(position_str).strip():
                # Split by "and" and clean up each position
                position_list = [pos.strip() for pos in str(position_str).split(" and ")]
                # Filter out empty strings
                position_list = [pos for pos in position_list if pos]
            else:
                position_list = []
            
            return {
                'type': 'Custom',  # Indicate this is custom wall cladding
                'width': int(width) if width and str(width).replace('.', '').isdigit() else None,
                'height': int(height) if height and str(height).replace('.', '').isdigit() else None,
                'position': position_list
            }
        else:
            # No wall cladding data found
            return {
                'type': 'None',
                'width': None,
                'height': None,
                'position': None
            }
    except Exception as e:
        print(f"Warning: Could not read wall cladding data from row {base_row}: {str(e)}")
        return {
            'type': 'None',
            'width': None,
            'height': None,
            'position': None
        }

def extract_tank_quantity(tank_value) -> int:
    """
    Extract tank quantity number from tank value strings like "1 TANK", "2 TANK", etc.
    
    Args:
        tank_value: Tank value from Excel cell (could be string, number, or None)
        
    Returns:
        int: Tank quantity number, or 0 if not found/invalid
    """
    if not tank_value:
        return 0
    
    # Convert to string and clean up
    tank_str = str(tank_value).strip().upper()
    
    # Handle empty or dash values
    if tank_str == "" or tank_str == "-":
        return 0
    
    # Extract number from strings like "1 TANK", "2 TANK", "3 TANKS", etc.
    try:
        # Split by space and look for the first part that's a number
        parts = tank_str.split()
        for part in parts:
            # Try to convert each part to int
            try:
                return int(part)
            except ValueError:
                continue
        
        # If no number found in parts, try to extract digits from the whole string
        import re
        numbers = re.findall(r'\d+', tank_str)
        if numbers:
            return int(numbers[0])  # Return the first number found
        
        return 0
    except (ValueError, AttributeError):
        return 0

def transform_recoair_model(model_str: str) -> str:
    """
    Transform RecoAir model names according to business rules.
    
    Examples:
    - RA1.0 STANDARD -> RAH1.0
    - RAH0.5 STANDARD -> RAH0.5 (already has H)
    - RAH0.5 VOID -> RAH0.5V
    - RA2.5 STANDARD (Prem Controls) -> RAH2.5 (no P suffix)
    - RA1.5 VOID (+10%) -> RAH1.5V
    
    Args:
        model_str (str): Original model string from Excel
        
    Returns:
        str: Transformed model string
    """
    if not model_str:
        return ""
    
    model = str(model_str).strip().upper()
    
    # Extract the base model number (e.g., "0.5", "1.0", "2.5")
    import re
    
    # Look for pattern like RA(H)?X.X where X.X is the model number
    base_match = re.search(r'RA(H)?(\d+\.\d+)', model)
    if not base_match:
        return model  # Return original if pattern doesn't match
    
    has_h = base_match.group(1) is not None  # Check if H is already present
    model_number = base_match.group(2)  # e.g., "1.0", "2.5"
    
    # Start with RAH + model number
    result = f"RAH{model_number}"
    
    # Add suffixes based on content (only VOID gets a suffix)
    if 'VOID' in model:
        result += 'V'
    # For STANDARD and PREM CONTROLS, no suffix needed
    
    return result

def get_recoair_specifications(model: str) -> dict:
    """
    Get technical specifications for RecoAir models based on the model name.
    
    Args:
        model (str): RecoAir model name (e.g., "RAH1.0", "RAH0.5V")
        
    Returns:
        dict: Technical specifications including p_drop, motor, weight
    """
    # RecoAir specifications table
    specifications = {
        # Standard models
        'RAH0.5': {'p_drop': 1050, 'motor': 2.2, 'weight': 436},
        'RAH0.8': {'p_drop': 1050, 'motor': 2.2, 'weight': 470},
        'RAH1.0': {'p_drop': 1050, 'motor': 4.7, 'weight': 572},
        'RAH1.5': {'p_drop': 1050, 'motor': 4.7, 'weight': 820},
        'RAH2.0': {'p_drop': 1050, 'motor': 5.25, 'weight': 974},
        'RAH2.5': {'p_drop': 1050, 'motor': 5.25, 'weight': 1170},
        'RAH3.0': {'p_drop': 1050, 'motor': 5.25, 'weight': 1210},
        'RAH3.5': {'p_drop': 1050, 'motor': 5.25, 'weight': 1395},
        'RAH4.0': {'p_drop': 1050, 'motor': 5.25, 'weight': 1500},
        
        # VOID models (V suffix)
        'RAH0.5V': {'p_drop': 1050, 'motor': 2.2, 'weight': 385},
        'RAH0.8V': {'p_drop': 1050, 'motor': 2.2, 'weight': 415},
        'RAH1.0V': {'p_drop': 1050, 'motor': 4.7, 'weight': 542},
        'RAH1.5V': {'p_drop': 1050, 'motor': 4.7, 'weight': 765},
        'RAH2.0V': {'p_drop': 1050, 'motor': 5.25, 'weight': 884},
        'RAH2.5V': {'p_drop': 1050, 'motor': 5.25, 'weight': 1093},
        'RAH3.0V': {'p_drop': 1050, 'motor': 5.25, 'weight': 1210},
        'RAH3.5V': {'p_drop': 1050, 'motor': 5.25, 'weight': 1395},
        'RAH4.0V': {'p_drop': 1050, 'motor': 5.25, 'weight': 1500},
    }
    
    # Return specifications for the model, or default values if not found
    return specifications.get(model, {'p_drop': 0, 'motor': 0, 'weight': 0})

def extract_recoair_volume(volume_str) -> float:
    """
    Extract volume number from RecoAir volume strings like "VERTICAL 1.2M3/S".
    
    Args:
        volume_str: Volume string from Excel cell (could be string or None)
        
    Returns:
        float: Volume number, or 0.0 if not found/invalid
    """
    if not volume_str:
        return 0.0
    
    # Convert to string and clean up
    volume_string = str(volume_str).strip()
    
    # Handle empty or dash values
    if volume_string == "" or volume_string == "-":
        return 0.0
    
    try:
        # Use regex to find decimal numbers in the string
        import re
        # Look for patterns like "1.2", "2.5", "10.0", etc.
        # But avoid matching single digits that are part of "M3/S"
        numbers = re.findall(r'\d+\.\d+|\d+(?!\d*[/])', volume_string)
        if numbers:
            # Return the first number found as float
            return float(numbers[0])
        
        return 0.0
    except (ValueError, AttributeError):
        return 0.0

def read_recoair_data_from_sheet(sheet: Worksheet) -> Dict:
    """
    Read RecoAir unit data from a RECOAIR sheet.
    
    Args:
        sheet (Worksheet): The RECOAIR worksheet to read from
        
    Returns:
        List[Dict]: List of RecoAir units found in the sheet
    """
    recoair_units = []
    
    try:
        # Get item reference from C12 (e.g., "1.01", "2.01")
        item_reference = sheet['C12'].value or ""
        
        # Get delivery and installation price from P36
        delivery_installation_price = sheet['P36'].value or 0
        
        # Get N29 value (default addition to RecoAir unit price)
        n29_value = sheet['N29'].value or 0
        
        # Get flat pack data from D40 and N40
        flat_pack_description = sheet['D40'].value or ""
        flat_pack_price = sheet['N40'].value or 0
        
        # Check rows 14 to 28 for RecoAir unit selections
        for row in range(14, 29):  # 14 to 28 inclusive
            # Check if there's a value of 1 or more in column E (selection indicator)
            selection_value = sheet[f'E{row}'].value
            
            if selection_value and str(selection_value).strip() != "":
                try:
                    # Try to convert to number
                    selection_num = float(str(selection_value).strip())
                    if selection_num >= 1:
                        # This row has a selected RecoAir unit
                        # Collect data from this row
                        model = sheet[f'C{row}'].value or ""
                        extract_volume_str = sheet[f'D{row}'].value or ""
                        width = sheet[f'F{row}'].value or 0
                        length = sheet[f'G{row}'].value or 0
                        height = sheet[f'H{row}'].value or 0
                        location_raw = sheet[f'I{row}'].value or "INTERNAL"  # Default to INTERNAL
                        unit_price = sheet[f'N{row}'].value or 0
                        
                        # Clean up location value - handle placeholder text
                        if location_raw:
                            location_str = str(location_raw).strip().upper()
                            # If location is a placeholder like "SELECT..." or empty, default to INTERNAL
                            if location_str in ["SELECT...", "SELECT", "", "-"] or "SELECT" in location_str:
                                location = "INTERNAL"
                            else:
                                location = location_str
                        else:
                            location = "INTERNAL"
                        
                        # Extract volume number from extract volume string
                        extract_volume = extract_recoair_volume(extract_volume_str)
                        
                        # Transform the model name according to business rules
                        original_model = str(model).strip() if model else ""
                        transformed_model = transform_recoair_model(original_model)
                        
                        # Get technical specifications for this model
                        specs = get_recoair_specifications(transformed_model)
                        
                        # Calculate final unit price (base price + N29 value)
                        base_unit_price = unit_price if isinstance(unit_price, (int, float)) else 0
                        n29_addition = n29_value if isinstance(n29_value, (int, float)) else 0
                        final_unit_price = base_unit_price + n29_addition
                        
                        # Create RecoAir unit data
                        recoair_unit = {
                            'item_reference': str(item_reference).strip() if item_reference else "",
                            'model': transformed_model,
                            'model_original': original_model,  # Keep original for reference
                            'extract_volume': extract_volume,
                            'extract_volume_raw': str(extract_volume_str).strip() if extract_volume_str else "",
                            'width': width if isinstance(width, (int, float)) else 0,
                            'length': length if isinstance(length, (int, float)) else 0,
                            'height': height if isinstance(height, (int, float)) else 0,
                            'location': location,
                            'unit_price': final_unit_price,  # Use final price (base + N29)
                            'base_unit_price': base_unit_price,  # Keep original base price for reference
                            'n29_addition': n29_addition,  # Keep N29 addition for reference
                            'quantity': selection_num,
                            'row': row,  # Keep track of which row this came from
                            
                            # Technical specifications
                            'p_drop': specs['p_drop'],  # Pressure drop (Pa)
                            'motor': specs['motor'],    # Motor power (kW/PH)
                            'weight': specs['weight']   # Weight (kg)
                        }
                        
                        recoair_units.append(recoair_unit)
                        
                except (ValueError, TypeError):
                    # Skip if selection value can't be converted to number
                    continue
        
        # Add delivery price to each unit (split equally if multiple units)
        if recoair_units and delivery_installation_price > 0:
            delivery_per_unit = delivery_installation_price / len(recoair_units)
            for unit in recoair_units:
                unit['delivery_installation_price'] = delivery_per_unit
        else:
            for unit in recoair_units:
                unit['delivery_installation_price'] = 0
        
        # Create result dictionary with units and flat pack data
        result = {
            'units': recoair_units,
            'flat_pack': {
                'item_reference': str(item_reference).strip() if item_reference else "",  # Add item reference to flat pack
                'description': flat_pack_description,
                'price': flat_pack_price if isinstance(flat_pack_price, (int, float)) else 0,
                'has_flat_pack': bool(flat_pack_description and str(flat_pack_description).strip())
            }
        }
        
        return result
        
    except Exception as e:
        print(f"Warning: Could not read RecoAir data from sheet: {str(e)}")
        return {
            'units': [],
            'flat_pack': {
                'item_reference': '',  # Add item reference to flat pack error case
                'description': '',
                'price': 0,
                'has_flat_pack': False
            }
        }

def write_project_metadata(sheet: Worksheet, project_data: Dict):
    """
    Write project metadata to the specified cells in the sheet.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        project_data (Dict): Project metadata
    """
    for field, cell in CELL_MAPPINGS.items():
        value = project_data.get(field)
        
        try:
            # Special handling for revision - use the value from project_data
            if field == "revision":
                sheet[cell] = value or "A"  # Use provided revision or default to "A"
            elif value:
                # Special handling for estimator/sales manager initials (only for sheet display)
                if field == "estimator":
                    # Generate combined initials (Sales Contact + Estimator)
                    from utils.word import get_combined_initials
                    from config.business_data import SALES_CONTACTS
                    
                    # Get sales contact info based on estimator
                    sales_contact_name = ""
                    for contact_name, phone in SALES_CONTACTS.items():
                        if value and any(name.lower() in value.lower() for name in contact_name.split()):
                            sales_contact_name = contact_name
                            break
                    
                    # If no match found, use first sales contact
                    if not sales_contact_name:
                        sales_contact_name = list(SALES_CONTACTS.keys())[0]
                    
                    # Generate combined initials
                    value = get_combined_initials(sales_contact_name, value)
                # Title case for other fields except date
                elif field != "date":
                    value = str(value).title()
                # Date handling
                elif field == "date" and not value:
                    value = datetime.now().strftime("%d/%m/%Y")
                
                sheet[cell] = value
        except Exception as e:
            # Handle merged cells or other write errors
            print(f"Warning: Could not write {field} to cell {cell}: {str(e)}")
            try:
                # Try to unmerge the cell and write
                if hasattr(sheet, 'merged_cells'):
                    for merged_range in list(sheet.merged_cells.ranges):
                        if cell in merged_range:
                            sheet.unmerge_cells(str(merged_range))
                            break
                # Try writing again after unmerging
                if field == "revision":
                    sheet[cell] = value or "A"  # Use provided revision or default to "A"
                elif value:
                    if field == "estimator":
                        # Generate combined initials (Sales Contact + Estimator)
                        from utils.word import get_combined_initials
                        from config.business_data import SALES_CONTACTS
                        
                        # Get sales contact info based on estimator
                        sales_contact_name = ""
                        for contact_name, phone in SALES_CONTACTS.items():
                            if value and any(name.lower() in value.lower() for name in contact_name.split()):
                                sales_contact_name = contact_name
                                break
                        
                        # If no match found, use first sales contact
                        if not sales_contact_name:
                            sales_contact_name = list(SALES_CONTACTS.keys())[0]
                        
                        # Generate combined initials
                        value = get_combined_initials(sales_contact_name, value)
                    elif field != "date":
                        value = str(value).title()
                    elif field == "date" and not value:
                        value = datetime.now().strftime("%d/%m/%Y")
                    sheet[cell] = value
            except Exception as e2:
                print(f"Warning: Still could not write {field} to cell {cell} after unmerging: {str(e2)}")
                continue

def write_canopy_data(sheet: Worksheet, canopy: Dict, row_index: int):
    """
    Write canopy specifications to the sheet at the specified row.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        canopy (Dict): Canopy specification data
        row_index (int): Starting row for this canopy's data (this is the model/config row)
    """
    try:
        # Reference number starts 2 rows before configuration/model
        ref_row = row_index - 2  # If row_index is 14, ref_row will be 12
        ref_number = canopy.get("reference_number", "")
        if ref_number:
            try:
                sheet[f"B{ref_row}"] = ref_number.upper()
            except Exception as e:
                print(f"Warning: Could not write reference number to B{ref_row}: {str(e)}")
        
        # Configuration and Model on same row
        configuration = canopy.get("configuration", "")
        if configuration:
            try:
                sheet[f"C{row_index}"] = configuration.upper()
            except Exception as e:
                print(f"Warning: Could not write configuration to C{row_index}: {str(e)}")
        
        # Model in D14, D31, D48, etc.
        model = canopy.get("model", "")
        if model:
            try:
                sheet[f"D{row_index}"] = model.upper()
                
                # For CMWF/CMWI canopies, initialize C27 (base_row + 13) to 0
                if model.upper() in ['CMWF', 'CMWI']:
                    initial_value_row = row_index + 13  # C27, C44, C61, etc.
                    try:
                        sheet[f"C{initial_value_row}"] = 0
                    except Exception as e:
                        print(f"Warning: Could not initialize C{initial_value_row} to 0 for CMWF/CMWI canopy: {str(e)}")
                        
            except Exception as e:
                print(f"Warning: Could not write model to D{row_index}: {str(e)}")
        
        # Options (only fire suppression at canopy level now)
        options_row = row_index + 4
        options = canopy.get("options", {})
        if options.get("fire_suppression"):
            try:
                sheet[f"B{options_row}"] = "FIRE SUPPRESSION SYSTEM"
            except Exception as e:
                print(f"Warning: Could not write fire suppression to B{options_row}: {str(e)}")
        
        # Wall cladding data (if present)
        wall_cladding = canopy.get("wall_cladding", {})
        if wall_cladding and wall_cladding.get('type') not in ['None', None, '']:
            cladding_indicator_row = row_index + 5  # Row for wall cladding indicator (row 19 for first canopy)
            cladding_data_row = row_index + 6  # Row for wall cladding data (row 20 for first canopy)
            try:
                # Write "2M² (HFL)" indicator in column C (C19, C36, C53, etc.)
                sheet[f"C{cladding_indicator_row}"] = "2M² (HFL)"
                
                # Write wall cladding width in column P
                if wall_cladding.get('width'):
                    sheet[f"P{cladding_data_row}"] = wall_cladding['width']
                
                # Write wall cladding height in column Q  
                if wall_cladding.get('height'):
                    sheet[f"Q{cladding_data_row}"] = wall_cladding['height']
                
                # Write wall cladding position in column S
                position = wall_cladding.get('position', [])
                if isinstance(position, list):
                    position_str = " and ".join(position) if position else ""
                else:
                    position_str = str(position) if position else ""
                
                if position_str:
                    sheet[f"S{cladding_data_row}"] = position_str
                    
            except Exception as e:
                print(f"Warning: Could not write wall cladding data to row {cladding_indicator_row}: {str(e)}")
    except Exception as e:
        raise Exception(f"Failed to write canopy data: {str(e)}")

def write_area_options(sheet: Worksheet, area: Dict):
    """
    Write area-level options (UV-C, SDU, RecoAir) to the sheet.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        area (Dict): Area data containing options
    """
    try:
        # Write area options in a dedicated section (e.g., starting at row 6)
        options_start_row = 6
        area_options = area.get("options", {})
        
        if area_options.get("uvc"):
            try:
                sheet[f"B{options_start_row}"] = "UV-C SYSTEM"
            except Exception as e:
                print(f"Warning: Could not write UV-C option to B{options_start_row}: {str(e)}")
        if area_options.get("sdu"):
            try:
                sheet[f"B{options_start_row + 1}"] = "SDU"
            except Exception as e:
                print(f"Warning: Could not write SDU option to B{options_start_row + 1}: {str(e)}")
        if area_options.get("recoair"):
            try:
                sheet[f"B{options_start_row + 2}"] = "RECOAIR"
            except Exception as e:
                print(f"Warning: Could not write RecoAir option to B{options_start_row + 2}: {str(e)}")
    except Exception as e:
        print(f"Warning: Could not write area options: {str(e)}")
        pass

def write_fire_suppression_canopy_data(sheet: Worksheet, canopy: Dict, row_index: int):
    """
    Write canopy reference number to the fire suppression sheet at the specified row.
    
    Args:
        sheet (Worksheet): The fire suppression worksheet to write to
        canopy (Dict): Canopy specification data
        row_index (int): Starting row for this canopy's data (this is the model/config row)
    """
    try:
        # Reference number starts 2 rows before configuration/model (same pattern as canopy sheets)
        ref_row = row_index - 2  # If row_index is 14, ref_row will be 12
        sheet[f"B{ref_row}"] = canopy["reference_number"].upper()
    except Exception as e:
        raise Exception(f"Failed to write fire suppression canopy data: {str(e)}")

def write_to_sheet(sheet: Worksheet, project_data: Dict, level_name: str, area_name: str, canopies: List[Dict]):
    """
    Write all data for an area to a sheet.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        project_data (Dict): Project metadata
        level_name (str): Name of the level
        area_name (str): Name of the area
        canopies (List[Dict]): List of canopy specifications
    """
    try:
        # Write level and area information in B1
        sheet["B1"] = f"{level_name} - {area_name}"
        
        # Write project metadata
        write_project_metadata(sheet, project_data)
        
        # Write each canopy with proper spacing (17 rows)
        for idx, canopy in enumerate(canopies):
            row_start = CANOPY_START_ROW + (idx * CANOPY_ROW_SPACING)  # Starts at 14, then 31, 48, etc.
            write_canopy_data(sheet, canopy, row_start)
            
            # If this canopy has fire suppression and fire suppression sheet exists, write to it
            if canopy.get("options", {}).get("fire_suppression") and fs_sheet:
                fs_row_start = CANOPY_START_ROW + (fs_canopy_idx * CANOPY_ROW_SPACING)
                write_fire_suppression_canopy_data(fs_sheet, canopy, fs_row_start)
                fs_canopy_idx += 1  # Only increment for canopies with fire suppression
        
        # Add dropdowns
        add_dropdowns_to_sheet(wb, current_canopy_sheet)
        if fs_sheet:
            # Add fire suppression specific dropdowns
            add_fire_suppression_dropdowns(fs_sheet)
        
        sheet_count += 1
    except Exception as e:
        raise Exception(f"Failed to write sheet data: {str(e)}")

def add_dropdowns_to_sheet(wb: Workbook, sheet: Worksheet, start_row: int = 12):
    """
    Add data validation (dropdowns) to specific cells in the sheet.
    
    Args:
        wb (Workbook): The workbook containing the sheet
        sheet (Worksheet): The worksheet to add dropdowns to
        start_row (int): Starting row for dropdowns
    """
    try:
        # Define dropdown options (keeping them shorter to avoid Excel corruption)
        lighting_options = [
            'LED STRIP L6 Inc DALI',
            'LED STRIP L12 inc DALI', 
            'LED STRIP L18 Inc DALI',
            'Small LED Spots inc DALI',
            'LARGE LED Spots inc DALI'
        ]
        
        special_works_options = [
            'ROUND CORNERS',
                'CUT OUT',
                'CASTELLE LOCKING ',
                'HEADER DUCT S/S',
                'HEADER DUCT',
                'PAINT FINSH',
                'UV ON DEMAND',
                'E/over for emergency strip light',
                'E/over for small emer. spot light',
                'E/over for large emer. spot light',
                'COLD MIST ON DEMAND',
                'CMW  PIPEWORK HWS/CWS',
                'CANOPY GROUND SUPPORT',
                ' 2nd EXTRACT PLENUM',
                'SUPPLY AIR PLENUM',
                'CAPTUREJET PLENUM',
                'COALESCER'
        ]
        
        cladding_options = [
            "Standard Stainless Steel",
            "Brushed Stainless Steel",
            "Painted Steel",
            "Galvanized Steel", 
            "Aluminum Composite",
            "No Cladding"
        ]
        
        # Wall cladding options for C19
        wall_cladding_options = [
            "",  # Empty option
            "2M² (HFL)"
        ]
        
        # Wall cladding position options for column S (comprehensive combinations)
        wall_cladding_position_options = [
            "",  # Empty option
            # Single positions
            "Rear",
            "Left", 
            "Right",
            "Front",
            # Two-position combinations
            "Rear and Left",
            "Rear and Right", 
            "Rear and Front",
            "Left and Right",
            "Left and Front",
            "Right and Front",
            # Three-position combinations
            "Rear and Left and Right",
            "Rear and Left and Front",
            "Rear and Right and Front",
            "Left and Right and Front",
            # All sides
            "All Sides"
        ]
        
        # CMWF/CMWI panel options for wash canopies
        cmw_panel_type_options = [
            "",  # Empty option
            "CP1S",
            "CP2S", 
            "CP3S",
            "CP4S"
        ]
        
        cmw_panel_size_options = [
            "",  # Empty option
            "1000-S",
            "1500-S",
            "2000-S",
            "2500-S",
            "3000-S",
            "1000-D",
            "1500-D",
            "2000-D",
            "2500-D",
            "3000-D"
        ]
        
        # Create data validations with proper escaping
        def create_validation(options):
            # Escape quotes and limit formula length
            formula = ",".join(options)
            if len(formula) > 255:  # Excel formula limit
                # Use only first few options if too long
                truncated_options = []
                current_length = 0
                for opt in options:
                    if current_length + len(opt) + 1 > 250:  # Leave some buffer
                        break
                    truncated_options.append(opt)
                    current_length += len(opt) + 1
                formula = ",".join(truncated_options)
            return DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
        
        lighting_dv = create_validation(lighting_options)
        special_works_dv = create_validation(special_works_options)
        cladding_dv = create_validation(cladding_options)
        wall_cladding_dv = create_validation(wall_cladding_options)
        wall_cladding_position_dv = create_validation(wall_cladding_position_options)
        cmw_panel_type_dv = create_validation(cmw_panel_type_options)
        cmw_panel_size_dv = create_validation(cmw_panel_size_options)
        
        # Add validations to sheet
        sheet.add_data_validation(lighting_dv)
        sheet.add_data_validation(special_works_dv)
        sheet.add_data_validation(cladding_dv)
        sheet.add_data_validation(wall_cladding_dv)
        sheet.add_data_validation(wall_cladding_position_dv)
        sheet.add_data_validation(cmw_panel_type_dv)
        sheet.add_data_validation(cmw_panel_size_dv)
        
        # Write wall cladding headers in row 19 (first canopy's cladding row - 1)
        try:
            sheet['P19'] = 'Width (mm)'
            sheet['Q19'] = 'Height (mm)'
            sheet['S19'] = 'Position'
        except Exception as e:
            print(f"Warning: Could not write wall cladding headers: {str(e)}")
        
        # Add wall cladding input fields for all 10 canopy rows (even if no canopy exists)
        for canopy_index in range(10):  # Support up to 10 canopies per sheet
            base_row = CANOPY_START_ROW + (canopy_index * CANOPY_ROW_SPACING)  # 14, 31, 48, 65, 82, 99, 116, 133, 150, 167
            cladding_row = base_row + 6  # 20, 37, 54, 71, 88, 105, 122, 139, 156, 173
            
            try:
                # Add placeholder values to ensure cells exist for dropdowns (only if cells are currently None)
                # Use direct cell assignment to ensure values are set
                width_cell = sheet[f'P{cladding_row}']
                height_cell = sheet[f'Q{cladding_row}']
                position_cell = sheet[f'S{cladding_row}']
                
                # Only set placeholder if cell is currently None (don't overwrite existing data)
                if width_cell.value is None:
                    width_cell.value = ""  # Width placeholder
                if height_cell.value is None:
                    height_cell.value = ""  # Height placeholder
                if position_cell.value is None:
                    position_cell.value = ""  # Position placeholder
                
                # Values are set to empty strings only if no data exists
            except Exception as e:
                print(f"Warning: Could not add wall cladding placeholders for row {cladding_row}: {str(e)}")
        
        # Apply dropdowns to multiple canopy sections (every 17 rows)
        for canopy_index in range(10):  # Support up to 10 canopies per sheet
            base_row = CANOPY_START_ROW + (canopy_index * CANOPY_ROW_SPACING)  # 14, 31, 48, 65, 82, 99, 116, 133, 150, 167
            
            try:
                # Lighting options - typically around row 15 (C15)
                lighting_row = base_row + 1
                lighting_dv.add(f"C{lighting_row}")
                
                # Special works options - C16, C17, C18 (base_row + 2, 3, 4)
                special_works_dv.add(f"C{base_row + 2}")  # C16
                special_works_dv.add(f"C{base_row + 3}")  # C17
                special_works_dv.add(f"C{base_row + 4}")  # C18
                
                # Wall cladding options - C19 (base_row + 5)
                wall_cladding_row = base_row + 5
                wall_cladding_dv.add(f"C{wall_cladding_row}")  # C19
                
                # Wall cladding position dropdown - S20 (base_row + 6)
                wall_cladding_position_row = base_row + 6
                wall_cladding_position_dv.add(f"S{wall_cladding_position_row}")  # S20, S37, S54, S71, S88, S105, S122, S139, S156, S173
                
                # Cladding options - typically around row 16 (D16) for cladding type
                cladding_row = base_row + 2
                cladding_dv.add(f"D{cladding_row}")
            except Exception as e:
                print(f"Warning: Could not add dropdown to canopy {canopy_index + 1}: {str(e)}")
                continue
        
        # Add some additional dropdowns for common fields
        # Configuration options for column C (model row)
        config_options = ["Wall", "Island", "Single", "Double", "Corner"]
        config_dv = create_validation(config_options)
        sheet.add_data_validation(config_dv)
        
        # Model options for column D (model row)
        model_dv = create_validation(VALID_CANOPY_MODELS)
        sheet.add_data_validation(model_dv)
        
        for canopy_index in range(10):
            try:
                base_row = CANOPY_START_ROW + (canopy_index * CANOPY_ROW_SPACING)
                config_dv.add(f"C{base_row}")  # Configuration in column C of the model row
                model_dv.add(f"D{base_row}")   # Model in column D of the model row (D14, D31, D48, etc.)
                
                # Add CMWF/CMWI panel options dropdowns for all canopies
                # These will be available for all canopies but are specifically for CMWF/CMWI models
                cmw_panel_type_row = base_row + 11  # C25, C42, C59, etc. (base_row + 11)
                cmw_panel_size_row = base_row + 12  # C26, C43, C60, etc. (base_row + 12)
                
                cmw_panel_type_dv.add(f"C{cmw_panel_type_row}")  # Panel type dropdown
                cmw_panel_size_dv.add(f"C{cmw_panel_size_row}")  # Panel size dropdown
                
            except Exception as e:
                print(f"Warning: Could not add config/model dropdown to canopy {canopy_index + 1}: {str(e)}")
                continue
        
    except Exception as e:
        # Silently fail for dropdown addition to avoid breaking the main process
        print(f"Warning: Could not add dropdowns to sheet {sheet.title}: {str(e)}")
        pass

def set_sheet_tab_color(sheet: Worksheet, area_index: int):
    """
    Set the tab color for a worksheet based on the area index.
    
    Args:
        sheet (Worksheet): The worksheet to color
        area_index (int): The area index (0-based)
    """
    # Get color from list, cycling through colors if area index exceeds available colors
    color = TAB_COLORS[area_index % len(TAB_COLORS)]
    sheet.sheet_properties.tabColor = color

def write_ebox_metadata(sheet: Worksheet, project_data: Dict):
    """
    Write project metadata to EBOX sheet with specific cell mappings.
    
    Args:
        sheet (Worksheet): The EBOX worksheet to write to
        project_data (Dict): Project metadata
    """
    try:
        # EBOX-specific cell mappings
        ebox_cell_mappings = {
            "project_number": "D3",    # Job No
            "customer": "D5",          # Customer
            "estimator": "D7",         # Sales Manager / Estimator Initials
            "project_name": "H3",      # Project Name
            "project_location": "H5",          # Project Location (was "location")
            "date": "H7",             # Date
            "revision": "O7",         # Revision
        }
        
        # Write EBOX-specific data
        try:
            sheet["C12"] = "UV-C"  # Model name
            sheet["D38"] = 1       # Quantity
            # Delivery location to E38
            delivery_location = project_data.get('delivery_location', '')
            if delivery_location:
                sheet["E38"] = delivery_location
        except Exception as e:
            print(f"Warning: Could not write EBOX-specific data: {str(e)}")
        
        for field, cell in ebox_cell_mappings.items():
            value = project_data.get(field)
            
            try:
                # Special handling for revision - use the value from project_data
                if field == "revision":
                    sheet[cell] = value or "A"  # Use provided revision or default to "A"
                elif value:
                    # Special handling for estimator/sales manager initials
                    if field == "estimator":
                        # Generate combined initials (Sales Contact + Estimator)
                        from utils.word import get_combined_initials
                        from config.business_data import SALES_CONTACTS
                        
                        # Get sales contact info based on estimator
                        sales_contact_name = ""
                        for contact_name, phone in SALES_CONTACTS.items():
                            if value and any(name.lower() in value.lower() for name in contact_name.split()):
                                sales_contact_name = contact_name
                                break
                        
                        # If no match found, use first sales contact
                        if not sales_contact_name:
                            sales_contact_name = list(SALES_CONTACTS.keys())[0]
                        
                        # Generate combined initials
                        value = get_combined_initials(sales_contact_name, value)
                    # Title case for other fields except date
                    elif field != "date":
                        value = str(value).title()
                    # Date handling
                    elif field == "date" and not value:
                        value = datetime.now().strftime("%d/%m/%Y")
                    
                    sheet[cell] = value
            except Exception as e:
                # Handle merged cells or other write errors
                print(f"Warning: Could not write {field} to EBOX cell {cell}: {str(e)}")
                try:
                    # Try to unmerge the cell and write
                    if hasattr(sheet, 'merged_cells'):
                        for merged_range in list(sheet.merged_cells.ranges):
                            if cell in merged_range:
                                sheet.unmerge_cells(str(merged_range))
                                break
                    # Try writing again after unmerging
                    if field == "revision":
                        sheet[cell] = value or "A"  # Use provided revision or default to "A"
                    elif value:
                        if field == "estimator":
                            # Generate combined initials (Sales Contact + Estimator)
                            from utils.word import get_combined_initials
                            from config.business_data import SALES_CONTACTS
                            
                            # Get sales contact info based on estimator
                            sales_contact_name = ""
                            for contact_name, phone in SALES_CONTACTS.items():
                                if value and any(name.lower() in value.lower() for name in contact_name.split()):
                                    sales_contact_name = contact_name
                                    break
                            
                            # If no match found, use first sales contact
                            if not sales_contact_name:
                                sales_contact_name = list(SALES_CONTACTS.keys())[0]
                            
                            # Generate combined initials
                            value = get_combined_initials(sales_contact_name, value)
                        elif field != "date":
                            value = str(value).title()
                        elif field == "date" and not value:
                            value = datetime.now().strftime("%d/%m/%Y")
                        sheet[cell] = value
                except Exception as e2:
                    print(f"Warning: Still could not write {field} to EBOX cell {cell} after unmerging: {str(e2)}")
                    continue
    except Exception as e:
        print(f"Warning: Could not write EBOX metadata: {str(e)}")
        pass

def write_recoair_metadata(sheet: Worksheet, project_data: Dict, item_number: str = "1.01"):
    """
    Write project metadata to RECOAIR sheet with specific cell mappings.
    
    Args:
        sheet (Worksheet): The RECOAIR worksheet to write to
        project_data (Dict): Project data dictionary
        item_number (str): Item number for this RecoAir sheet (e.g., "1.01", "2.01")
    """
    try:
        # RECOAIR-specific cell mappings (same as EBOX - D/H columns)
        recoair_cell_mappings = {
            "project_number": "D3",  # Job No
            "customer": "D5",        # Customer
            "estimator": "D7",       # Sales Manager / Estimator Initials
            "project_name": "H3",    # Project Name
            "project_location": "H5",        # Project Location (was "location")
            "date": "H7",           # Date
            "revision": "O7",       # Revision
        }
        
        # Write RECOAIR-specific data
        try:
            sheet['C12'] = item_number  # Item number (1.01, 2.01, etc.)
            sheet['D37'] = 1  # Quantity
            sheet['E37'] = project_data.get('delivery_location', '')  # Delivery location
            # N9 cell ready for RecoAir price (to be implemented)
        except Exception as e:
            print(f"Warning: Could not write RECOAIR-specific data: {str(e)}")
        
        for field, cell in recoair_cell_mappings.items():
            try:
                value = project_data.get(field, "")
                
                # Handle special cases
                if field == "estimator":
                    # Generate combined initials (Sales Contact + Estimator) for RECOAIR sheets
                    from utils.word import get_combined_initials
                    from config.business_data import SALES_CONTACTS
                    
                    estimator_name = project_data.get("estimator", "")
                    
                    # Get sales contact info based on estimator
                    sales_contact_name = ""
                    for contact_name, phone in SALES_CONTACTS.items():
                        if estimator_name and any(name.lower() in estimator_name.lower() for name in contact_name.split()):
                            sales_contact_name = contact_name
                            break
                    
                    # If no match found, use first sales contact
                    if not sales_contact_name:
                        sales_contact_name = list(SALES_CONTACTS.keys())[0]
                    
                    # Generate combined initials
                    value = get_combined_initials(sales_contact_name, estimator_name)
                elif field == "revision":
                    value = project_data.get("revision", "A")  # Use provided revision or default to "A"
                elif field == "date":
                    # Keep date as is from project data
                    value = project_data.get("date", "")
                
                # Write the value to the cell
                sheet[cell] = value
                
            except Exception as e:
                print(f"Warning: Could not write {field} to RECOAIR cell {cell}: {str(e)}")
                
                # Try to handle merged cells
                try:
                    # Check if the cell is part of a merged range
                    for merged_range in sheet.merged_cells.ranges:
                        if cell in merged_range:
                            # Unmerge the range temporarily
                            sheet.unmerge_cells(str(merged_range))
                            # Write the value
                            sheet[cell] = value
                            # Re-merge the range
                            sheet.merge_cells(str(merged_range))
                            break
                    else:
                        # Cell is not merged, try writing again
                        sheet[cell] = value
                        
                except Exception as e2:
                    print(f"Warning: Still could not write {field} to RECOAIR cell {cell} after unmerging: {str(e2)}")
        
    except Exception as e:
        print(f"Warning: Could not write RECOAIR metadata: {str(e)}")

def write_sdu_metadata(sheet: Worksheet, project_data: Dict):
    """
    Write project metadata to SDU sheet with specific cell mappings.
    
    Args:
        sheet (Worksheet): The SDU worksheet to write to
        project_data (Dict): Project data dictionary
    """
    try:
        
        # Write SDU-specific data
        try:
            # Write model name to C12
            sheet['C12'] = "SDU"  # Model name
            
            # Write quantity (1) to C97
            sheet['C97'] = 1
            
            # Write delivery location to D97
            delivery_location = project_data.get('delivery_location', '')
            if delivery_location:
                sheet['D97'] = delivery_location
        except Exception as e:
            print(f"Warning: Could not write SDU-specific data: {str(e)}")
        
        # Write project metadata to SDU-specific cells with merged cell handling
        def write_to_cell_safe(sheet, cell_ref, value):
            """Safely write to a cell, handling merged cells by unmerging first."""
            try:
                # Check if the cell is part of a merged range
                for merged_range in list(sheet.merged_cells.ranges):
                    if cell_ref in merged_range:
                        # Unmerge the range temporarily
                        sheet.unmerge_cells(str(merged_range))
                        # Write the value
                        sheet[cell_ref] = value
                        # Re-merge the range
                        sheet.merge_cells(str(merged_range))
                        return
                # If not merged, write directly
                sheet[cell_ref] = value
            except Exception as e:
                print(f"Warning: Could not write to {cell_ref}: {str(e)}")
        
        try:
            # Job No at C4
            write_to_cell_safe(sheet, 'C4', project_data.get('project_number', ''))
            
            # Customer at C6
            write_to_cell_safe(sheet, 'C6', project_data.get('customer', ''))
            
            # Sales Manager/Estimator Initials at C8
            estimator_name = project_data.get('estimator', '')
            if estimator_name:
                # Generate combined initials (Sales Contact + Estimator)
                from utils.word import get_combined_initials
                from config.business_data import SALES_CONTACTS
                
                # Get sales contact info based on estimator
                sales_contact_name = ""
                for contact_name, phone in SALES_CONTACTS.items():
                    if estimator_name and any(name.lower() in estimator_name.lower() for name in contact_name.split()):
                        sales_contact_name = contact_name
                        break
                
                # If no match found, use first sales contact
                if not sales_contact_name:
                    sales_contact_name = list(SALES_CONTACTS.keys())[0]
                
                # Generate combined initials
                combined_initials = get_combined_initials(sales_contact_name, estimator_name)
                write_to_cell_safe(sheet, 'C8', combined_initials)
            
            # Project Name at F4 (corrected from G4)
            write_to_cell_safe(sheet, 'F4', project_data.get('project_name', ''))
            
            # Location at F6 (corrected from G6)
            write_to_cell_safe(sheet, 'F6', project_data.get('project_location', ''))
            
            # Date at F8 (corrected from G8)
            write_to_cell_safe(sheet, 'F8', project_data.get('date', ''))
            
            # Revision at K9
            write_to_cell_safe(sheet, 'K8', project_data.get('revision', 'A'))
            
        except Exception as e:
            print(f"Warning: Could not write SDU project metadata: {str(e)}")
        
    except Exception as e:
        print(f"Warning: Could not write SDU metadata: {str(e)}")

def add_fire_suppression_dropdowns(sheet: Worksheet):
    """
    Add specific dropdowns for fire suppression sheets.
    
    Args:
        sheet (Worksheet): The fire suppression worksheet to add dropdowns to
    """
    try:
        # Fire suppression specific options (shortened to avoid Excel corruption)
        system_types = [
            "1 TANK SYSTEM",
            "1 TANK TRAVEL HUB",
            "1 TANK DISTANCE",
            "NOBEL",
            "AMAREX",
            "OTHER",
            "2 TANK SYSTEM",
            "2 TANK TRAVEL HUB",
            "2 TANK DISTANCE",
            "3 TANK SYSTEM",
            "3 TANK TRAVEL HUB",
            "3 TANK DISTANCE",
            "4 TANK SYSTEM",
            "4 TANK TRAVEL HUB",
            "4 TANK DISTANCE",
            "5 TANK SYSTEM",
            "5 TANK TRAVEL HUB",
            "5 TANK DISTANCE",
            "6 TANK SYSTEM",
            "6 TANK TRAVEL HUB",
            "6 TANK DISTANCE"
            ]
        
        tank_sizes = [
            "1 TANK",
            "1 TANK DISTANCE",
            "2 TANK",
            "2 TANK DISTANCE",
            "3 TANK",
            "3 TANK DISTANCE",
            "4 TANK",
            "5 TANK",
            "6 TANK"
        ]
        
        # Create data validations with proper formula length checking
        def create_validation(options):
            formula = ",".join(options)
            if len(formula) > 255:  # Excel formula limit
                truncated_options = []
                current_length = 0
                for opt in options:
                    if current_length + len(opt) + 1 > 250:
                        break
                    truncated_options.append(opt)
                    current_length += len(opt) + 1
                formula = ",".join(truncated_options)
            return DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
        
        system_dv = create_validation(system_types)
        tank_dv = create_validation(tank_sizes)
        
        # Add validations to sheet
        sheet.add_data_validation(system_dv)
        sheet.add_data_validation(tank_dv)
        
        # Apply to specific cells with error handling
        try:
            # Fire suppression system type (C16, C33, C50)
            system_dv.add("C16")
            system_dv.add("C33") 
            system_dv.add("C50")
            
            # Tank installation options (C17, C34, C51)
            tank_dv.add("C17")
            tank_dv.add("C34")
            tank_dv.add("C51")
        except Exception as e:
            print(f"Warning: Could not add fire suppression dropdown cells: {str(e)}")
        
    except Exception as e:
        print(f"Warning: Could not add fire suppression dropdowns to sheet {sheet.title}: {str(e)}")
        pass

def add_recoair_dropdowns(sheet: Worksheet):
    """
    Add specific dropdowns for RecoAir sheets.
    
    Args:
        sheet (Worksheet): The RecoAir worksheet to add dropdowns to
    """
    try:
        # RecoAir specific options
        internal_external_options = [
            "INTERNAL",
            "EXTERNAL"
        ]
        
        # Create data validation
        def create_validation(options):
            formula = ",".join(options)
            return DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
        
        internal_external_dv = create_validation(internal_external_options)
        
        # Add validation to sheet
        sheet.add_data_validation(internal_external_dv)
        
        # Apply to range I14 to I28
        try:
            for row in range(14, 29):  # 14 to 28 inclusive
                internal_external_dv.add(f"I{row}")
        except Exception as e:
            print(f"Warning: Could not add RecoAir dropdown cells: {str(e)}")
        
    except Exception as e:
        print(f"Warning: Could not add RecoAir dropdowns to sheet {sheet.title}: {str(e)}")
        pass

def add_sdu_dropdowns(sheet: Worksheet):
    """
    Add specific dropdowns for SDU sheets.
    
    Args:
        sheet (Worksheet): The SDU worksheet to add dropdowns to
    """
    try:
        # Water connection types for E89 to E91 (CWS/HWS)
        water_types_options = [
            "",  # Empty option
            "CWS",
            "HWS"
        ]
        
        # Water connection sizes for D89 to D91
        water_sizes_options = [
            "",  # Empty option
            "15mm",
            "22mm", 
            "28mm"
        ]
        
        # Create data validation
        def create_validation(options):
            formula = ",".join(options)
            if len(formula) > 255:  # Excel formula limit
                # Truncate if too long
                truncated_options = []
                current_length = 0
                for opt in options:
                    if current_length + len(opt) + 1 > 250:
                        break
                    truncated_options.append(opt)
                    current_length += len(opt) + 1
                formula = ",".join(truncated_options)
            return DataValidation(type="list", formula1=f'"{formula}"', allow_blank=True)
        
        water_types_dv = create_validation(water_types_options)
        water_sizes_dv = create_validation(water_sizes_options)
        
        # Add validations to sheet
        sheet.add_data_validation(water_types_dv)
        sheet.add_data_validation(water_sizes_dv)
        
        # Apply water types to cells E89 to E91 and set default to CWS
        try:
            for row in range(89, 92):  # E89 to E91 inclusive
                water_types_dv.add(f"E{row}")
                # Set default value to CWS
                sheet[f"E{row}"] = "CWS"
        except Exception as e:
            print(f"Warning: Could not add SDU water types dropdown cells: {str(e)}")
        
        # Apply water sizes to cells D89 to D91
        try:
            for row in range(89, 92):  # D89 to D91 inclusive
                water_sizes_dv.add(f"D{row}")
        except Exception as e:
            print(f"Warning: Could not add SDU water sizes dropdown cells: {str(e)}")
        
    except Exception as e:
        print(f"Warning: Could not add SDU dropdowns to sheet {sheet.title}: {str(e)}")
        pass

def organize_sheets_by_area(wb: Workbook):
    """
    Organize sheets so that JOB TOTAL is first, then related sheets (CANOPY, FIRE SUPP, etc.) are grouped by area.
    
    Args:
        wb (Workbook): The workbook to reorganize
    """
    try:
        # Get all sheet names and group them by area
        sheet_groups = {}
        other_sheets = []
        
        for sheet_name in wb.sheetnames:
            if ' - ' in sheet_name and '(' in sheet_name and ')' in sheet_name:
                # Extract area identifier: "CANOPY - Level 1 (1)" -> "Level 1 (1)"
                parts = sheet_name.split(' - ', 1)
                if len(parts) == 2:
                    sheet_type = parts[0]  # CANOPY, FIRE SUPP, etc.
                    area_identifier = parts[1]  # Level 1 (1)
                    
                    if area_identifier not in sheet_groups:
                        sheet_groups[area_identifier] = []
                    sheet_groups[area_identifier].append((sheet_type, sheet_name))
                else:
                    other_sheets.append(sheet_name)
            else:
                other_sheets.append(sheet_name)
        
        # Create ordered list of sheets
        ordered_sheets = []
        
        # Put JOB TOTAL first
        job_total_sheets = [s for s in other_sheets if 'JOB TOTAL' in s]
        ordered_sheets.extend(job_total_sheets)
        
        # Then add area-grouped sheets
        # Sort area identifiers to maintain consistent ordering
        for area_identifier in sorted(sheet_groups.keys()):
            # Sort sheets within each area: CANOPY first, then FIRE SUPP, then EBOX, then others
            area_sheets = sheet_groups[area_identifier]
            area_sheets.sort(key=lambda x: (
                0 if x[0] == 'CANOPY' else
                1 if x[0] == 'FIRE SUPP' else
                2 if x[0] == 'EBOX' else
                3 if x[0] == 'RECOAIR' else
                4 if x[0] == 'SDU' else 5,
                x[1]  # Then by sheet name as secondary sort
            ))
            
            # Add sheets from this area to the ordered list
            for _, sheet_name in area_sheets:
                ordered_sheets.append(sheet_name)
        
        # Add other sheets at the end (excluding JOB TOTAL which is already first)
        other_important_sheets = [s for s in other_sheets if s not in job_total_sheets and s != 'Lists']
        lists_sheets = [s for s in other_sheets if s == 'Lists']
        
        ordered_sheets.extend(other_important_sheets)
        ordered_sheets.extend(lists_sheets)
        
        # Reorder the sheets in the workbook
        current_sheets = wb.sheetnames.copy()
        
        # Move sheets to the correct order
        for i, target_sheet in enumerate(ordered_sheets):
            if target_sheet in current_sheets:
                current_index = wb.sheetnames.index(target_sheet)
                if current_index != i:
                    # Move the sheet to the correct position
                    sheet_to_move = wb[target_sheet]
                    wb.move_sheet(sheet_to_move, offset=i - current_index)
        
    except Exception as e:
        print(f"Warning: Could not organize sheets by area: {str(e)}")
        pass

def write_company_data_to_hidden_sheet(wb: Workbook, project_data: Dict):
    """
    Write company and estimator data to a hidden sheet for later extraction.
    
    Args:
        wb (Workbook): The workbook to add the hidden sheet to
        project_data (Dict): Project data including company and estimator info
    """
    try:
        # Create or get the hidden sheet
        sheet_name = "ProjectData"
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            sheet = wb.create_sheet(sheet_name)
        
        # Hide the sheet
        sheet.sheet_state = 'hidden'
        
        # Write company information
        sheet['A1'] = 'Company'
        sheet['B1'] = project_data.get('company', '')
        
        sheet['A2'] = 'Address'
        sheet['B2'] = project_data.get('address', '')
        
        # Write estimator information (full name, not initials)
        sheet['A3'] = 'Estimator_Full_Name'
        sheet['B3'] = project_data.get('estimator', '')
        
        # Get estimator rank from business data
        estimator_name = project_data.get('estimator', '')
        estimator_rank = 'Estimator'  # Default
        
        # Look up the rank from ESTIMATORS dictionary
        from config.business_data import ESTIMATORS
        for name, rank in ESTIMATORS.items():
            if name.lower() in estimator_name.lower():
                estimator_rank = rank
                break
        
        sheet['A4'] = 'Estimator_Rank'
        sheet['B4'] = estimator_rank
        
        # Write sales contact information
        sheet['A5'] = 'Sales_Contact'
        sheet['B5'] = project_data.get('sales_contact', '')
        
        # Write delivery location if available
        sheet['A6'] = 'Delivery_Location'
        sheet['B6'] = project_data.get('delivery_location', '')
        
        # Write revision information
        sheet['A7'] = 'Revision'
        sheet['B7'] = project_data.get('revision', 'A')
        
    except Exception as e:
        print(f"Warning: Could not write company data to hidden sheet: {str(e)}")
        pass

def write_delivery_location_to_sheet(sheet: Worksheet, delivery_location: str):
    """
    Write delivery location to cell D183 on the given sheet.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        delivery_location (str): The delivery location to write
    """
    try:
        if delivery_location and delivery_location != "Select...":
            sheet['D183'] = delivery_location
    except Exception as e:
        print(f"Warning: Could not write delivery location to sheet {sheet.title}: {str(e)}")
        pass

def save_to_excel(project_data: Dict) -> str:
    """
    Generate a complete Excel workbook from project data.
    
    Args:
        project_data (Dict): Complete project specification data
    
    Returns:
        str: Path to the saved Excel file
    """
    try:
        wb = load_template_workbook()
        project_type = project_data.get("project_type")
        if not project_type:
            raise ValueError("Project type not specified in project data")
        
        # Get all sheets once and create lists of available sheets
        all_sheets = wb.sheetnames
        canopy_sheets = [sheet for sheet in all_sheets if 'CANOPY' in sheet]
        fire_supp_sheets = [sheet for sheet in all_sheets if 'FIRE SUPP' in sheet or 'FIRE SUPPRESSION' in sheet]
        edge_box_sheets = [sheet for sheet in all_sheets if 'EBOX' in sheet or 'EDGE BOX' in sheet]
        recoair_sheets = [sheet for sheet in all_sheets if 'RECOAIR' in sheet]
        sdu_sheets = [sheet for sheet in all_sheets if 'SDU' in sheet and 'CANOPY' not in sheet and 'FIRE' not in sheet]
        
        # Hide the Lists sheet if it exists
        if 'Lists' in wb.sheetnames:
            wb['Lists'].sheet_state = 'hidden'
        
        # Add project metadata to JOB TOTAL sheet by default
        if 'JOB TOTAL' in wb.sheetnames:
            job_total_sheet = wb['JOB TOTAL']
            write_project_metadata(job_total_sheet, project_data)
            job_total_sheet.sheet_state = 'visible'
        
        # Write company and estimator data to hidden sheet
        write_company_data_to_hidden_sheet(wb, project_data)
        
        # Counters for sheet numbering
        recoair_sheet_count = 0
        
        # Keep track of total areas for coloring
        area_count = 0
        
        # Create a mapping of level names to area numbers for proper sheet numbering
        level_area_numbers = {}
        
        # Process each level and area
        for level in project_data.get("levels", []):
            level_number = level["level_number"]
            level_name = level.get("level_name", f"Level {level_number}")
            
            # Initialize area counter for this level if not exists
            if level_name not in level_area_numbers:
                level_area_numbers[level_name] = 0
            
            for idx, area in enumerate(level["areas"], 1):
                # Increment area number for this level
                level_area_numbers[level_name] += 1
                area_number = level_area_numbers[level_name]
                area_name = area["name"]
                area_canopies = area.get("canopies", [])
                
                # Get tab color for this area
                tab_color = TAB_COLORS[area_count % len(TAB_COLORS)]
                
                # Check if area has fire suppression
                has_fire_suppression = any(canopy.get("options", {}).get("fire_suppression", False) for canopy in area_canopies)
                
                # Check if area has UV-C system (area-level option)
                has_uvc = area.get("options", {}).get("uvc", False)
                
                # Check if area has SDU system (area-level option)
                has_sdu = area.get("options", {}).get("sdu", False)
                
                # Check if area has RecoAir system (area-level option)
                has_recoair = area.get("options", {}).get("recoair", False)
                
                current_canopy_sheet = None
                fs_sheet = None
                ebox_sheet = None
                sdu_sheet = None
                recoair_sheet = None
                
                # Process canopy sheet if canopies exist for this area
                if area_canopies:
                    if canopy_sheets:
                        sheet_name = canopy_sheets.pop(0)
                        current_canopy_sheet = wb[sheet_name]
                        
                        # Set title in B1
                        sheet_title_display = f"{level_name} - {area_name}"
                        current_canopy_sheet['B1'] = sheet_title_display
                        
                        # Rename the sheet tab
                        canopy_sheet_tab_name = f"CANOPY - {level_name} ({area_number})"
                        current_canopy_sheet.title = canopy_sheet_tab_name
                        current_canopy_sheet.sheet_state = 'visible'
                        current_canopy_sheet.sheet_properties.tabColor = tab_color
                        
                        # Write project metadata to canopy sheet (C/G columns)
                        write_project_metadata(current_canopy_sheet, project_data)
                        
                        # Create fire suppression sheet if needed
                        if has_fire_suppression:
                            if fire_supp_sheets:
                                fs_sheet_name = fire_supp_sheets.pop(0)
                                fs_sheet = wb[fs_sheet_name]
                                new_fs_name = f"FIRE SUPP - {level_name} ({area_number})"
                                fs_sheet.title = new_fs_name
                                fs_sheet.sheet_state = 'visible'
                                fs_sheet.sheet_properties.tabColor = tab_color
                                
                                # Write project metadata to fire suppression sheet
                                write_project_metadata(fs_sheet, project_data)
                                # Set fire suppression sheet title in B1
                                fs_sheet['B1'] = f"{level_name} - {area_name} - FIRE SUPPRESSION"
                        
                        # Create EBOX sheet if UV-C is selected for this area
                        if has_uvc:
                            if edge_box_sheets:
                                ebox_sheet_name = edge_box_sheets.pop(0)
                                ebox_sheet = wb[ebox_sheet_name]
                                new_ebox_name = f"EBOX - {level_name} ({area_number})"
                                ebox_sheet.title = new_ebox_name
                                ebox_sheet.sheet_state = 'visible'
                                ebox_sheet.sheet_properties.tabColor = tab_color
                                
                                # Write EBOX-specific metadata to EBOX sheet (D/H columns)
                                write_ebox_metadata(ebox_sheet, project_data)
                                # Set EBOX sheet title in B1
                                ebox_sheet['B1'] = f"{level_name} - {area_name} - UV-C SYSTEM"
                            else:
                                print(f"Warning: Not enough EBOX sheets in template for UV-C system in area {area_name}")
                        
                        # Create SDU sheet if SDU is selected for this area
                        if has_sdu:
                            if sdu_sheets:
                                sdu_sheet_name = sdu_sheets.pop(0)
                                sdu_sheet = wb[sdu_sheet_name]
                                new_sdu_name = f"SDU - {level_name} ({area_number})"
                                sdu_sheet.title = new_sdu_name
                                sdu_sheet.sheet_state = 'visible'
                                sdu_sheet.sheet_properties.tabColor = tab_color
                                
                                # Write SDU-specific metadata to SDU sheet (C/G columns)
                                write_sdu_metadata(sdu_sheet, project_data)
                                # Set SDU sheet title in B1
                                sdu_sheet['B1'] = f"{level_name} - {area_name} - SDU SYSTEM"
                                
                                # Add SDU specific dropdowns
                                add_sdu_dropdowns(sdu_sheet)
                            else:
                                print(f"Warning: Not enough SDU sheets in template for SDU system in area {area_name}")
                        
                        # Create RECOAIR sheet if RecoAir is selected for this area
                        if has_recoair:
                            if recoair_sheets:
                                recoair_sheet_name = recoair_sheets.pop(0)
                                recoair_sheet = wb[recoair_sheet_name]
                                new_recoair_name = f"RECOAIR - {level_name} ({area_number})"
                                recoair_sheet.title = new_recoair_name
                                recoair_sheet.sheet_state = 'visible'
                                recoair_sheet.sheet_properties.tabColor = tab_color
                                
                                # Generate item number for this RecoAir sheet
                                recoair_sheet_count += 1
                                item_number = f"{recoair_sheet_count}.01"
                                
                                # Write RECOAIR-specific metadata to RECOAIR sheet (D/H columns)
                                write_recoair_metadata(recoair_sheet, project_data, item_number)
                                # Set RECOAIR sheet title in C1
                                recoair_sheet['C1'] = f"{level_name} - {area_name} - RECOAIR SYSTEM"
                                
                                # Add RecoAir specific dropdowns
                                add_recoair_dropdowns(recoair_sheet)
                            else:
                                print(f"Warning: Not enough RECOAIR sheets in template for RecoAir system in area {area_name}")
                        
                        # Write each canopy with proper spacing
                        fs_canopy_idx = 0  # Track fire suppression canopies separately
                        for canopy_idx, canopy in enumerate(area_canopies):
                            row_start = CANOPY_START_ROW + (canopy_idx * CANOPY_ROW_SPACING)
                            write_canopy_data(current_canopy_sheet, canopy, row_start)
                            
                            # If this canopy has fire suppression and fire suppression sheet exists, write to it
                            if canopy.get("options", {}).get("fire_suppression") and fs_sheet:
                                fs_row_start = CANOPY_START_ROW + (fs_canopy_idx * CANOPY_ROW_SPACING)
                                write_fire_suppression_canopy_data(fs_sheet, canopy, fs_row_start)
                                fs_canopy_idx += 1  # Only increment for canopies with fire suppression
                        
                        # Add dropdowns
                        add_dropdowns_to_sheet(wb, current_canopy_sheet)
                        if fs_sheet:
                            # Add fire suppression specific dropdowns
                            add_fire_suppression_dropdowns(fs_sheet)
                    else:
                        raise Exception(f"Not enough CANOPY sheets in template for area {area_name}")
                
                # Handle case where UV-C, SDU, and/or RecoAir are selected but no canopies exist (edge case)
                elif (has_uvc or has_sdu or has_recoair) and not area_canopies:
                    # Create EBOX sheet if UV-C is selected
                    if has_uvc:
                        if edge_box_sheets:
                            ebox_sheet_name = edge_box_sheets.pop(0)
                            ebox_sheet = wb[ebox_sheet_name]
                            new_ebox_name = f"EBOX - {level_name} ({area_number})"
                            ebox_sheet.title = new_ebox_name
                            ebox_sheet.sheet_state = 'visible'
                            ebox_sheet.sheet_properties.tabColor = tab_color
                            
                            # Write EBOX-specific metadata to EBOX sheet (D/H columns)
                            write_ebox_metadata(ebox_sheet, project_data)
                            # Set EBOX sheet title in C1
                            ebox_sheet['C1'] = f"{level_name} - {area_name} - UV-C SYSTEM"
                        else:
                            print(f"Warning: Not enough EBOX sheets in template for UV-C system in area {area_name}")
                    
                    # Create SDU sheet if SDU is selected
                    if has_sdu:
                        if sdu_sheets:
                            sdu_sheet_name = sdu_sheets.pop(0)
                            sdu_sheet = wb[sdu_sheet_name]
                            new_sdu_name = f"SDU - {level_name} ({area_number})"
                            sdu_sheet.title = new_sdu_name
                            sdu_sheet.sheet_state = 'visible'
                            sdu_sheet.sheet_properties.tabColor = tab_color
                            
                            # Write SDU-specific metadata to SDU sheet (C/G columns)
                            write_sdu_metadata(sdu_sheet, project_data)
                            # Set SDU sheet title in B1
                            sdu_sheet['B1'] = f"{level_name} - {area_name} - SDU SYSTEM"
                            
                            # Add SDU specific dropdowns
                            add_sdu_dropdowns(sdu_sheet)
                        else:
                            print(f"Warning: Not enough SDU sheets in template for SDU system in area {area_name}")
                    
                    # Create RECOAIR sheet if RecoAir is selected
                    if has_recoair:
                        if recoair_sheets:
                            recoair_sheet_name = recoair_sheets.pop(0)
                            recoair_sheet = wb[recoair_sheet_name]
                            new_recoair_name = f"RECOAIR - {level_name} ({area_number})"
                            recoair_sheet.title = new_recoair_name
                            recoair_sheet.sheet_state = 'visible'
                            recoair_sheet.sheet_properties.tabColor = tab_color
                            
                            # Generate item number for this RecoAir sheet
                            recoair_sheet_count += 1
                            item_number = f"{recoair_sheet_count}.01"
                            
                            # Write RECOAIR-specific metadata to RECOAIR sheet (D/H columns)
                            write_recoair_metadata(recoair_sheet, project_data, item_number)
                            # Set RECOAIR sheet title in C1
                            recoair_sheet['C1'] = f"{level_name} - {area_name} - RECOAIR SYSTEM"
                            
                            # Add RecoAir specific dropdowns
                            add_recoair_dropdowns(recoair_sheet)
                        else:
                            print(f"Warning: Not enough RECOAIR sheets in template for RecoAir system in area {area_name}")
                
                area_count += 1
        
        # Write project metadata to any other visible sheets that might exist
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if (sheet.sheet_state == 'visible' and 
                not sheet_name.startswith(('CANOPY', 'FIRE SUPP', 'EBOX', 'RECOAIR', 'SDU')) and 
                sheet_name not in ['Lists', 'JOB TOTAL']):
                # Write metadata to any other visible sheets (excluding EBOX, RECOAIR, and SDU which have their own metadata)
                try:
                    write_project_metadata(sheet, project_data)
                except Exception as e:
                    print(f"Warning: Could not write metadata to sheet {sheet_name}: {str(e)}")
        
        # Organize sheets by area for better navigation
        organize_sheets_by_area(wb)
        
        # Write delivery location to D183 for all sheets except JOB TOTAL, EBOX, RECOAIR, and SDU
        delivery_location = project_data.get('delivery_location', '')
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if (sheet.sheet_state == 'visible' and 
                sheet_name != 'JOB TOTAL' and 
                sheet_name != 'Lists' and
                not sheet_name.startswith(('EBOX', 'RECOAIR', 'SDU'))):
                write_delivery_location_to_sheet(sheet, delivery_location)
        
        # Remove any unused template sheets
        unused_sheets = canopy_sheets + fire_supp_sheets + edge_box_sheets + sdu_sheets + recoair_sheets
        for sheet_name in unused_sheets:
            if sheet_name in wb.sheetnames:
                wb.remove(wb[sheet_name])
        
        # Create pricing summary sheet for dynamic pricing aggregation
        print("Creating PRICING_SUMMARY sheet...")
        create_pricing_summary_sheet(wb)
        
        # Update JOB TOTAL sheet to reference pricing summary
        print("Updating JOB TOTAL sheet with dynamic pricing formulas...")
        update_job_total_sheet(wb)
        
        # Save the workbook
        project_number = project_data.get('project_number', 'unknown')
        date_str = project_data.get('date', '')
        
        # Format date for filename (remove slashes and make it filename-safe)
        if date_str:
            # Convert DD/MM/YYYY to DDMMYYYY or similar format
            formatted_date = date_str.replace('/', '').replace('-', '')
        else:
            # Use current date if no date provided
            formatted_date = datetime.now().strftime("%d%m%Y")
        
        output_path = f"output/{project_number} Cost Sheet {formatted_date}.xlsx"
        os.makedirs("output", exist_ok=True)
        wb.save(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Failed to generate Excel file: {str(e)}")

def read_excel_project_data(excel_path: str) -> Dict:
    """
    Read project data back from a generated Excel file.
    
    Args:
        excel_path (str): Path to the Excel file to read
        
    Returns:
        Dict: Project data extracted from the Excel file
    """
    try:
        wb = load_workbook(excel_path, data_only=True)
        
        # Try to get data from JOB TOTAL sheet first, then any canopy sheet
        data_sheet = None
        if 'JOB TOTAL' in wb.sheetnames:
            data_sheet = wb['JOB TOTAL']
        else:
            # Find first CANOPY sheet
            for sheet_name in wb.sheetnames:
                if 'CANOPY' in sheet_name:
                    data_sheet = wb[sheet_name]
                    break
        
        if not data_sheet:
            raise Exception("No suitable sheet found to extract project data")
        
        # Extract project metadata using the same cell mappings
        project_data = {}
        
        # Read basic project info
        project_data['project_number'] = data_sheet['C3'].value or ""
        project_data['customer'] = data_sheet['C5'].value or ""
        project_data['estimator_initials'] = data_sheet['C7'].value or ""  # This is the initials version
        project_data['project_name'] = data_sheet['G3'].value or ""
        project_data['project_location'] = data_sheet['G5'].value or ""  # Project location from G5
        project_data['location'] = data_sheet['G5'].value or ""  # Keep for backward compatibility
        project_data['date'] = data_sheet['G7'].value or ""
        project_data['revision'] = data_sheet['O7'].value or "A"  # Revision from O7, default to "A"
        
        # Read company and estimator data from hidden ProjectData sheet
        if 'ProjectData' in wb.sheetnames:
            hidden_sheet = wb['ProjectData']
            
            # Read company information
            project_data['company'] = hidden_sheet['B1'].value or ""
            project_data['address'] = hidden_sheet['B2'].value or ""
            
            # Read full estimator information
            project_data['estimator'] = hidden_sheet['B3'].value or project_data['estimator_initials']
            project_data['estimator_rank'] = hidden_sheet['B4'].value or "Estimator"
            
            # Read additional data
            project_data['sales_contact'] = hidden_sheet['B5'].value or ""
            project_data['delivery_location'] = hidden_sheet['B6'].value or ""
            
            # Read revision from ProjectData sheet if not already set
            if not project_data.get('revision') or project_data['revision'] == 'A':
                project_data['revision'] = hidden_sheet['B7'].value or project_data.get('revision', 'A')
        else:
            # Fallback if no hidden sheet exists
            project_data['estimator'] = project_data['estimator_initials']
            project_data['estimator_rank'] = "Estimator"
            project_data['company'] = ""
            project_data['address'] = ""
            project_data['sales_contact'] = ""
            project_data['delivery_location'] = ""
        
        # Get level and area information from sheet titles
        levels_data = {}
        canopy_data = {}
        
        # First pass: Create areas from all sheet types (CANOPY, FIRE SUPP, EBOX, RECOAIR, SDU)
        for sheet_name in wb.sheetnames:
            if any(prefix in sheet_name for prefix in ['CANOPY - ', 'FIRE SUPP - ', 'EBOX - ', 'RECOAIR - ', 'SDU - ']):
                sheet = wb[sheet_name]
                
                # Determine which cell contains the title based on sheet type
                if 'EBOX - ' in sheet_name or 'RECOAIR - ' in sheet_name:
                    title_cell = sheet['C1'].value  # EBOX and RECOAIR sheets have title in C1
                else:
                    title_cell = sheet['B1'].value  # CANOPY, FIRE SUPP, and SDU sheets have title in B1
                
                if title_cell and ' - ' in title_cell:
                    # Handle different title formats
                    if 'UV-C SYSTEM' in title_cell or 'RECOAIR SYSTEM' in title_cell or 'SDU SYSTEM' in title_cell:
                        # For EBOX/RECOAIR/SDU: "Level 1 - Main Kitchen - UV-C SYSTEM" or "Level 1 - Main Kitchen - RECOAIR SYSTEM" or "Level 1 - Main Kitchen - SDU SYSTEM"
                        title_parts = title_cell.split(' - ')
                        if len(title_parts) >= 2:
                            level_name = title_parts[0]
                            area_name = title_parts[1]
                        else:
                            continue
                    else:
                        # For CANOPY/FIRE SUPP: "Level 1 - Main Kitchen" or "Level 1 - Main Kitchen - FIRE SUPPRESSION"
                        level_area = title_cell.split(' - ')
                        if len(level_area) >= 2:
                            level_name = level_area[0]
                            area_name = level_area[1]
                        else:
                            continue
                    
                    # Create level if it doesn't exist
                    if level_name not in levels_data:
                        levels_data[level_name] = []
                    
                    # Create area if it doesn't exist
                    if area_name not in [area['name'] for area in levels_data[level_name]]:
                        levels_data[level_name].append({
                            'name': area_name,
                            'canopies': []
                        })
        
        # Second pass: Read canopy data from CANOPY sheets
        for sheet_name in wb.sheetnames:
            if 'CANOPY - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value
                
                if title_cell and ' - ' in title_cell:
                    level_area = title_cell.split(' - ')
                    if len(level_area) >= 2:
                        level_name = level_area[0]
                        area_name = level_area[1]
                        
                        # Read canopy specifications from the sheet
                        # This is a simplified read - you might want to enhance this
                        for canopy_idx in range(10):  # Support up to 10 canopies
                            base_row = CANOPY_START_ROW + (canopy_idx * CANOPY_ROW_SPACING)
                            ref_row = base_row - 2
                            
                            ref_number = sheet[f'B{ref_row}'].value
                            if ref_number:
                                # Get model to check for placeholder rows
                                model = sheet[f'D{base_row}'].value or ""
                                
                                # Skip placeholder rows
                                if (ref_number.upper() == "ITEM" or 
                                    model.upper() == "CANOPY TYPE" or
                                    ref_number.upper().strip() == "ITEM" or
                                    model.upper().strip() == "CANOPY TYPE"):
                                    continue
                                
                                canopy_info = {
                                    'reference_number': ref_number,
                                    'configuration': sheet[f'C{base_row}'].value or "",
                                    'model': model,
                                    
                                    # Additional specification data
                                    'length': sheet[f'F{base_row}'].value or "",
                                    'width': sheet[f'E{base_row}'].value or "",
                                    'height': sheet[f'G{base_row}'].value or "",
                                    'sections': sheet[f'H{base_row}'].value or "",
                                    'lighting_type': sheet[f'C{base_row + 1}'].value or "",  # C15 (base_row + 1)
                                    
                                    # Volume and static data (if available in your template)
                                    'extract_volume': sheet[f'I{base_row}'].value or "",
                                    'extract_static': sheet[f'F{base_row + 8}'].value or "",  # F22, F39, F56, etc.
                                    'mua_volume': sheet[f'K{base_row}'].value or "",
                                    'supply_static': sheet[f'L{base_row}'].value or "",
                                    
                                    # Pricing data - individual canopy price
                                    'canopy_price': sheet[f'P{ref_row}'].value or 0,  # P12, P29, P46, etc. (ref_row = base_row - 2)
                                    
                                    # Fire suppression data - will be populated from FIRE SUPP sheet
                                    'fire_suppression_tank_quantity': 0,  # Default to 0, will be updated from FIRE SUPP sheet
                                    'fire_suppression_price': 0,  # Default to 0, will be updated from FIRE SUPP sheet
                                    'fire_suppression_system_type': None,  # Default to None, will be updated from FIRE SUPP sheet
                                    
                                    # Read wall cladding data from Excel
                                    'wall_cladding': read_wall_cladding_from_canopy(sheet, base_row),
                                    
                                    # Read wall cladding price from Excel (N19, N20, N21, etc.)
                                    'cladding_price': sheet[f'N{base_row + 5}'].value or 0  # N19, N36, N53, etc. (base_row + 5)
                                }
                                
                                # Add CWS/HWS data for CMWF and CMWI canopies
                                if model.upper() in ['CMWF', 'CMWI']:
                                    # Calculate the row offset for CWS/HWS data (F25, F26, F27 relative to canopy)
                                    cws_row = base_row + 11  # F25 relative to base_row (14 + 11 = 25)
                                    hws_row = base_row + 12  # F26 relative to base_row
                                    storage_row = base_row + 13  # F27 relative to base_row
                                    
                                    canopy_info.update({
                                        'cws_capacity': sheet[f'F{cws_row}'].value or "",  # CWS @ 2 Bar
                                        'hws_requirement': sheet[f'F{hws_row}'].value or "",  # HWS @ 2 Bar  
                                        'hw_storage': sheet[f'F{storage_row}'].value or "",  # HWS Storage
                                        'has_wash_capabilities': True
                                    })
                                else:
                                    canopy_info.update({
                                        'cws_capacity': "",
                                        'hws_requirement': "",
                                        'hw_storage': "",
                                        'has_wash_capabilities': False
                                    })
                                
                                # Find the area and add canopy data
                                for area in levels_data[level_name]:
                                    if area['name'] == area_name:
                                        area['canopies'].append(canopy_info)
                                        break
        
        # Read fire suppression data from FIRE SUPP sheets
        for sheet_name in wb.sheetnames:
            if 'FIRE SUPP - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value
                
                if title_cell and ' - ' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - FIRE SUPPRESSION"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Get fire suppression commissioning price from N193 and delivery price from N182
                        fs_commissioning_price = sheet['N193'].value or 0
                        fs_delivery_price = sheet['N182'].value or 0
                        
                        # Count how many fire suppression units are in this sheet
                        fs_units = []
                        for canopy_idx in range(10):  # Support up to 10 canopies
                            base_row = CANOPY_START_ROW + (canopy_idx * CANOPY_ROW_SPACING)
                            ref_row = base_row - 2
                            system_row = base_row + 2  # C16 relative to base_row (14 + 2 = 16)
                            tank_row = base_row + 3  # C17 relative to base_row (14 + 3 = 17)
                            
                            ref_number = sheet[f'B{ref_row}'].value
                            system_type = sheet[f'C{system_row}'].value  # Fire suppression system type from C16
                            tank_value = sheet[f'C{tank_row}'].value
                            base_fire_suppression_price = sheet[f'N{ref_row}'].value or 0  # Fire suppression base price at N12, N29, N46, etc.
                            
                            # Only count actual fire suppression units, not template entries
                            if (ref_number and tank_value and 
                                str(ref_number).upper() != "ITEM" and 
                                str(tank_value).upper() not in ["TANK INSTALL", "TANK INSTALLATION"]):
                                tank_quantity = extract_tank_quantity(tank_value)
                                fs_units.append({
                                    'ref_number': ref_number,
                                    'system_type': system_type,  # Add system type from C16
                                    'tank_quantity': tank_quantity,
                                    'base_price': base_fire_suppression_price
                                })
                        
                        # Calculate delivery price per unit (split equally among all units, or full amount if only one unit)
                        if len(fs_units) == 1:
                            delivery_per_unit = fs_delivery_price  # Single unit gets full delivery price
                        else:
                            delivery_per_unit = fs_delivery_price / len(fs_units) if fs_units else 0  # Multiple units split delivery
                        
                        # Update fire suppression data for each canopy
                        for fs_unit in fs_units:
                            # Calculate fire suppression price: base price + delivery share (no commissioning share)
                            total_fs_price = fs_unit['base_price'] + delivery_per_unit
                            
                            # Find the corresponding canopy and update its fire suppression data
                            for level_areas in levels_data.get(level_name, []):
                                if level_areas['name'] == area_name:
                                    for canopy in level_areas['canopies']:
                                        if canopy['reference_number'] == fs_unit['ref_number']:
                                            canopy['fire_suppression_tank_quantity'] = fs_unit['tank_quantity']
                                            canopy['fire_suppression_price'] = total_fs_price
                                            canopy['fire_suppression_system_type'] = fs_unit['system_type']  # Add system type
                                            break
        
        # Read area-level pricing data (delivery & installation, commissioning)
        for sheet_name in wb.sheetnames:
            if 'CANOPY - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value
                
                if title_cell and ' - ' in title_cell:
                    level_area = title_cell.split(' - ')
                    if len(level_area) >= 2:
                        level_name = level_area[0]
                        area_name = level_area[1]
                        
                        # Read area-level pricing
                        delivery_installation_price = sheet['P182'].value or 0
                        commissioning_price = sheet['N193'].value or 0
                        
                        # Find the area and add pricing data
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                level_areas.update({
                                    'delivery_installation_price': delivery_installation_price,
                                    'commissioning_price': commissioning_price
                                })
                                break
        
        # Read UV-C pricing from EBOX sheets
        for sheet_name in wb.sheetnames:
            if 'EBOX - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['C1'].value  # EBOX sheets have title in C1
                
                if title_cell and ' - ' in title_cell and 'UV-C SYSTEM' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - UV-C SYSTEM"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Read UV-C price from N9
                        uvc_price = sheet['N9'].value or 0
                        
                        # Find the area and add UV-C pricing data
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                level_areas.update({
                                    'uvc_price': uvc_price
                                })
                                break
        
        # Read RecoAir data from RECOAIR sheets
        for sheet_name in wb.sheetnames:
            if 'RECOAIR - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['C1'].value  # RECOAIR sheets have title in C1
                
                if title_cell and ' - ' in title_cell and 'RECOAIR SYSTEM' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - RECOAIR SYSTEM"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Read RecoAir data from this sheet (now returns dict with units and flat pack)
                        recoair_data = read_recoair_data_from_sheet(sheet)
                        recoair_units = recoair_data['units']
                        flat_pack_data = recoair_data['flat_pack']
                        
                        # Read RecoAir commissioning price from N46
                        recoair_commissioning_price = sheet['N46'].value or 0
                        
                        # Calculate total RecoAir price (sum of all unit prices + delivery + commissioning + flat pack)
                        total_unit_price = sum(unit['unit_price'] for unit in recoair_units)
                        total_delivery_price = sum(unit['delivery_installation_price'] for unit in recoair_units)
                        flat_pack_price = flat_pack_data['price'] if flat_pack_data['has_flat_pack'] else 0
                        recoair_price = total_unit_price + total_delivery_price + recoair_commissioning_price + flat_pack_price
                        
                        # Find the area and add RecoAir data
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                level_areas.update({
                                    'recoair_price': recoair_price,
                                    'recoair_commissioning_price': recoair_commissioning_price,  # Add commissioning price separately
                                    'recoair_units': recoair_units,  # Add detailed unit data
                                    'recoair_flat_pack': flat_pack_data  # Add flat pack data
                                })
                                break
        
        # Read area-level options from sheets
        for sheet_name in wb.sheetnames:
            if 'CANOPY - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value
                
                if title_cell and ' - ' in title_cell:
                    level_area = title_cell.split(' - ')
                    if len(level_area) >= 2:
                        level_name = level_area[0]
                        area_name = level_area[1]
                        
                        # Initialize area options
                        area_options = {'uvc': False, 'sdu': False, 'recoair': False}
                        
                        # Check for area options in rows 6-8 (where write_area_options writes them)
                        for row in range(6, 9):
                            cell_value = sheet[f'B{row}'].value
                            if cell_value:
                                cell_value_upper = str(cell_value).upper()
                                if 'UV-C' in cell_value_upper:
                                    area_options['uvc'] = True
                                elif 'SDU' in cell_value_upper:
                                    area_options['sdu'] = True
                                elif 'RECOAIR' in cell_value_upper:
                                    area_options['recoair'] = True
                        
                        # Find the area and add options data
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                level_areas['options'] = area_options
                                break
            
            # Also check EBOX sheets for UV-C option
            elif 'EBOX - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['C1'].value  # EBOX sheets have title in C1
                
                if title_cell and ' - ' in title_cell and 'UV-C SYSTEM' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - UV-C SYSTEM"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Find the area and set UV-C option to True
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                if 'options' not in level_areas:
                                    level_areas['options'] = {'uvc': False, 'sdu': False, 'recoair': False}
                                level_areas['options']['uvc'] = True
                                break
            
            # Also check SDU sheets for SDU option
            elif 'SDU - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value  # SDU sheets have title in B1
                
                if title_cell and ' - ' in title_cell and 'SDU SYSTEM' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - SDU SYSTEM"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Find the area and set SDU option to True
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                if 'options' not in level_areas:
                                    level_areas['options'] = {'uvc': False, 'sdu': False, 'recoair': False}
                                level_areas['options']['sdu'] = True
                                break
            
            # Also check RECOAIR sheets for RecoAir option
            elif 'RECOAIR - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['C1'].value  # RECOAIR sheets have title in C1
                
                if title_cell and ' - ' in title_cell and 'RECOAIR SYSTEM' in title_cell:
                    # Extract level and area from title like "Level 1 - Main Kitchen - RECOAIR SYSTEM"
                    title_parts = title_cell.split(' - ')
                    if len(title_parts) >= 2:
                        level_name = title_parts[0]
                        area_name = title_parts[1]
                        
                        # Find the area and set RecoAir option to True
                        for level_areas in levels_data.get(level_name, []):
                            if level_areas['name'] == area_name:
                                if 'options' not in level_areas:
                                    level_areas['options'] = {'uvc': False, 'sdu': False, 'recoair': False}
                                level_areas['options']['recoair'] = True
                                break
        
        # Convert levels_data to the format expected by the system
        project_data['levels'] = []
        for level_idx, (level_name, areas) in enumerate(levels_data.items(), 1):
            project_data['levels'].append({
                'level_number': level_idx,  # Add level_number field required by save_to_excel
                'level_name': level_name,
                'areas': areas
            })
        

        
        return project_data
        
    except Exception as e:
        raise Exception(f"Failed to read Excel project data: {str(e)}")

def collect_wall_cladding_data(project_data: Dict) -> List[Dict]:
    """
    Collect all wall cladding data from canopies across all levels and areas.
    
    Args:
        project_data (Dict): Project data containing levels and areas
        
    Returns:
        List[Dict]: List of wall cladding specifications with item numbers
    """
    cladding_data = []
    
    levels = project_data.get("levels", [])
    for level in levels:
        for area in level.get("areas", []):
            for canopy in area.get("canopies", []):
                wall_cladding = canopy.get("wall_cladding", {})
                
                # Check if this canopy has wall cladding
                if wall_cladding.get("type") != "None" and wall_cladding.get("type"):
                    # Handle position as list or string
                    position = wall_cladding.get("position", [])
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
                    
                    # Join positions for other uses (use "and" format for consistency)
                    position_str = " and ".join(position_list) if position_list else ""
                    
                    cladding_info = {
                        'item_number': canopy.get("reference_number", ""),  # Use canopy reference number
                        'description': description,
                        'width': wall_cladding.get("width", 0),
                        'height': wall_cladding.get("height", 0),
                        'dimensions': f"{wall_cladding.get('width', 0)}X{wall_cladding.get('height', 0)}",
                        'position_description': position_str,
                        'canopy_ref': canopy.get("reference_number", ""),
                        'level_name': level.get("level_name", ""),
                        'area_name': area.get("name", "")
                    }
                    
                    cladding_data.append(cladding_info)
    
    return cladding_data

def write_wall_cladding_summary(sheet: Worksheet, cladding_data: List[Dict]):
    """
    Write wall cladding summary to the sheet starting at row 19.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        cladding_data (List[Dict]): List of wall cladding specifications
    """
    if not cladding_data:
        return
    
    # Starting row for wall cladding section
    start_row = 19
    
    # Write "2M² (HFL)" in C19 to indicate cladding is present
    sheet[f"C{start_row}"] = "2M² (HFL)"
    
    # Write each cladding item
    for idx, cladding in enumerate(cladding_data):
        current_row = start_row + idx
        
        # Item number in column A (if needed)
        # sheet[f"A{current_row}"] = cladding['item_number']
        
        # Description in column B (or appropriate column based on template)
        # sheet[f"B{current_row}"] = cladding['description']
        
        # Dimensions in column P
        sheet[f"P{current_row}"] = cladding['dimensions']
        
        # Position description in column Q  
        sheet[f"Q{current_row}"] = cladding['position_description'] 

# Add this new function after the write_delivery_location_to_sheet function

def create_pricing_summary_sheet(wb: Workbook) -> None:
    """
    Create a hidden PRICING_SUMMARY sheet that aggregates totals from all sheets.
    Uses Excel formulas to reference N9 cells from individual sheets for dynamic updates.
    
    Args:
        wb (Workbook): The workbook to add the pricing summary sheet to
    """
    try:
        # Create or get the PRICING_SUMMARY sheet
        sheet_name = "PRICING_SUMMARY"
        if sheet_name in wb.sheetnames:
            wb.remove(wb[sheet_name])  # Remove existing sheet to recreate
        
        summary_sheet = wb.create_sheet(sheet_name)
        summary_sheet.sheet_state = 'hidden'  # Hide the sheet
        
        # Set up headers
        summary_sheet['A1'] = 'Sheet Type'
        summary_sheet['B1'] = 'Sheet Name'
        summary_sheet['C1'] = 'Total Price (N9)'
        summary_sheet['D1'] = 'Total Cost (K9)'
        summary_sheet['E1'] = 'Price Formula Reference'
        summary_sheet['F1'] = 'Cost Formula Reference'
        
        # Get all visible sheets and categorize them
        canopy_sheets = []
        fire_supp_sheets = []
        ebox_sheets = []
        sdu_sheets = []
        recoair_sheets = []
        other_sheets = []
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if sheet.sheet_state == 'visible':
                if 'CANOPY - ' in sheet_name:
                    canopy_sheets.append(sheet_name)
                elif 'FIRE SUPP - ' in sheet_name:
                    fire_supp_sheets.append(sheet_name)
                elif 'EBOX - ' in sheet_name:
                    ebox_sheets.append(sheet_name)
                elif 'SDU - ' in sheet_name:
                    sdu_sheets.append(sheet_name)
                elif 'RECOAIR - ' in sheet_name:
                    recoair_sheets.append(sheet_name)
                elif sheet_name not in ['JOB TOTAL', 'Lists', 'PRICING_SUMMARY', 'ProjectData']:
                    other_sheets.append(sheet_name)
        
        # Write individual sheet references
        current_row = 2
        
        # CANOPY sheets
        for sheet_name in canopy_sheets:
            summary_sheet[f'A{current_row}'] = 'CANOPY'
            summary_sheet[f'B{current_row}'] = sheet_name
            # Create formulas to reference N9 (price) and K9 (cost) from the sheet
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!N9,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!K9,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!N9"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!K9"  # Cost reference
            current_row += 1
        
        # FIRE SUPP sheets
        for sheet_name in fire_supp_sheets:
            summary_sheet[f'A{current_row}'] = 'FIRE SUPP'
            summary_sheet[f'B{current_row}'] = sheet_name
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!N9,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!K9,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!N9"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!K9"  # Cost reference
            current_row += 1
        
        # EBOX sheets
        for sheet_name in ebox_sheets:
            summary_sheet[f'A{current_row}'] = 'EBOX'
            summary_sheet[f'B{current_row}'] = sheet_name
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!N9,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!K9,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!N9"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!K9"  # Cost reference
            current_row += 1
        
        # SDU sheets
        for sheet_name in sdu_sheets:
            summary_sheet[f'A{current_row}'] = 'SDU'
            summary_sheet[f'B{current_row}'] = sheet_name
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!J10,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!G10,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!J10"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!G10"  # Cost reference
            current_row += 1
        
        # RECOAIR sheets
        for sheet_name in recoair_sheets:
            summary_sheet[f'A{current_row}'] = 'RECOAIR'
            summary_sheet[f'B{current_row}'] = sheet_name
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!N9,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!K9,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!N9"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!K9"  # Cost reference
            current_row += 1
        
        # OTHER sheets
        for sheet_name in other_sheets:
            summary_sheet[f'A{current_row}'] = 'OTHER'
            summary_sheet[f'B{current_row}'] = sheet_name
            safe_sheet_name = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
            summary_sheet[f'C{current_row}'] = f"=IFERROR({safe_sheet_name}!N9,0)"  # Price
            summary_sheet[f'D{current_row}'] = f"=IFERROR({safe_sheet_name}!K9,0)"  # Cost
            summary_sheet[f'E{current_row}'] = f"{safe_sheet_name}!N9"  # Price reference
            summary_sheet[f'F{current_row}'] = f"{safe_sheet_name}!K9"  # Cost reference
            current_row += 1
        
        # Add summary totals by type
        summary_row = current_row + 2
        summary_sheet[f'A{summary_row}'] = 'SUMMARY TOTALS'
        summary_sheet[f'B{summary_row}'] = 'PRICE TOTALS'
        summary_sheet[f'C{summary_row}'] = 'COST TOTALS'
        
        summary_sheet[f'A{summary_row + 1}'] = 'CANOPY TOTAL'
        summary_sheet[f'A{summary_row + 2}'] = 'FIRE SUPP TOTAL'
        summary_sheet[f'A{summary_row + 3}'] = 'EBOX TOTAL'
        summary_sheet[f'A{summary_row + 4}'] = 'SDU TOTAL'
        summary_sheet[f'A{summary_row + 5}'] = 'RECOAIR TOTAL'
        summary_sheet[f'A{summary_row + 6}'] = 'OTHER TOTAL'
        summary_sheet[f'A{summary_row + 7}'] = 'PROJECT TOTAL'
        
        # Create SUMIF formulas to sum by sheet type for PRICES (column C)
        summary_sheet[f'B{summary_row + 1}'] = f'=SUMIF(A:A,"CANOPY",C:C)'  # Sum all CANOPY sheet prices
        summary_sheet[f'B{summary_row + 2}'] = f'=SUMIF(A:A,"FIRE SUPP",C:C)'  # Sum all FIRE SUPP sheet prices
        summary_sheet[f'B{summary_row + 3}'] = f'=SUMIF(A:A,"EBOX",C:C)'  # Sum all EBOX sheet prices
        summary_sheet[f'B{summary_row + 4}'] = f'=SUMIF(A:A,"SDU",C:C)'  # Sum all SDU sheet prices
        summary_sheet[f'B{summary_row + 5}'] = f'=SUMIF(A:A,"RECOAIR",C:C)'  # Sum all RECOAIR sheet prices
        summary_sheet[f'B{summary_row + 6}'] = f'=SUMIF(A:A,"OTHER",C:C)'  # Sum all OTHER sheet prices
        summary_sheet[f'B{summary_row + 7}'] = f'=B{summary_row + 1}+B{summary_row + 2}+B{summary_row + 3}+B{summary_row + 4}+B{summary_row + 5}+B{summary_row + 6}'  # Project price total (including RecoAir)
        
        # Create SUMIF formulas to sum by sheet type for COSTS (column D)
        summary_sheet[f'C{summary_row + 1}'] = f'=SUMIF(A:A,"CANOPY",D:D)'  # Sum all CANOPY sheet costs
        summary_sheet[f'C{summary_row + 2}'] = f'=SUMIF(A:A,"FIRE SUPP",D:D)'  # Sum all FIRE SUPP sheet costs
        summary_sheet[f'C{summary_row + 3}'] = f'=SUMIF(A:A,"EBOX",D:D)'  # Sum all EBOX sheet costs
        summary_sheet[f'C{summary_row + 4}'] = f'=SUMIF(A:A,"SDU",D:D)'  # Sum all SDU sheet costs
        summary_sheet[f'C{summary_row + 5}'] = f'=SUMIF(A:A,"RECOAIR",D:D)'  # Sum all RECOAIR sheet costs
        summary_sheet[f'C{summary_row + 6}'] = f'=SUMIF(A:A,"OTHER",D:D)'  # Sum all OTHER sheet costs
        summary_sheet[f'C{summary_row + 7}'] = f'=C{summary_row + 1}+C{summary_row + 2}+C{summary_row + 3}+C{summary_row + 4}+C{summary_row + 5}+C{summary_row + 6}'  # Project cost total (including RecoAir)
        
        # Store the summary row positions for JOB TOTAL to reference
        summary_sheet['H1'] = 'Reference Cells for JOB TOTAL'
        summary_sheet['H2'] = f'CANOPY_PRICE_TOTAL=B{summary_row + 1}'
        summary_sheet['H3'] = f'FIRE_SUPP_PRICE_TOTAL=B{summary_row + 2}'
        summary_sheet['H4'] = f'EBOX_PRICE_TOTAL=B{summary_row + 3}'
        summary_sheet['H5'] = f'SDU_PRICE_TOTAL=B{summary_row + 4}'
        summary_sheet['H6'] = f'RECOAIR_PRICE_TOTAL=B{summary_row + 5}'
        summary_sheet['H7'] = f'OTHER_PRICE_TOTAL=B{summary_row + 6}'
        summary_sheet['H8'] = f'PROJECT_PRICE_TOTAL=B{summary_row + 7}'
        summary_sheet['H9'] = f'CANOPY_COST_TOTAL=C{summary_row + 1}'
        summary_sheet['H10'] = f'FIRE_SUPP_COST_TOTAL=C{summary_row + 2}'
        summary_sheet['H11'] = f'EBOX_COST_TOTAL=C{summary_row + 3}'
        summary_sheet['H12'] = f'SDU_COST_TOTAL=C{summary_row + 4}'
        summary_sheet['H13'] = f'RECOAIR_COST_TOTAL=C{summary_row + 5}'
        summary_sheet['H14'] = f'OTHER_COST_TOTAL=C{summary_row + 6}'
        summary_sheet['H15'] = f'PROJECT_COST_TOTAL=C{summary_row + 7}'
        
        print(f"Created PRICING_SUMMARY sheet with {current_row - 2} individual sheet references")
        
    except Exception as e:
        print(f"Warning: Could not create PRICING_SUMMARY sheet: {str(e)}")

def update_job_total_sheet(wb: Workbook) -> None:
    """
    Update the JOB TOTAL sheet to reference the PRICING_SUMMARY sheet for dynamic pricing.
    
    Args:
        wb (Workbook): The workbook containing the JOB TOTAL sheet
    """
    try:
        if 'JOB TOTAL' not in wb.sheetnames:
            print("Warning: JOB TOTAL sheet not found")
            return
        
        if 'PRICING_SUMMARY' not in wb.sheetnames:
            print("Warning: PRICING_SUMMARY sheet not found")
            return
        
        job_total_sheet = wb['JOB TOTAL']
        
        # Get the summary row positions from PRICING_SUMMARY sheet
        pricing_summary = wb['PRICING_SUMMARY']
        
        # Find the summary totals section
        summary_row = None
        for row in range(1, 100):  # Search first 100 rows
            cell_value = pricing_summary[f'A{row}'].value
            if cell_value and 'SUMMARY TOTALS' in str(cell_value):
                summary_row = row
                break
        
        if not summary_row:
            print("Warning: Could not find SUMMARY TOTALS section in PRICING_SUMMARY")
            return
        
        # Update JOB TOTAL sheet with formulas referencing PRICING_SUMMARY
        # Only populate rows 16-25 for columns S (cost) and T (price)
        
        # Write PRICE totals to column T:
        job_total_sheet['T16'] = f"=PRICING_SUMMARY!B{summary_row + 1}"  # Canopy Total (Price)
        job_total_sheet['T17'] = f"=PRICING_SUMMARY!B{summary_row + 2}"  # Fire Suppression Total (Price)
        job_total_sheet['T18'] = f"=PRICING_SUMMARY!B{summary_row + 4}"  # SDU Total (SDU category - Price)
        job_total_sheet['T19'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # Vent Clg Total (OTHER category - Price)
        job_total_sheet['T20'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # MARVEL Total (OTHER category - Price)
        job_total_sheet['T21'] = f"=PRICING_SUMMARY!B{summary_row + 3}"  # Edge Total (EBOX/UV-C - Price)
        job_total_sheet['T22'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # Aerolys Total (OTHER category - Price)
        job_total_sheet['T23'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # Pollustop Total (OTHER category - Price)
        job_total_sheet['T24'] = f"=PRICING_SUMMARY!B{summary_row + 5}"  # Reco Total (RecoAir - Price)
        job_total_sheet['T25'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # Reactaway Total (OTHER category - Price)
        
        # Write COST totals to column S:
        job_total_sheet['S16'] = f"=PRICING_SUMMARY!C{summary_row + 1}"  # Canopy Total (Cost)
        job_total_sheet['S17'] = f"=PRICING_SUMMARY!C{summary_row + 2}"  # Fire Suppression Total (Cost)
        job_total_sheet['S18'] = f"=PRICING_SUMMARY!C{summary_row + 4}"  # SDU Total (SDU category - Cost)
        job_total_sheet['S19'] = f"=PRICING_SUMMARY!C{summary_row + 6}"  # Vent Clg Total (OTHER category - Cost)
        job_total_sheet['S20'] = f"=PRICING_SUMMARY!C{summary_row + 6}"  # MARVEL Total (OTHER category - Cost)
        job_total_sheet['S21'] = f"=PRICING_SUMMARY!C{summary_row + 3}"  # Edge Total (EBOX/UV-C - Cost)
        job_total_sheet['S22'] = f"=PRICING_SUMMARY!C{summary_row + 6}"  # Aerolys Total (OTHER category - Cost)
        job_total_sheet['S23'] = f"=PRICING_SUMMARY!C{summary_row + 6}"  # Pollustop Total (OTHER category - Cost)
        job_total_sheet['S24'] = f"=PRICING_SUMMARY!C{summary_row + 5}"  # Reco Total (RecoAir - Cost)
        job_total_sheet['S25'] = f"=PRICING_SUMMARY!C{summary_row + 6}"  # Reactaway Total (OTHER category - Cost)
        
        print("Updated JOB TOTAL sheet with dynamic pricing formulas")
        
    except Exception as e:
        print(f"Warning: Could not update JOB TOTAL sheet: {str(e)}")

# Add this new function after the save_to_excel function

def create_revision_from_existing(excel_path: str, new_revision: str, new_date: str = None) -> str:
    """
    Create a new revision by copying an existing Excel file and updating only the revision and date.
    This preserves all existing data, formulas, and pricing.
    
    Args:
        excel_path (str): Path to the existing Excel file
        new_revision (str): New revision letter (e.g., "B", "C")
        new_date (str, optional): New date in DD/MM/YYYY format. If None, keeps existing date.
    
    Returns:
        str: Path to the new revision file
    """
    try:
        # Load the existing workbook (without data_only to preserve formulas)
        wb = load_workbook(excel_path)
        
        # Update revision in all sheets that have the revision field (O7)
        sheets_to_update = []
        
        # Check all visible sheets for revision field
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if sheet.sheet_state == 'visible':
                try:
                    # Check if O7 exists and has a value (indicating it's a sheet with revision field)
                    if sheet['O7'].value is not None:
                        sheets_to_update.append(sheet_name)
                except:
                    # Skip sheets that don't have O7 or can't access it
                    continue
        
        # Update revision in all identified sheets
        for sheet_name in sheets_to_update:
            sheet = wb[sheet_name]
            try:
                sheet['O7'] = new_revision
                print(f"Updated revision to {new_revision} in sheet: {sheet_name}")
            except Exception as e:
                print(f"Warning: Could not update revision in sheet {sheet_name}: {str(e)}")
        
        # Update date if provided
        if new_date:
            for sheet_name in sheets_to_update:
                sheet = wb[sheet_name]
                try:
                    # Update date in G7 (standard date field)
                    sheet['G7'] = new_date
                    print(f"Updated date to {new_date} in sheet: {sheet_name}")
                except Exception as e:
                    print(f"Warning: Could not update date in sheet {sheet_name}: {str(e)}")
        
        # Update revision in ProjectData sheet if it exists
        if 'ProjectData' in wb.sheetnames:
            try:
                hidden_sheet = wb['ProjectData']
                hidden_sheet['B7'] = new_revision  # Update revision in ProjectData
                if new_date:
                    # Add date to ProjectData if not already there
                    hidden_sheet['A8'] = 'Date'
                    hidden_sheet['B8'] = new_date
                print(f"Updated revision in ProjectData sheet to {new_revision}")
            except Exception as e:
                print(f"Warning: Could not update ProjectData sheet: {str(e)}")
        
        # Generate output filename
        # Extract project number from the original filename or from the data
        original_filename = os.path.basename(excel_path)
        
        # Try to extract project number from filename or from sheet data
        project_number = "unknown"
        try:
            # Try to get project number from JOB TOTAL sheet
            if 'JOB TOTAL' in wb.sheetnames:
                job_total_sheet = wb['JOB TOTAL']
                project_number = job_total_sheet['C3'].value or "unknown"
            elif sheets_to_update:
                # Get from first available sheet
                first_sheet = wb[sheets_to_update[0]]
                project_number = first_sheet['C3'].value or "unknown"
        except:
            # If we can't get project number from sheets, try to extract from filename
            if " Cost Sheet " in original_filename:
                project_number = original_filename.split(" Cost Sheet ")[0]
        
        # Format date for filename
        if new_date:
            formatted_date = new_date.replace('/', '')
        else:
            # Try to get existing date from sheet
            try:
                if 'JOB TOTAL' in wb.sheetnames:
                    existing_date = wb['JOB TOTAL']['G7'].value or ""
                elif sheets_to_update:
                    existing_date = wb[sheets_to_update[0]]['G7'].value or ""
                else:
                    existing_date = ""
                
                if existing_date:
                    formatted_date = str(existing_date).replace('/', '')
                else:
                    formatted_date = datetime.now().strftime("%d%m%Y")
            except:
                formatted_date = datetime.now().strftime("%d%m%Y")
        
        # Create output filename: "Project Number Cost Sheet Date"
        output_filename = f"{project_number} Cost Sheet {formatted_date}.xlsx"
        output_path = f"output/{output_filename}"
        
        # Ensure output directory exists
        os.makedirs("output", exist_ok=True)
        
        # Save the updated workbook
        wb.save(output_path)
        wb.close()
        
        print(f"Created revision {new_revision} at: {output_path}")
        return output_path
        
    except Exception as e:
        raise Exception(f"Failed to create revision from existing file: {str(e)}")

def extract_sdu_electrical_services(sheet: Worksheet) -> Dict:
    """
    Extract electrical and gas services data from SDU sheet.
    Reads from the electrical services section (B35-B68) and gas services section (C71-C82).
    
    Args:
        sheet (Worksheet): The SDU worksheet to read from
        
    Returns:
        Dict: Electrical and gas services data with mapped values
    """
    try:
        electrical_services = {
            'distribution_board': 0,
            'single_phase_switched_spur': 0,
            'three_phase_socket_outlet': 0,
            'switched_socket_outlet': 0,
            'emergency_knock_off': 0,
            'ring_main_inc_2no_sso': 0
        }
        
        # Check distribution board value at C35
        distribution_board_value = sheet['C35'].value
        if distribution_board_value and str(distribution_board_value).strip() not in ['', '0', '-']:
            try:
                electrical_services['distribution_board'] = int(float(str(distribution_board_value).strip()))
            except (ValueError, TypeError):
                electrical_services['distribution_board'] = 0
        
        # If distribution board has a value, check C40-C47 for single phase switched spur
        if electrical_services['distribution_board'] > 0:
            for row in range(40, 48):  # C40 to C47
                cell_value = sheet[f'C{row}'].value
                if cell_value and str(cell_value).strip() not in ['', '0', '-']:
                    try:
                        electrical_services['single_phase_switched_spur'] = int(float(str(cell_value).strip()))
                        break  # Take the first non-zero value found
                    except (ValueError, TypeError):
                        continue
        else:
            # If no distribution board value, check C49-C56 for three phase socket outlet
            for row in range(49, 57):  # C49 to C56
                cell_value = sheet[f'C{row}'].value
                if cell_value and str(cell_value).strip() not in ['', '0', '-']:
                    try:
                        electrical_services['three_phase_socket_outlet'] = int(float(str(cell_value).strip()))
                        break  # Take the first non-zero value found
                    except (ValueError, TypeError):
                        continue
        
        # Switched socket outlet value at C65
        switched_socket_value = sheet['C65'].value
        if switched_socket_value and str(switched_socket_value).strip() not in ['', '0', '-']:
            try:
                electrical_services['switched_socket_outlet'] = int(float(str(switched_socket_value).strip()))
            except (ValueError, TypeError):
                electrical_services['switched_socket_outlet'] = 0
        
        # Emergency knock-off value (assuming it's around the electrical services section)
        # You may need to specify the exact cell for this
        emergency_value = sheet['C57'].value  # Adjust this cell reference as needed
        if emergency_value and str(emergency_value).strip() not in ['', '0', '-']:
            try:
                electrical_services['emergency_knock_off'] = int(float(str(emergency_value).strip()))
            except (ValueError, TypeError):
                electrical_services['emergency_knock_off'] = 0
        
        # Ring main inc. 2no SSO value at C68
        ring_main_value = sheet['C68'].value
        if ring_main_value and str(ring_main_value).strip() not in ['', '0', '-']:
            try:
                electrical_services['ring_main_inc_2no_sso'] = int(float(str(ring_main_value).strip()))
            except (ValueError, TypeError):
                electrical_services['ring_main_inc_2no_sso'] = 0
        
        # Gas services extraction
        gas_services = {
            'gas_manifold': 0,
            'gas_connection_15mm': 0,
            'gas_connection_20mm': 0,
            'gas_connection_25mm': 0,
            'gas_connection_32mm': 0,
            'gas_solenoid_valve': 0
        }
        
        # Gas manifold value from C71-C74 (take first non-zero value)
        for row in range(71, 75):  # C71 to C74
            cell_value = sheet[f'C{row}'].value
            if cell_value and str(cell_value).strip() not in ['', '0', '-']:
                try:
                    gas_services['gas_manifold'] = int(float(str(cell_value).strip()))
                    break  # Take the first non-zero value found
                except (ValueError, TypeError):
                    continue
        
        # Gas connections - specific cell locations from C75 to C78
        gas_connection_cells = {
            'gas_connection_15mm': 'C75',   # 15MM gas connection
            'gas_connection_20mm': 'C76',   # 20MM gas connection  
            'gas_connection_25mm': 'C77',   # 25MM gas connection
            'gas_connection_32mm': 'C78'    # 32MM gas connection
        }
        
        for service_name, cell_ref in gas_connection_cells.items():
            try:
                cell_value = sheet[cell_ref].value
                if cell_value and str(cell_value).strip() not in ['', '0', '-']:
                    gas_services[service_name] = int(float(str(cell_value).strip()))
            except (ValueError, TypeError, KeyError):
                gas_services[service_name] = 0
        
        # Gas solenoid valve from C79-C82 (take first non-zero value)
        for row in range(79, 83):  # C79 to C82
            cell_value = sheet[f'C{row}'].value
            if cell_value and str(cell_value).strip() not in ['', '0', '-']:
                try:
                    gas_services['gas_solenoid_valve'] = int(float(str(cell_value).strip()))
                    break  # Take the first non-zero value found
                except (ValueError, TypeError):
                    continue
        
        # Water services extraction
        water_services = {
            'cws_manifold_22mm': 0,
            'cws_manifold_15mm': 0,
            'hws_manifold': 0,
            'water_connection_15mm': 0,
            'water_connection_22mm': 0,
            'water_connection_28mm': 0
        }
        
        # Extract manifold values
        # 22mm CWS manifold at C86
        cws_22mm_manifold = sheet['C86'].value
        if cws_22mm_manifold and str(cws_22mm_manifold).strip() not in ['', '0', '-']:
            try:
                water_services['cws_manifold_22mm'] = int(float(str(cws_22mm_manifold).strip()))
            except (ValueError, TypeError):
                water_services['cws_manifold_22mm'] = 0
        
        # 15mm CWS manifold at C87
        cws_15mm_manifold = sheet['C87'].value
        if cws_15mm_manifold and str(cws_15mm_manifold).strip() not in ['', '0', '-']:
            try:
                water_services['cws_manifold_15mm'] = int(float(str(cws_15mm_manifold).strip()))
            except (ValueError, TypeError):
                water_services['cws_manifold_15mm'] = 0
        
        # HWS manifold at C88
        hws_manifold = sheet['C88'].value
        if hws_manifold and str(hws_manifold).strip() not in ['', '0', '-']:
            try:
                water_services['hws_manifold'] = int(float(str(hws_manifold).strip()))
            except (ValueError, TypeError):
                water_services['hws_manifold'] = 0
        
        # Extract water connection values from fixed cells
        # C89: 15mm connection
        connection_15mm_value = sheet['C89'].value
        if connection_15mm_value and str(connection_15mm_value).strip() not in ['', '0', '-']:
            try:
                water_services['water_connection_15mm'] = int(float(str(connection_15mm_value).strip()))
            except (ValueError, TypeError):
                water_services['water_connection_15mm'] = 0
        else:
            water_services['water_connection_15mm'] = 0
        
        # C90: 22mm connection
        connection_22mm_value = sheet['C90'].value
        if connection_22mm_value and str(connection_22mm_value).strip() not in ['', '0', '-']:
            try:
                water_services['water_connection_22mm'] = int(float(str(connection_22mm_value).strip()))
            except (ValueError, TypeError):
                water_services['water_connection_22mm'] = 0
        else:
            water_services['water_connection_22mm'] = 0
        
        # C91: 28mm connection
        connection_28mm_value = sheet['C91'].value
        if connection_28mm_value and str(connection_28mm_value).strip() not in ['', '0', '-']:
            try:
                water_services['water_connection_28mm'] = int(float(str(connection_28mm_value).strip()))
            except (ValueError, TypeError):
                water_services['water_connection_28mm'] = 0
        else:
            water_services['water_connection_28mm'] = 0
        
        # Extract pricing information
        pricing = {
            'carcass_only_price': 0,
            'electrical_mechanical_price': 0,
            'live_site_test_price': 0,
            'delivery_price': 0,
            'final_carcass_price': 0,
            'final_electrical_price': 0,
            'has_live_test': False
        }
        
        # Carcass only price at C105
        carcass_price = sheet['C105'].value
        if carcass_price and str(carcass_price).strip() not in ['', '0', '-']:
            try:
                pricing['carcass_only_price'] = float(str(carcass_price).strip())
            except (ValueError, TypeError):
                pricing['carcass_only_price'] = 0
        
        # Electrical & mechanical services price at C106
        electrical_price = sheet['C106'].value
        if electrical_price and str(electrical_price).strip() not in ['', '0', '-']:
            try:
                pricing['electrical_mechanical_price'] = float(str(electrical_price).strip())
            except (ValueError, TypeError):
                pricing['electrical_mechanical_price'] = 0
        
        # Check for live site test at C102 and cost at J102
        live_test_quantity = sheet['C102'].value
        if live_test_quantity and str(live_test_quantity).strip() not in ['', '0', '-']:
            try:
                test_qty = float(str(live_test_quantity).strip())
                if test_qty > 0:
                    pricing['has_live_test'] = True
                    # Get cost from J102
                    live_test_cost = sheet['J102'].value
                    if live_test_cost and str(live_test_cost).strip() not in ['', '0', '-']:
                        try:
                            pricing['live_site_test_price'] = float(str(live_test_cost).strip())
                        except (ValueError, TypeError):
                            pricing['live_site_test_price'] = 0
            except (ValueError, TypeError):
                pass
        
        # Delivery price at C107 (to be divided by 2)
        delivery_price = sheet['C107'].value
        if delivery_price and str(delivery_price).strip() not in ['', '0', '-']:
            try:
                pricing['delivery_price'] = float(str(delivery_price).strip())
            except (ValueError, TypeError):
                pricing['delivery_price'] = 0
        
        # Calculate final prices
        # Delivery price divided by 2, half added to each
        half_delivery = pricing['delivery_price'] / 2
        pricing['final_carcass_price'] = pricing['carcass_only_price'] + half_delivery
        pricing['final_electrical_price'] = pricing['electrical_mechanical_price'] + half_delivery
        
        # Combine electrical, gas, water services, and pricing
        result = {
            'electrical_services': electrical_services,
            'gas_services': gas_services,
            'water_services': water_services,
            'pricing': pricing
        }
        
        return result
        
    except Exception as e:
        print(f"Warning: Could not extract electrical, gas, and water services data from SDU sheet: {str(e)}")
        return {
            'electrical_services': {
                'distribution_board': 0,
                'single_phase_switched_spur': 0,
                'three_phase_socket_outlet': 0,
                'switched_socket_outlet': 0,
                'emergency_knock_off': 0,
                'ring_main_inc_2no_sso': 0
            },
            'gas_services': {
                'gas_manifold': 0,
                'gas_connection_15mm': 0,
                'gas_connection_20mm': 0,
                'gas_connection_25mm': 0,
                'gas_connection_32mm': 0,
                'gas_solenoid_valve': 0
            },
            'water_services': {
                'cws_manifold_22mm': 0,
                'cws_manifold_15mm': 0,
                'hws_manifold': 0,
                'water_connection_15mm': 0,
                'water_connection_22mm': 0,
                'water_connection_28mm': 0
            },
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