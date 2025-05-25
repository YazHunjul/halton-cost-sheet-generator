"""
Excel generation utilities for HVAC quotation system.
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

# Cell mappings for different data points
CELL_MAPPINGS = {
    "project_number": "C3",  # Job No
    "customer": "C5",        # Customer
    "estimator": "C7",       # Sales Manager / Estimator Initials
    "project_name": "G3",    # Project Name
    "location": "G5",        # Location
    "date": "G7",           # Date
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

def write_project_metadata(sheet: Worksheet, project_data: Dict):
    """
    Write project metadata to the specified cells in the sheet.
    
    Args:
        sheet (Worksheet): The worksheet to write to
        project_data (Dict): Project metadata
    """
    for field, cell in CELL_MAPPINGS.items():
        value = project_data.get(field)
        
        if value:
            # Special handling for estimator/sales manager initials (only for sheet display)
            if field == "estimator":
                value = get_initials(value)  # Convert to initials for sheet display
            # Title case for other fields except date
            elif field != "date":
                value = str(value).title()
            # Date handling
            elif field == "date" and not value:
                value = datetime.now().strftime("%d/%m/%Y")
            
            sheet[cell] = value

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
            sheet[f"B{ref_row}"] = ref_number.upper()
        
        # Configuration and Model on same row
        configuration = canopy.get("configuration", "")
        if configuration:
            sheet[f"C{row_index}"] = configuration.upper()
        
        # Model in D14, D31, D48, etc.
        model = canopy.get("model", "")
        if model:
            sheet[f"D{row_index}"] = model.upper()
        
        # Options
        options_row = row_index + 4
        options = canopy.get("options", {})
        if options.get("fire_suppression"):
            sheet[f"B{options_row}"] = "FIRE SUPPRESSION SYSTEM"
        if options.get("uvc"):
            sheet[f"B{options_row + 1}"] = "UV-C SYSTEM"
        if options.get("sdu"):
            sheet[f"B{options_row + 2}"] = "SDU"
        if options.get("recoair"):
            sheet[f"B{options_row + 3}"] = "RECOAIR"
    except Exception as e:
        raise Exception(f"Failed to write canopy data: {str(e)}")

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
        
        # Add validations to sheet
        sheet.add_data_validation(lighting_dv)
        sheet.add_data_validation(special_works_dv)
        sheet.add_data_validation(cladding_dv)
        sheet.add_data_validation(wall_cladding_dv)
        
        # Apply dropdowns to multiple canopy sections (every 17 rows)
        for canopy_index in range(5):  # Support up to 5 canopies per sheet
            base_row = CANOPY_START_ROW + (canopy_index * CANOPY_ROW_SPACING)  # 14, 31, 48, 65, 82
            
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
        
        for canopy_index in range(5):
            try:
                base_row = CANOPY_START_ROW + (canopy_index * CANOPY_ROW_SPACING)
                config_dv.add(f"C{base_row}")  # Configuration in column C of the model row
                model_dv.add(f"D{base_row}")   # Model in column D of the model row (D14, D31, D48, etc.)
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

def organize_sheets_by_area(wb: Workbook):
    """
    Organize sheets so that related sheets (CANOPY, FIRE SUPP, etc.) for the same area are grouped together.
    
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
        
        # Sort area identifiers to maintain consistent ordering
        for area_identifier in sorted(sheet_groups.keys()):
            # Sort sheets within each area: CANOPY first, then FIRE SUPP, then others
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
        
        # Add other sheets at the end (JOB TOTAL, Lists, etc.)
        # Put JOB TOTAL at the end
        job_total_sheets = [s for s in other_sheets if 'JOB TOTAL' in s]
        other_important_sheets = [s for s in other_sheets if s not in job_total_sheets and s != 'Lists']
        lists_sheets = [s for s in other_sheets if s == 'Lists']
        
        ordered_sheets.extend(other_important_sheets)
        ordered_sheets.extend(job_total_sheets)
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
        sheet_count = 0
        fs_sheet_count = 0
        
        # Keep track of total areas for coloring
        area_count = 0
        
        # Process each level and area
        for level in project_data.get("levels", []):
            level_number = level["level_number"]
            level_name = level.get("level_name", f"Level {level_number}")
            
            for idx, area in enumerate(level["areas"], 1):
                area_name = area["name"]
                area_canopies = area.get("canopies", [])
                
                # Get tab color for this area
                tab_color = TAB_COLORS[area_count % len(TAB_COLORS)]
                
                # Check if area has fire suppression
                has_fire_suppression = any(canopy.get("options", {}).get("fire_suppression", False) for canopy in area_canopies)
                
                current_canopy_sheet = None
                fs_sheet = None
                
                # Process canopy sheet if canopies exist for this area
                if area_canopies:
                    if canopy_sheets:
                        sheet_name = canopy_sheets.pop(0)
                        current_canopy_sheet = wb[sheet_name]
                        
                        # Set title in B1
                        sheet_title_display = f"{level_name} - {area_name}"
                        current_canopy_sheet['B1'] = sheet_title_display
                        
                        # Rename the sheet tab
                        canopy_sheet_tab_name = f"CANOPY - {level_name} ({sheet_count + 1})"
                        current_canopy_sheet.title = canopy_sheet_tab_name
                        current_canopy_sheet.sheet_state = 'visible'
                        current_canopy_sheet.sheet_properties.tabColor = tab_color
                        
                        # Write project metadata to canopy sheet
                        write_project_metadata(current_canopy_sheet, project_data)
                        
                        # Create fire suppression sheet if needed
                        if has_fire_suppression:
                            if fire_supp_sheets:
                                fs_sheet_name = fire_supp_sheets.pop(0)
                                fs_sheet = wb[fs_sheet_name]
                                new_fs_name = f"FIRE SUPP - {level_name} ({sheet_count + 1})"
                                fs_sheet.title = new_fs_name
                                fs_sheet.sheet_state = 'visible'
                                fs_sheet.sheet_properties.tabColor = tab_color
                                
                                # Write project metadata to fire suppression sheet
                                write_project_metadata(fs_sheet, project_data)
                                # Set fire suppression sheet title in B1
                                fs_sheet['B1'] = f"{level_name} - {area_name} - FIRE SUPPRESSION"
                                
                                fs_sheet_count += 1
                        
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
                        
                        sheet_count += 1
                    else:
                        raise Exception(f"Not enough CANOPY sheets in template for area {area_name}")
                
                area_count += 1
        
        # Write project metadata to any other visible sheets that might exist
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if (sheet.sheet_state == 'visible' and 
                not sheet_name.startswith(('CANOPY', 'FIRE SUPP')) and 
                sheet_name not in ['Lists', 'JOB TOTAL']):
                # Write metadata to any other visible sheets
                try:
                    write_project_metadata(sheet, project_data)
                except Exception as e:
                    print(f"Warning: Could not write metadata to sheet {sheet_name}: {str(e)}")
        
        # Organize sheets by area for better navigation
        organize_sheets_by_area(wb)
        
        # Collect wall cladding data from all canopies
        wall_cladding_data = collect_wall_cladding_data(project_data)
        
        # Write wall cladding summary to all visible canopy sheets
        if wall_cladding_data:
            for sheet_name in wb.sheetnames:
                if sheet_name.startswith('CANOPY') and wb[sheet_name].sheet_state == 'visible':
                    write_wall_cladding_summary(wb[sheet_name], wall_cladding_data)
        
        # Write delivery location to D183 for all sheets except JOB TOTAL
        delivery_location = project_data.get('location', '')
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            if (sheet.sheet_state == 'visible' and 
                sheet_name != 'JOB TOTAL' and 
                sheet_name != 'Lists'):
                write_delivery_location_to_sheet(sheet, delivery_location)
        
        # Remove any unused template sheets
        unused_sheets = canopy_sheets + fire_supp_sheets + edge_box_sheets + recoair_sheets
        for sheet_name in unused_sheets:
            if sheet_name in wb.sheetnames:
                wb.remove(wb[sheet_name])
        
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
        project_data['location'] = data_sheet['G5'].value or ""
        project_data['date'] = data_sheet['G7'].value or ""
        
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
        
        for sheet_name in wb.sheetnames:
            if 'CANOPY - ' in sheet_name or 'FIRE SUPP - ' in sheet_name:
                sheet = wb[sheet_name]
                title_cell = sheet['B1'].value
                
                if title_cell and ' - ' in title_cell:
                    level_area = title_cell.split(' - ')
                    if len(level_area) >= 2:
                        level_name = level_area[0]
                        area_name = level_area[1]
                        
                        if level_name not in levels_data:
                            levels_data[level_name] = []
                        
                        if area_name not in [area['name'] for area in levels_data[level_name]]:
                            levels_data[level_name].append({
                                'name': area_name,
                                'canopies': []
                            })
                        
                        # If this is a canopy sheet, try to read canopy data
                        if 'CANOPY - ' in sheet_name:
                            # Read canopy specifications from the sheet
                            # This is a simplified read - you might want to enhance this
                            for canopy_idx in range(5):  # Support up to 5 canopies
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
                                        
                                        # Fire suppression data - will be populated from FIRE SUPP sheet
                                        'fire_suppression_tank_quantity': 0,  # Default to 0, will be updated from FIRE SUPP sheet
                                        
                                        # Initialize wall cladding as None - will be populated later if found
                                        'wall_cladding': {
                                            'type': 'None',
                                            'width': None,
                                            'height': None,
                                            'position': None
                                        }
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
                        
                        # Read fire suppression tank quantities for canopies in this area
                        for canopy_idx in range(5):  # Support up to 5 canopies
                            base_row = CANOPY_START_ROW + (canopy_idx * CANOPY_ROW_SPACING)
                            ref_row = base_row - 2
                            tank_row = base_row + 3  # C17 relative to base_row (14 + 3 = 17)
                            
                            ref_number = sheet[f'B{ref_row}'].value
                            tank_value = sheet[f'C{tank_row}'].value
                            
                            if ref_number and tank_value:
                                tank_quantity = extract_tank_quantity(tank_value)
                                
                                # Find the corresponding canopy and update its tank quantity
                                for level in levels_data.get(level_name, []):
                                    if level['name'] == area_name:
                                        for canopy in level['canopies']:
                                            if canopy['reference_number'] == ref_number:
                                                canopy['fire_suppression_tank_quantity'] = tank_quantity
                                                break
        
        # Convert levels_data to the format expected by the system
        project_data['levels'] = []
        for level_name, areas in levels_data.items():
            project_data['levels'].append({
                'level_name': level_name,
                'areas': areas
            })
        
        # Read wall cladding data from consolidated section (starting at row 19)
        # We'll read from any CANOPY sheet and apply the cladding data to appropriate canopies
        for sheet_name in wb.sheetnames:
            if 'CANOPY - ' in sheet_name:
                sheet = wb[sheet_name]
                
                # Check if there's wall cladding data (indicated by "2M² (HFL)" in C19)
                if sheet['C19'].value and "2M²" in str(sheet['C19'].value):
                    # Read wall cladding entries starting from row 19
                    for row_idx in range(19, 25):  # Check rows 19-24 for wall cladding entries
                        dimensions_cell = sheet[f'P{row_idx}'].value
                        position_cell = sheet[f'Q{row_idx}'].value
                        
                        if dimensions_cell and position_cell:
                            # Parse dimensions (e.g., "1000X2100" -> width=1000, height=2100)
                            try:
                                if 'X' in str(dimensions_cell):
                                    width_str, height_str = str(dimensions_cell).split('X')
                                    width = int(width_str.strip())
                                    height = int(height_str.strip())
                                else:
                                    continue  # Skip if format is not width X height
                            except (ValueError, AttributeError):
                                continue  # Skip if parsing fails
                            
                            # Convert position to list format (e.g., "rear/left hand" -> ["rear", "left hand"])
                            position_list = [pos.strip() for pos in str(position_cell).split('/')]
                            
                            # Create wall cladding data
                            wall_cladding_data = {
                                'type': 'Custom',
                                'width': width,
                                'height': height,
                                'position': position_list
                            }
                            
                            # For now, we'll assign this wall cladding to the first canopy in the first area
                            # In a more sophisticated approach, you might want to track which canopy
                            # the cladding belongs to based on reference numbers or other identifiers
                            if project_data['levels']:
                                first_level = project_data['levels'][0]
                                if first_level.get('areas'):
                                    first_area = first_level['areas'][0]
                                    if first_area.get('canopies'):
                                        # Find a canopy without wall cladding assigned yet
                                        for canopy in first_area['canopies']:
                                            if canopy['wall_cladding']['type'] == 'None':
                                                canopy['wall_cladding'] = wall_cladding_data
                                                break
                
                # Only process wall cladding from the first CANOPY sheet found
                break
        
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
                    
                    # Join positions for other uses
                    position_str = "/".join(position_list) if position_list else ""
                    
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