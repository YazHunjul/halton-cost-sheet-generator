# HVAC Project Management Tool

A comprehensive Streamlit application for managing HVAC canopy projects, generating Excel cost sheets, and creating Word quotation documents.

## Features

### 🏗️ Project Management

- **Project Type Support**: Canopy Projects and RecoAir Projects
- **Multi-level Structure**: Support for multiple levels and areas
- **Canopy Configuration**: Wall, Island, Single, Double, Corner configurations
- **Options Management**: Fire suppression, UV-C systems, SDU, RecoAir

### 📊 Excel Generation

- **Automated Cost Sheets**: Generate detailed Excel workbooks from project data
- **Template-based**: Uses professional Excel templates with proper formatting
- **Multiple Sheet Types**: CANOPY, FIRE SUPP, EDGE BOX, RECOAIR sheets
- **Data Validation**: Dropdown menus for consistent data entry
- **Business Logic**: Automatic calculations and formatting rules

### 📄 Word Document Generation

- **Professional Quotations**: Generate Word documents from Excel data
- **Jinja2 Templates**: Flexible template system for customization
- **Business Rules**: Automatic formatting and data transformation
- **Fire Suppression Detection**: Automatically includes fire suppression sections when applicable

## Project Structure

```
UKCS/
├── src/
│   ├── components/          # Streamlit form components
│   │   ├── forms.py        # General project forms
│   │   └── project_forms.py # Project structure forms
│   ├── config/             # Configuration files
│   │   ├── business_data.py # Business rules and data
│   │   └── constants.py    # Application constants
│   ├── utils/              # Utility modules
│   │   ├── excel.py        # Excel generation and reading
│   │   └── word.py         # Word document generation
│   └── app.py              # Main Streamlit application
├── templates/
│   ├── excel/              # Excel templates
│   └── word/               # Word templates
├── output/                 # Generated files (gitignored)
└── debug_*.py             # Debug and test scripts
```

## Installation

1. **Clone the repository**:

   ```bash
   git clone <repository-url>
   cd UKCS
   ```

2. **Install dependencies**:

   ```bash
   pip install streamlit openpyxl python-docx-template
   ```

3. **Run the application**:
   ```bash
   streamlit run src/app.py
   ```

## Usage

### Creating a New Project

1. **Select Project Type**: Choose between "Canopy Project" or "RecoAir Project"
2. **Fill Project Information**: Enter customer details, project name, location, etc.
3. **Define Project Structure**: Add levels, areas, and canopies
4. **Configure Canopies**: Set models, configurations, and options
5. **Generate Cost Sheet**: Export to Excel format
6. **Generate Quotation**: Create Word document from Excel data

### Excel Cost Sheet Features

- **Automatic Sheet Creation**: Creates separate sheets for each area
- **Fire Suppression Sheets**: Automatically generated when fire suppression is enabled
- **Data Validation**: Dropdown menus for models, configurations, lighting types
- **Business Logic**:
  - KVI models (without 'F') get "-" for MUA volume and supply static
  - CMWF/CMWI models get "-" for extract static
  - Automatic formatting of volumes and static pressures

### Word Document Features

- **Template-based Generation**: Uses Jinja2 templates for flexibility
- **Data Transformation**:
  - Empty values become "-"
  - "LED STRIP L12 inc DALI" → "LED STRIP"
  - "LIGHT SELECTION" → "-"
  - Remove "Pa" from static pressure values
  - Round MUA volumes to 1 decimal place
- **Fire Suppression Detection**: Automatically includes fire suppression sections when sheets exist
- **Wall Cladding Support**: Organized by area with proper descriptions

## Business Rules

### Canopy Models

- **KVI Models**: No 'F' in name → MUA volume and supply static = "-"
- **CMWF/CMWI Models**: Extract static = "-"
- **Models with 'F'**: Full volume and static data

### Fire Suppression

- **Sheet Detection**: Fire suppression section appears if FIRE SUPP sheets exist
- **Tank Quantities**: Shows actual numbers or "TBD" if not specified
- **Area-based**: Organized by level and area

### Formatting

- **Extract Static**: Removes "Pa" units (e.g., "150 Pa" → "150")
- **MUA Volume**: Rounded to 1 decimal place (e.g., "2.345" → "2.3")
- **Lighting Types**: Standardized to "LED STRIP", "LED SPOTS", or "-"

## Recent Updates

### Fire Suppression Enhancement

- ✅ Fire suppression items now appear when FIRE SUPP sheets exist, even if tank quantities are 0
- ✅ Shows "TBD" for empty tank quantities instead of hiding the section
- ✅ Area-based detection for better organization

### Formatting Improvements

- ✅ Remove "Pa" from extract static values
- ✅ Round MUA volume to 1 decimal place
- ✅ Enhanced lighting type transformation
- ✅ Consistent empty value handling

### Excel Reading Enhancements

- ✅ Improved fire suppression data extraction from FIRE SUPP sheets
- ✅ Better handling of placeholder rows and empty values
- ✅ Enhanced wall cladding data processing

## Debug Tools

The project includes several debug scripts for troubleshooting:

- `debug_fire_suppression_data.py`: Check fire suppression data in Excel files
- `debug_word_template.py`: Debug Word template processing
- Various test scripts for specific functionality

## Templates

### Excel Templates

- Located in `templates/excel/`
- Professional formatting with data validation
- Multiple sheet types for different project needs

### Word Templates

- Located in `templates/word/`
- Jinja2-based templates for flexible document generation
- Professional quotation format

## Contributing

1. Follow the existing code structure
2. Add appropriate error handling
3. Update documentation for new features
4. Test with various project configurations

## License

[Add your license information here]

## Support

For issues or questions, please [add contact information or issue tracker link].
