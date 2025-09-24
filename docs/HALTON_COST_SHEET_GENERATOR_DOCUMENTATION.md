# Halton Cost Sheet Generator - Comprehensive Documentation

## Overview

The **Halton Cost Sheet Generator** is a comprehensive Streamlit web application designed for managing Halton canopy projects, generating professional Excel cost sheets, and creating Word quotation documents. The application supports multi-level projects with complex canopy configurations, fire suppression systems, and specialized Halton air handling (RecoAir) systems.

### Key Features
- ✅ Multi-level project management with hierarchical structure
- ✅ Professional Excel cost sheet generation with multiple templates
- ✅ Automated Word quotation document creation
- ✅ Fire suppression system integration
- ✅ RecoAir system support with specialized templates
- ✅ Wall cladding configuration and tracking
- ✅ UV-C system integration
- ✅ Project data import/export functionality
- ✅ URL-based state sharing
- ✅ Advanced document preview capabilities

## Architecture & Project Structure

### Application Entry Points
```
📁 UKCS/
├── app.py                    # Main entry point (imports from src/app.py)
└── 📁 src/
    ├── app.py               # Core Streamlit application
    ├── 📁 components/       # Reusable UI components
    ├── 📁 config/          # Configuration and business data
    ├── 📁 utils/           # Core business logic utilities
    └── 📁 supabase/        # Database integration (optional)
```

### Core Architecture Components

#### 1. **Main Application** (`src/app.py`)
- **Purpose**: Primary Streamlit orchestrator and UI controller
- **Key Functions**:
  - `main()`: Application entry point and navigation
  - `display_project_summary()`: Project data visualization
  - `word_generation_page()`: Document generation interface
  - `project_creation_page()`: Project data input interface

#### 2. **Business Logic Layer** (`src/utils/`)
- **Excel Generation** (`excel.py`): Template-based Excel workbook creation
- **Word Generation** (`word.py`): Jinja2-templated quotation documents
- **State Management** (`state_manager.py`): URL serialization and data persistence
- **Date Utilities** (`date_utils.py`): Consistent date formatting and handling
- **Word Preview** (`word_preview.py`): Document preview with pandoc integration

#### 3. **Configuration Layer** (`src/config/`)
- **Constants** (`constants.py`): Feature flags and system configuration
- **Business Data** (`business_data.py`): Sales contacts, estimators, addresses, valid models

#### 4. **Component Layer** (`src/components/`)
- **Project Forms** (`project_forms.py`): Reusable form components for project data
- **Forms** (`forms.py`): Additional UI components and validators

## Application Functionality & Features

### 1. Project Creation Interface

#### Project Structure
The application supports a hierarchical project structure:
```
Project
├── General Information (Name, Number, Customer, etc.)
├── Level 1
│   ├── Area A
│   │   ├── Canopy 1 (Model, Configuration, Options)
│   │   ├── Canopy 2
│   │   └── Area Options (UV-C, RecoAir, Marvel)
│   └── Area B
└── Level 2 (Additional levels as needed)
```

#### Input Forms & Validation
- **Project Metadata**: Name, number, customer, location, sales contact, estimator
- **Canopy Configuration**:
  - Reference numbers with validation
  - Model selection from `VALID_CANOPY_MODELS`
  - Configuration types: Wall, Island, Single, Double, Corner
  - Fire suppression options
  - Wall cladding specifications (type, dimensions, positioning)
- **Area-Level Options**: UV-C systems, RecoAir integration, Marvel systems
- **Data Validation**: Required fields, model validation, reference uniqueness

### 2. Excel Cost Sheet Generation

#### Template System
- **Template Versions**: R19.1 (May 2025) and R19.2 (June 2025)
- **Default Template**: R19.2 (latest version)
- **Template Location**: `templates/excel/`

#### Sheet Types & Generation
1. **CANOPY Sheets**: One per area with canopy details
2. **FIRE SUPP Sheets**: Fire suppression system specifications
3. **EDGE BOX Sheets**: Edge box configurations
4. **RECOAIR Sheets**: RecoAir system specifications
5. **Lists Sheet**: Supporting data and validation lists

#### Business Logic Implementation
- **KVI Models**: Without 'F' → MUA volume and supply static = "-"
- **CMWF/CMWI Models**: Extract static = "-"
- **Level-Based Color Coding**: Each level gets unique tab colors
- **Cell Mappings**: Project metadata populated at specific cells:
  - C3: Project Number, C5: Company, C7: Estimator
  - G3: Project Name, G5: Location, G7: Date
  - K7: Revision information

#### Advanced Features
- **External Link Removal**: Prevents "unsafe external sources" warnings
- **Dynamic Sheet Creation**: Creates sheets based on project structure
- **Data Validation**: Dropdown lists and input validation
- **Formatting**: Professional styling with conditional formatting

### 3. Word Document Generation

#### Template System
- **Main Quotation**: `templates/word/Halton Quote Feb 2024.docx`
- **RecoAir Quotation**: `templates/word/Halton RECO Quotation Jan 2025 (2).docx`
- **Template Engine**: Jinja2 for dynamic content generation

#### Document Types & Logic
1. **Standard Canopy Quotations**: For projects with canopy systems
2. **RecoAir-Only Quotations**: Specialized template for RecoAir projects
3. **Mixed Projects**: Creates ZIP file with multiple documents

#### Data Transformation Features
- **Fire Suppression Detection**: Automatic detection from FIRE SUPP sheets
- **System Type Mapping**: NOBEL → NOBEL System, AMAREX → AMAREX System, default → Ansul R102
- **Sales Contact Integration**: Automatic phone number inclusion
- **Empty Value Handling**: Professional formatting for missing data
- **Lighting Type Normalization**: Consistent lighting descriptions
- **Static Pressure Formatting**: Proper engineering units

#### Document Analysis
- **Project Type Detection**: Analyzes areas to determine quotation type
- **Multi-Document Logic**: Creates separate documents for different system types
- **File Naming**: Standardized format: `[Project Number] [Type] [Date].ext`

### 4. Advanced Features

#### State Management
- **Session Persistence**: Maintains form data across page refreshes
- **URL Serialization**: Share projects via encoded URLs
- **Data Export**: Save project data to Excel for later use
- **Import Functionality**: Load existing Excel files for modification

#### Document Preview
- **Preview Requirements**: Automatic detection of pandoc/pypandoc availability
- **Enhanced Preview**: Advanced formatting with pypandoc integration
- **Basic Preview**: Fallback preview without external dependencies
- **Table Preservation**: Maintains table formatting in previews
- **Document Statistics**: Paragraph count, table count, file size metrics

#### Feature Flag System
The application uses feature flags to control system visibility:

**Enabled Systems**:
- ✅ Canopy Systems (Core functionality)
- ✅ RecoAir Systems
- ✅ Fire Suppression Systems
- ✅ UV-C Systems
- ✅ SDU (Supply Distribution Units)
- ✅ Wall Cladding
- ✅ Cyclocell Cassette Ceiling

**Hidden/Development Systems**:
- 🚧 Kitchen Extract Systems
- 🚧 M.A.R.V.E.L. System (DCKV)
- 🚧 Reactaway Units
- 🚧 Dishwasher Extract
- 🚧 Gas Interlocking
- 🚧 Pollustop Units

## Business Rules & Logic

### Fire Suppression System Rules
- **Automatic Detection**: Based on presence of FIRE SUPP sheets in Excel
- **System Type Mapping**:
  - `NOBEL` → "NOBEL System. Supplied, installed & commissioned."
  - `AMAREX` → "AMAREX System. Supplied, installed & commissioned."
  - Default → "Ansul R102 system. Supplied, installed & commissioned."
- **Tank Quantities**: Display as "TBD" when values are empty
- **Integration**: Seamless integration with both Excel and Word generation

### Canopy Model Validation
- **Valid Models**: Maintained in `VALID_CANOPY_MODELS` list in `business_data.py`
- **Configuration Types**: Wall, Island, Single, Double, Corner
- **Model-Specific Logic**: KVI models have special volume/static pressure rules
- **Wall Cladding**: Dimensions and positioning tracked per canopy

### Date Handling Standards
- **Internal Format**: datetime objects or DD/MM/YYYY strings
- **Display Format**: DD/MM/YYYY format consistently
- **File Naming Format**: DDMMYYYY (no separators for file names)
- **Revision Integration**: Date-based revision tracking

### Sales & Business Data
- **Sales Contacts**: 14+ contacts with phone numbers
- **Estimators**: 3 estimators with role designations
- **Company Addresses**: 100+ customer addresses
- **Delivery Locations**: Comprehensive location database

## Technical Implementation

### Dependencies & Technology Stack
- **Frontend**: Streamlit (Web UI framework)
- **Excel Processing**: openpyxl (Excel file manipulation)
- **Word Processing**: python-docx, python-docxtpl (Jinja2 templating)
- **Preview System**: pypandoc (optional, for enhanced previews)
- **Data Processing**: pandas (data manipulation)
- **Date Handling**: datetime, dateutil
- **File Management**: zipfile, tempfile, os

### File Structure & Templates
```
📁 templates/
├── 📁 excel/
│   ├── Cost Sheet R19.1 May 2025.xlsx
│   ├── Cost Sheet R19.2 Jun 2025.xlsx
│   └── Halton Cost Sheet Jan 2025.xlsx
└── 📁 word/
    ├── Halton Quote Feb 2024.docx
    ├── Halton RECO Quotation Jan 2025 (2).docx
    └── Halton Quote Feb 2024_TEST.docx
```

### Output Management
- **Output Directory**: `output/` (gitignored)
- **Temporary Files**: Automatic cleanup of temporary files
- **File Naming**: Consistent naming convention across all outputs
- **ZIP Creation**: Automatic ZIP creation for multi-document projects

## Development & Deployment

### Running the Application
```bash
# Local development
streamlit run app.py

# Python version: 3.13.5
python3 -m pip install -r requirements.txt
```

### Testing Framework
The project uses manual test scripts rather than formal testing frameworks:
```bash
# Test specific functionality
python3 test_sdu_collection.py
python3 test_fire_suppression.py
python3 test_recoair_fix.py

# Debug Excel generation
python3 debug_excel_fire_suppression.py
python3 debug_sheet_values.py
python3 check_excel_cladding.py

# Debug Word generation
python3 debug_word_template.py
python3 debug_word_comprehensive.py
```

### Common Development Tasks

#### Adding New Canopy Models
1. Update `VALID_CANOPY_MODELS` in `src/config/business_data.py`
2. Ensure proper categorization (KVI vs others)
3. Test with various project configurations

#### Modifying Excel Templates
1. Update template files in `templates/excel/`
2. Adjust cell mappings in `src/utils/excel.py` if needed
3. Test with different project structures
4. Verify external link removal

#### Updating Word Templates
1. Modify Jinja2 templates in `templates/word/`
2. Update data transformation logic in `src/utils/word.py`
3. Test fire suppression detection
4. Validate multi-area scenarios

## User Interface & Navigation

### Main Navigation
The application provides a clean, intuitive interface with clear navigation:

1. **Project Creation Tab**: Complete project setup and configuration
2. **Word Generation Tab**: Upload existing Excel files and generate quotations
3. **Preview & Download**: Real-time document preview and download options

### Form Organization
- **Progressive Disclosure**: Complex forms broken into manageable sections
- **Validation Feedback**: Real-time validation with helpful error messages
- **State Preservation**: Form data maintained across navigation
- **Quick Actions**: Save progress and share project URLs

### Download & Export
- **Instant Download**: One-click download buttons for generated documents
- **Format Options**: Excel cost sheets and Word quotations
- **ZIP Files**: Automatic ZIP creation for multi-document projects
- **Preview Integration**: See before you download

## Error Handling & Validation

### Input Validation
- **Required Fields**: Clear marking and validation of mandatory fields
- **Model Validation**: Canopy models validated against approved list
- **Reference Uniqueness**: Duplicate reference number detection
- **Date Format**: Consistent date format validation and conversion

### Error Recovery
- **Graceful Degradation**: Application continues working with partial data
- **Clear Error Messages**: User-friendly error descriptions
- **Debug Information**: Detailed error logging for development
- **Fallback Options**: Alternative approaches when features fail

### File Handling
- **Template Validation**: Verification of template file integrity
- **External Link Removal**: Automatic cleanup to prevent security warnings
- **Temporary File Management**: Proper cleanup of temporary files
- **Path Handling**: Robust file path management across operating systems

## Future Enhancements & Roadmap

### Planned Features (Currently Feature-Flagged)
- **Kitchen Extract Systems**: Comprehensive kitchen ventilation
- **M.A.R.V.E.L. System**: Advanced DCKV control integration
- **Advanced Control Systems**: Expanded automation options
- **Additional Equipment**: Reactaway units, dishwasher extract, gas interlocking

### Technical Improvements
- **Database Integration**: Full Supabase integration for project persistence
- **User Authentication**: Multi-user support with role-based access
- **API Development**: RESTful API for external integrations
- **Enhanced Testing**: Comprehensive test suite development
- **Performance Optimization**: Caching and performance improvements

### User Experience Enhancements
- **Mobile Responsiveness**: Improved mobile interface
- **Bulk Operations**: Batch processing capabilities
- **Advanced Reporting**: Enhanced reporting and analytics
- **Template Management**: User-configurable templates
- **Workflow Automation**: Automated project workflows

---

*This documentation covers the comprehensive functionality of the Halton Cost Sheet Generator as of September 2024. The application continues to evolve with new features and improvements being added regularly.*