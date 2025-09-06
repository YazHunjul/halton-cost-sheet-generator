# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Halton Cost Sheet Generator** - a Streamlit application for managing Halton canopy projects, generating Excel cost sheets, and creating Word quotation documents. The application supports multi-level projects with complex canopy configurations, fire suppression systems, and specialized Halton air handling (RecoAir) systems.

## Development Commands

### Running the Application
```bash
# Run locally
streamlit run app.py

# Python version: 3.13.5
python3 -m pip install -r requirements.txt
```

### Testing
The project uses manual test scripts rather than a formal test framework:
```bash
# Run specific test scripts
python3 test_sdu_collection.py
python3 test_fire_suppression.py
python3 test_recoair_fix.py
```

## Architecture & Key Components

### Core Application Structure
The application follows a modular architecture with clear separation of concerns:

- **Entry Point**: `app.py` (root) → imports from `src/app.py`
- **Main Application**: `src/app.py` - Streamlit UI orchestrator
- **Business Logic**: `src/utils/` - Core functionality for Excel/Word generation
- **Configuration**: `src/config/` - Business rules, constants, and feature flags
- **Components**: `src/components/` - Reusable UI components and forms

### Excel Generation System (`src/utils/excel.py`)
Handles template-based Excel generation with complex business rules:
- **Template Versions**: R19.1 and R19.2 templates in `templates/excel/`
- **Sheet Types**: CANOPY, FIRE SUPP, EDGE BOX, RECOAIR, Lists
- **Dynamic Sheet Creation**: Creates sheets per area with level-based color coding
- **Business Logic**: 
  - KVI models (without 'F') → MUA volume and supply static = "-"
  - CMWF/CMWI models → Extract static = "-"
- **Cell Mappings**: Project metadata starts at specific cells (C3, C5, C7, G3, G5, G7, K7)

### Word Document Generation (`src/utils/word.py`)
Uses Jinja2 templating for professional quotation generation:
- **Templates**: Main template and RecoAir-specific template in `templates/word/`
- **Data Transformation**: Handles empty values, lighting type normalization, static pressure formatting
- **Fire Suppression Logic**: Auto-detects from FIRE SUPP sheets
- **Multi-document Support**: Creates ZIP when multiple documents needed

### Project Data Flow
1. User inputs project data through Streamlit forms
2. Data stored in session state with structured hierarchy (levels → areas → canopies)
3. Excel generation creates cost sheets from templates
4. Word generation reads Excel data and creates quotations
5. Files saved with standardized naming: `[Project Number] [Type] [Date].ext`

### Feature Flag System
The application uses feature flags in `src/config/constants.py` to control system visibility:
- Enabled: Canopy Systems, RecoAir, Fire Suppression, UV-C, SDU, Wall Cladding
- Hidden/Development: Kitchen Extract, M.A.R.V.E.L., Cyclocell, etc.

### State Management
- Session state for form data persistence
- URL-based state serialization for sharing (`src/utils/state_manager.py`)
- Project data export/import through Excel files

## Important Business Rules

### Fire Suppression Detection
- Automatic detection based on presence of FIRE SUPP sheets
- System type mapping: NOBEL → NOBEL System, AMAREX → AMAREX System, default → Ansul R102
- Tank quantities display as "TBD" when empty

### Canopy Model Logic
- Models validated against `VALID_CANOPY_MODELS` list
- Configuration types: Wall, Island, Single, Double, Corner
- Wall cladding dimensions and positioning tracked per canopy

### Date Formatting
- Internal: datetime objects or DD/MM/YYYY strings
- Display: DD/MM/YYYY format
- File naming: DDMMYYYY format (no separators)

## Common Development Tasks

### Adding New Canopy Models
Update `src/config/business_data.py`:
- Add to `VALID_CANOPY_MODELS` list
- Ensure proper categorization (KVI vs others)

### Modifying Excel Templates
1. Update template in `templates/excel/`
2. Adjust cell mappings in `src/utils/excel.py` if needed
3. Test with various project configurations

### Updating Word Templates
1. Modify Jinja2 template in `templates/word/`
2. Update data transformation logic in `src/utils/word.py`
3. Test fire suppression detection and multi-area scenarios

### Debugging Excel Generation
Use debug scripts to inspect:
```bash
python3 debug_excel_fire_suppression.py
python3 debug_sheet_values.py
python3 check_excel_cladding.py
```

### Debugging Word Generation
```bash
python3 debug_word_template.py
python3 debug_word_comprehensive.py
```

## Key File Locations

- **Templates**: `templates/excel/` and `templates/word/`
- **Generated Files**: `output/` (gitignored)
- **Business Data**: `src/config/business_data.py`
- **Constants**: `src/config/constants.py`
- **Debug Scripts**: Root directory (`debug_*.py`, `test_*.py`)