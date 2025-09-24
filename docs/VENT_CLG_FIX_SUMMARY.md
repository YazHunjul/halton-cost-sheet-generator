# VENT CLG Fix Summary

## Issue Description
The VENT CLG sheet metadata was not being written correctly:
1. **Project name, location, and date** were being written to column C instead of column F
2. **Sales manager/estimator initials** were correctly mapped to C7, but there was a conflict because date was also trying to use C7
3. **Sales manager/estimator initials** were using plain estimator name instead of combined initials (Sales Contact + Estimator)

## Root Cause
The `write_vent_clg_metadata` function in `src/utils/excel.py` had two issues:
1. **Incorrect cell mappings** (lines 1958-1966):
   - `"project_name": "C3"` (should be F3)
   - `"project_location": "C5"` (should be F5)  
   - `"date": "C7"` (should be F7, was conflicting with estimator at C7)
2. **Missing combined initials logic** (line 1978): Used plain `project_data.get("estimator", "")` instead of generating combined initials like other sheet types

## Fix Applied
**File**: `src/utils/excel.py:1958-1966, 1975-1990`
**Function**: `write_vent_clg_metadata`

### Before:
```python
# Use C column for basic project info as specified by user
cell_mappings = {
    "project_number": "C3",  # Job No
    "company": "C5",         # Customer
    "estimator": "C7",       # Sales Manager / Estimator
    "project_name": "C3",    # Project Name (same as job number)
    "project_location": "C5", # Project Location (same as customer)
    "date": "C7",            # Date (same as estimator)
    "revision": "K7",        # Revision
}

# Write project metadata using the mappings
write_to_cell_safe(sheet, cell_mappings["estimator"], project_data.get("estimator", ""))
```

### After:
```python
# VENT CLG-specific cell mappings (F columns for project name/location/date)
cell_mappings = {
    "project_number": "C3",  # Job No
    "company": "C5",         # Customer
    "estimator": "C7",       # Sales Manager / Estimator Initials
    "project_name": "F3",    # Project Name (F column as requested)
    "project_location": "F5", # Project Location (F column as requested)
    "date": "F7",            # Date (F column as requested)
    "revision": "K7",        # Revision
}

# Generate combined initials for sales manager/estimator (Sales Contact + Estimator)
estimator_name = project_data.get("estimator", "")
if estimator_name:
    from utils.word import get_combined_initials
    sales_contact_name = project_data.get('sales_contact', '')
    combined_initials = get_combined_initials(sales_contact_name, estimator_name)
    write_to_cell_safe(sheet, cell_mappings["estimator"], combined_initials)
```

## Result
Now the VENT CLG sheet will correctly populate:
- **C7**: Sales manager/estimator **combined initials** (Sales Contact + Estimator, no conflicts)
- **F3**: Project name
- **F5**: Project location  
- **F7**: Date

## Date Fixed
July 17, 2025

## Files Modified
- `src/utils/excel.py` (lines 1958-1966, 1975-1990)