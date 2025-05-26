# 📁 Halton Cost Sheet Generator - File Naming Conventions

## Overview

The Halton Cost Sheet Generator now uses standardized filename conventions for all generated files to ensure consistency and easy identification.

## 📊 Excel Cost Sheets

### Format: `Project Number Cost Sheet Date`

**Examples:**

- `P12345 Cost Sheet 15012025.xlsx`
- `HAL001 Cost Sheet 28022025.xlsx`
- `PROJ789 Cost Sheet 10032025.xlsx`

**Components:**

- **Project Number**: From project data (e.g., "P12345")
- **"Cost Sheet"**: Fixed identifier
- **Date**: DDMMYYYY format (e.g., "15012025" for 15/01/2025)

**Notes:**

- Date format removes slashes for filename compatibility
- If no date provided, uses current date
- Revision information is stored internally but not in filename

## 📄 Word Quotation Documents

### Main Quotation Format: `Project Number Quotation Date`

**Examples:**

- `P12345 Quotation 15012025.docx`
- `HAL001 Quotation 28022025.docx`
- `PROJ789 Quotation 10032025.docx`

### RecoAir Quotation Format: `Project Number RecoAir Quotation Date`

**Examples:**

- `P12345 RecoAir Quotation 15012025.docx`
- `HAL001 RecoAir Quotation 28022025.docx`
- `PROJ789 RecoAir Quotation 10032025.docx`

### Multiple Documents (ZIP) Format: `Project Number Quotations Date`

**Examples:**

- `P12345 Quotations 15012025.zip`
- `HAL001 Quotations 28022025.zip`
- `PROJ789 Quotations 10032025.zip`

**ZIP Contents:**

- Main quotation document (for canopies, SDU, etc.)
- RecoAir quotation document (for RecoAir systems)

## 🔄 Document Generation Logic

### Single Document Scenarios

1. **Canopy-only projects**: Generate main quotation only

   - Filename: `Project Number Quotation Date.docx`

2. **RecoAir-only projects**: Generate RecoAir quotation only
   - Filename: `Project Number RecoAir Quotation Date.docx`

### Multiple Document Scenarios

3. **Mixed projects** (Canopies + RecoAir): Generate both documents in ZIP
   - ZIP filename: `Project Number Quotations Date.zip`
   - Contains:
     - `Project Number Quotation Date.docx` (main quotation)
     - `Project Number RecoAir Quotation Date.docx` (RecoAir quotation)

## 📅 Date Formatting

### Input Formats Supported

- `DD/MM/YYYY` (e.g., "15/01/2025")
- `DD-MM-YYYY` (e.g., "15-01-2025")
- Empty/null (uses current date)

### Output Format

- `DDMMYYYY` (e.g., "15012025")
- Removes slashes and hyphens for filename compatibility
- Always 8 digits for consistency

## 🔧 Implementation Details

### Excel Generation

```python
# Format: "Project Number Cost Sheet Date"
output_filename = f"{project_number} Cost Sheet {formatted_date}.xlsx"
```

### Word Generation

```python
# Main quotation: "Project Number Quotation Date"
main_filename = f"{project_number} Quotation {date_str}.docx"

# RecoAir quotation: "Project Number RecoAir Quotation Date"
recoair_filename = f"{project_number} RecoAir Quotation {date_str}.docx"

# ZIP file: "Project Number Quotations Date"
zip_filename = f"{project_number} Quotations {date_str}.zip"
```

### Date Formatting Function

```python
def format_date_for_filename(date_str: str) -> str:
    if date_str:
        return date_str.replace('/', '').replace('-', '')
    else:
        return datetime.now().strftime("%d%m%Y")
```

## 📋 Benefits

### Consistency

- All files follow the same naming pattern
- Easy to identify file type and project
- Chronological sorting by date

### Professional Appearance

- Clean, readable filenames
- No underscores or special characters
- Proper capitalization

### File Management

- Easy to search and filter
- Clear identification of document type
- Date-based organization

### User Experience

- Predictable filename patterns
- No confusion about file contents
- Professional delivery to clients

## 🔄 Revision Handling

### Excel Revisions

- Filename format remains: `Project Number Cost Sheet Date.xlsx`
- Revision letter stored internally in spreadsheet
- New revision creates new file with updated date

### Word Document Revisions

- Generate new documents from updated Excel file
- Revision information included in document content
- Filename reflects the date of generation

## 📝 Examples by Project Type

### Canopy Project

```
P12345 Cost Sheet 15012025.xlsx
P12345 Quotation 15012025.docx
```

### RecoAir Project

```
P12345 Cost Sheet 15012025.xlsx
P12345 RecoAir Quotation 15012025.docx
```

### Mixed Project

```
P12345 Cost Sheet 15012025.xlsx
P12345 Quotations 15012025.zip
  ├── P12345 Quotation 15012025.docx
  └── P12345 RecoAir Quotation 15012025.docx
```

## 🎯 Quality Assurance

### Validation

- Project number validation (no special characters)
- Date format validation and conversion
- Filename length limits respected

### Error Handling

- Fallback to current date if date invalid
- Default project number if missing
- Graceful handling of special characters

### Testing

- Test with various project numbers
- Test with different date formats
- Test with missing data scenarios

---

## 📞 Support

For questions about filename conventions or to request changes, please refer to the development team or update this documentation accordingly.
