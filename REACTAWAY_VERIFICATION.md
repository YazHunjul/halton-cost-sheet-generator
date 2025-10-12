# Reactaway Feature Verification Guide

## ✅ Feature Status: **WORKING**

The Reactaway option has been successfully implemented and tested.

## Test Results

A test was run with the following configuration:
- **Project**: Reactaway Test Project
- **Area**: Test Area with Reactaway
- **Reactaway Option**: ✅ Enabled
- **Result**: ✅ REACTAWAY - Ground Floor (1) sheet was created successfully

## How to Use Reactaway

### 1. Enable Reactaway in the UI

When creating or editing a project, for each area you can check the **Reactaway** checkbox:

**Location in UI**:
- Single Page Setup → Area Options
- Create Revision → Levels & Areas → Area Options
- Step 2: Project Structure → Area Options

The checkbox appears alongside other options:
- UV-C
- RecoAir
- Marvel
- Vent CLG
- Pollustop
- Aerolys
- XEU
- **Reactaway** ← New option

### 2. Generate Excel

When you generate the Excel file, a REACTAWAY sheet will be automatically created for each area that has the Reactaway option enabled.

**Sheet naming format**: `REACTAWAY - [Level Name] ([Area Number])`

Example: `REACTAWAY - Ground Floor (1)`

### 3. Finding the Reactaway Sheet

The REACTAWAY sheet will appear in your Excel workbook. Since the template has many sheets (262+), you may need to:

1. **Scroll through the sheet tabs** at the bottom of Excel
2. **Use sheet navigation**: Right-click on the sheet tab navigation arrows → "More Sheets..." → Find REACTAWAY
3. **Use Ctrl+F (Windows) or Cmd+F (Mac)** in the sheet list to search for "REACTAWAY"

The sheet will have:
- Tab color matching the area's color scheme
- Project metadata (project name, number, date, etc.)
- Title in B1: `[Level Name] - [Area Name] - REACTAWAY SYSTEM`

## Verification Steps

To verify Reactaway is working in your project:

1. Create a new project or open existing one
2. Add an area
3. Check the **Reactaway** checkbox for that area
4. Add at least one canopy to the area (or leave empty if testing area-only options)
5. Generate Excel
6. Open the generated Excel file
7. Look for the REACTAWAY sheet tab (scroll through tabs or use search)

## Technical Details

### Files Modified:
- `src/app.py`: Added Reactaway checkbox to all area forms
- `src/utils/excel.py`: Added REACTAWAY sheet creation logic

### Implementation:
- Reactaway follows the same pattern as Pollustop, Aerolys, and XEU
- Sheet is created from template "REACTAWAY" sheets
- Works with both canopy areas and non-canopy areas
- Properly handles sheet naming, coloring, and metadata

## Troubleshooting

**Issue**: "I don't see the Reactaway sheet"

**Solutions**:
1. Verify the Reactaway checkbox was checked before generating Excel
2. Scroll through all sheet tabs in Excel (there are many sheets)
3. Right-click sheet navigation arrows → "More Sheets..." → Search for "REACTAWAY"
4. Check the project summary before generating to confirm Reactaway shows as "Yes"

**Issue**: "Checkbox not showing up"

**Solutions**:
1. Refresh the Streamlit app (F5 or Cmd+R)
2. Check you're on the latest code version
3. The checkbox should appear in a row with 8 options (UV-C, RecoAir, Marvel, Vent CLG, Pollustop, Aerolys, XEU, Reactaway)

## Test File

Run the test script to verify functionality:

```bash
cd /Users/yazan/Desktop/Efficiency/UKCS
source venv/bin/activate
python test_reactaway.py
```

Expected output: `✅ SUCCESS: Found REACTAWAY sheet(s)`
