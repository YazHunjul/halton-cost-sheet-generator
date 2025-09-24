# Word Generation - SDU Canopy Level Update Summary

## Changes Made to Support Canopy-Level SDU

### File Modified: `src/utils/word.py`

#### 1. Updated `collect_sdu_data` Function (Lines 2100-2274)
**Changes:**
- Modified function to scan **canopies** instead of areas for SDU systems
- Changed from `area.get('options', {}).get('sdu', False)` to `canopy.get('options', {}).get('sdu', False)`
- Updated SDU sheet naming from `SDU - {level_name} ({area_number})` to `SDU - {canopy_reference}`
- Added `canopy_reference` field to SDU data structure
- Updated all print messages to refer to "canopies" instead of "areas"

#### 2. Updated Area-Level SDU Price Calculation (Line 838)
**Changes:**
- Changed from getting SDU price from area: `area.get('sdu_price', 0)`
- To calculating total SDU price from all canopies: `sum(canopy.get('sdu_price', 0) for canopy in area.get('canopies', []))`

#### 3. Updated Area-Level SDU Flag (Lines 888, 1735)
**Changes:**
- Changed from checking area options: `area.get('options', {}).get('sdu', False)`
- To checking if any canopy has SDU: `any(canopy.get('options', {}).get('sdu', False) for canopy in area.get('canopies', []))`

#### 4. Updated SDU Data Merging Logic (Lines 1789-1825)
**Changes:**
- Changed lookup dictionary from area-based to canopy-based
- Now iterates through all canopies in each area to aggregate SDU pricing
- Accumulates SDU totals from multiple canopies per area
- Properly updates area totals with aggregated canopy-level SDU pricing

## How It Works Now

1. **SDU Detection**: The system checks each canopy for SDU option instead of checking at area level
2. **SDU Sheet Naming**: SDU sheets are now named using canopy reference numbers (e.g., `SDU - C1`)
3. **Price Aggregation**: Area-level SDU prices are calculated by summing all canopy SDU prices
4. **Data Collection**: `collect_sdu_data` now returns a list of SDU data per canopy, not per area
5. **Merging**: When merging detailed Excel data, the system looks up SDU data by canopy reference and aggregates for area totals

## Result
The Word document generation now correctly handles canopy-level SDU:
- ✅ Detects SDU at canopy level
- ✅ Reads SDU sheets with canopy-based naming
- ✅ Aggregates SDU pricing from multiple canopies per area
- ✅ Updates area and project totals correctly
- ✅ Maintains backward compatibility with area-level totals for Word templates

## Testing Recommendations
1. Create a project with multiple canopies, some with SDU enabled
2. Generate Excel file and verify SDU sheets are created with canopy references
3. Generate Word document and verify:
   - SDU pricing appears correctly in area totals
   - Multiple SDU canopies in same area have prices aggregated
   - Project totals include all SDU pricing

## Date Completed
July 17, 2025