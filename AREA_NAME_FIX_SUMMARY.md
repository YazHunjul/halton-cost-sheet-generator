# Area Name Preservation Fix Summary

## Issue
Area 1 in the main Kitchen loses its name when navigating away and back, even after trying to sync session state.

## Root Cause
The issue was in the `step2_project_structure()` function where the code was unconditionally syncing session state with the area's name from the data structure:

```python
# Always update session state to match the actual area name
st.session_state[area_name_key] = area['name']
```

If the area's name in the data structure (`st.session_state.levels[level_idx]['areas'][area_idx]['name']`) was empty or had been lost during navigation, this would overwrite the session state with an empty value.

## Solution Implemented

### 1. Added Smart Session State Synchronization
Updated the area name handling logic to:
- Only sync from data structure to session state if the area has a name
- If area name is empty but session state has a value, restore it to the data structure
- Initialize session state for area names if not present

### 2. Added Area Name Preservation System
Created a preservation system with three components:

#### a. Preservation Dictionary
Added `preserved_area_names` to session state to store area names independently of the main data structure.

#### b. Preserve Function
Created `preserve_area_names()` function that saves all area names before navigation:
```python
def preserve_area_names():
    """Preserve area names before navigation to prevent loss."""
    for level_idx, level in enumerate(st.session_state.levels):
        for area_idx, area in enumerate(level.get('areas', [])):
            key = f"level_{level_idx}_area_{area_idx}"
            if area.get('name'):
                st.session_state.preserved_area_names[key] = area['name']
```

#### c. Restore Function
Created `restore_area_names()` function that restores area names after navigation:
```python
def restore_area_names():
    """Restore area names after navigation if they were lost."""
    for level_idx, level in enumerate(st.session_state.levels):
        for area_idx, area in enumerate(level.get('areas', [])):
            key = f"level_{level_idx}_area_{area_idx}"
            # If area name is empty but we have a preserved name, restore it
            if not area.get('name') and key in st.session_state.preserved_area_names:
                area['name'] = st.session_state.preserved_area_names[key]
```

### 3. Integration Points
- Called `preserve_area_names()` in navigation buttons before changing steps
- Called `restore_area_names()` at the beginning of each step function
- Updated area name change handler to also save to preservation dictionary
- Added preservation when loading data from Excel

### 4. Debug Features
Added a debug checkbox in step 2 to show:
- Current area names in the data structure
- Area names stored in session state
- Helps identify when/where names are being lost

## Files Modified
1. `/Users/yazan/Desktop/Efficiency/UKCS/src/app.py` - Main application file with all the preservation logic

## Testing
Created test script `/Users/yazan/Desktop/Efficiency/UKCS/test_area_name_preservation.py` to verify the fix works correctly.

## Expected Behavior After Fix
1. Enter "Kitchen" as Area 1 name in Step 2
2. Navigate to Step 3 or Step 4
3. Navigate back to Step 2
4. "Kitchen" name should still be displayed and preserved

The fix ensures area names are preserved across navigation by maintaining them in a separate dictionary that survives step changes.