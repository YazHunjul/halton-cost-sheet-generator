# Level Name Preservation Fix Summary

## Problem
Level names (like "ddd" for Level 1) were disappearing when navigating between steps in the application. This was similar to the area name issue but affected level names.

## Root Cause
When users navigate between steps (using Previous/Next buttons), Streamlit reruns the entire app. During this rerun:
1. The level data structure (`st.session_state.levels`) might lose the level names
2. The text input widgets get recreated, potentially losing their values
3. There was no preservation mechanism for level names (unlike area names which had a preservation system)

## Solution Implemented
Added a comprehensive level name preservation system that mirrors the area name preservation:

### 1. Session State Initialization
```python
# Initialize level name preservation dictionary
if 'preserved_level_names' not in st.session_state:
    st.session_state.preserved_level_names = {}
```

### 2. Preserve Function
```python
def preserve_level_names():
    """Preserve level names before navigation to prevent loss."""
    for level_idx, level in enumerate(st.session_state.levels):
        key = f"level_{level_idx}"
        # First try to get from the level data structure
        name_to_preserve = level.get('level_name', '')
        
        # If empty, try to get from session state key
        if not name_to_preserve:
            level_name_key = f"level_name_{level_idx}"
            if level_name_key in st.session_state:
                name_to_preserve = st.session_state[level_name_key]
        
        # Always preserve, even if empty
        st.session_state.preserved_level_names[key] = name_to_preserve
```

### 3. Restore Function
```python
def restore_level_names():
    """Restore level names after navigation if they were lost."""
    for level_idx, level in enumerate(st.session_state.levels):
        key = f"level_{level_idx}"
        if key in st.session_state.preserved_level_names:
            preserved_name = st.session_state.preserved_level_names[key]
            if preserved_name or not level.get('level_name'):
                level['level_name'] = preserved_name
                
                # Also update session state key
                level_name_key = f"level_name_{level_idx}"
                st.session_state[level_name_key] = preserved_name
```

### 4. Navigation Integration
Both Previous and Next buttons now call `preserve_level_names()` before navigation:
```python
if st.button("← Previous", key="nav_prev"):
    preserve_level_names()  # Preserve level names before navigation
    preserve_area_names()   # Preserve area names before navigation
    st.session_state.current_step -= 1
    st.rerun()
```

### 5. Step Function Integration
All step functions (step2, step3, step4) now restore level names at the beginning:
```python
def step2_project_structure():
    """Step 2: Project Structure"""
    # ... header code ...
    
    # Restore names if they were lost during navigation
    restore_level_names()
    restore_area_names()
```

### 6. Update Function Integration
The `update_level_name()` function now preserves changes:
```python
def update_level_name():
    st.session_state.levels[level_idx]['level_name'] = st.session_state[f"level_name_{level_idx}"]
    # Also preserve in the preservation dictionary
    preservation_key = f"level_{level_idx}"
    st.session_state.preserved_level_names[preservation_key] = st.session_state[f"level_name_{level_idx}"]
```

### 7. Excel Loading Integration
When loading from Excel, level names are now preserved:
```python
# Also preserve level names when loading from Excel
for level_idx, level in enumerate(st.session_state.levels):
    if level.get('level_name'):
        key = f"level_{level_idx}"
        st.session_state.preserved_level_names[key] = level['level_name']
        # Also set the session state key
        level_name_key = f"level_name_{level_idx}"
        st.session_state[level_name_key] = level['level_name']
```

## Result
Level names are now preserved during navigation:
- When users click Previous/Next, level names are saved before navigation
- When the new step loads, level names are restored from the preserved state
- Level names persist correctly even when navigating back and forth between steps
- Level names are properly preserved when loading from Excel files

## Testing
To verify the fix works:
1. Enter a level name (e.g., "ddd" for Level 1)
2. Navigate to the next step using the Next button
3. Navigate back using the Previous button
4. The level name "ddd" should still be visible and preserved

## Files Modified
- `src/app.py`: Added the complete level name preservation system

## Backup
A backup of the original file was created at `src/app.py.backup` before implementing the fix.