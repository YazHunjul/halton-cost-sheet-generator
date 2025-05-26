# UV Canopy Detection Implementation

## Overview

Implemented a global project-level flag `has_uv` that detects if there are any UV canopies (like UVI, UVF models) anywhere in the entire project.

## Changes Made

### 1. Updated `analyze_project_areas()` function in `src/utils/word.py`

**Before:** Checked for area-level UV-C system option (`area.options.uvc`)
**After:** Checks for UV canopy models (models starting with "UV") across all canopies in the project

```python
# Check for UV canopy models (UVI, UVF, etc.) across all canopies in the project
for canopy in canopies:
    model = canopy.get('model', '').upper().strip()
    if model.startswith('UV'):  # UV models like UVI, UVF, etc.
        has_uv = True
```

**Return value changed:** `(has_canopies, has_recoair, is_recoair_only, has_uv)` (was `has_uvc`)

### 2. Updated all function calls to use `has_uv` instead of `has_uvc`

**Files updated:**

- `src/utils/word.py`: Updated 3 locations where `analyze_project_areas()` is called
- `src/app.py`: Updated 2 locations where `analyze_project_areas()` is called

### 3. Updated template context

**Template variable changed:** `has_uvc` → `has_uv`

The global flag is now available in Word templates as `has_uv`.

## Detection Logic

The flag detects UV canopies at the **project level**, not per area:

- **UVI models**: ✅ Detected (model starts with "UV")
- **UVF models**: ✅ Detected (model starts with "UV")
- **KV, KVF, CMWF models**: ❌ Not detected (don't start with "UV")

## Template Usage

In Word templates, you can now use:

```jinja2
{% if has_uv %}
Specific Notes:
UV canopy systems are included in this project.
{% endif %}
```

## Testing

Created comprehensive test cases covering:

- ✅ Projects with UVI models
- ✅ Projects with UVF models
- ✅ Projects without UV models
- ✅ Mixed projects (UV + RecoAir)
- ✅ UV models in different areas

All tests pass successfully.

## Important Notes

### Area-level UV-C vs Project-level UV Canopies

- **Area-level UV-C option** (`area.options.uvc`): Still used for creating EBOX sheets in Excel
- **Project-level UV canopy detection** (`has_uv`): New global flag for Word templates

These are **separate concepts**:

- UV-C system option = Area-level feature for UV-C control systems
- UV canopy detection = Project-level flag for UV canopy models (UVI, UVF, etc.)

### Backward Compatibility

- Excel generation unchanged (still uses area-level UV-C options)
- Only Word template generation uses the new global UV canopy flag
- No breaking changes to existing functionality
