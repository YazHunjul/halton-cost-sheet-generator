# Feature Flags Documentation

## Overview

The Halton Cost Sheet Generator uses feature flags to control which systems and components are displayed to users. This allows you to keep systems in the documentation and templates while hiding them from the user interface until they are ready for production use.

## Feature Flag Configuration

Feature flags are defined in `src/config/constants.py` in the `FEATURE_FLAGS` dictionary:

```python
FEATURE_FLAGS = {
    # Kitchen systems
    "kitchen_extract_system": False,
    "kitchen_makeup_air_system": False,

    # Advanced control systems
    "marvel_system": False,  # M.A.R.V.E.L. System (DCKV)

    # Ceiling systems
    "cyclocell_cassette_ceiling": False,

    # Additional equipment
    "reactaway_unit": False,

    # Future systems (placeholders)
    "dishwasher_extract": False,
    "gas_interlocking": False,
    "pollustop_unit": False,
}
```

## How Feature Flags Work

### 1. Checking Feature Status

Use the `is_feature_enabled()` function to check if a feature is enabled:

```python
from config.constants import is_feature_enabled

if is_feature_enabled('marvel_system'):
    # Show MARVEL system to user
    pass
else:
    # Hide MARVEL system from user
    pass
```

### 2. Excel Generation

In the Excel generation (`src/utils/excel.py`), feature flags control:

- **JOB TOTAL Sheet**: Whether pricing formulas are included for disabled systems
- **System Sheets**: Whether specific system sheets are created
- **Pricing Calculations**: Whether disabled systems contribute to totals

Example from Excel generation:

```python
# Conditionally add MARVEL system if enabled
if is_feature_enabled('marvel_system'):
    job_total_sheet['T20'] = f"=PRICING_SUMMARY!B{summary_row + 6}"  # MARVEL Total
else:
    job_total_sheet['T20'] = 0  # Hide MARVEL if not enabled
```

### 3. Word Document Generation

In the Word generation (`src/utils/word.py`), feature flags are passed to templates as context variables:

```python
context = {
    # ... other context variables ...
    'show_marvel_system': is_feature_enabled('marvel_system'),
    'show_reactaway_unit': is_feature_enabled('reactaway_unit'),
    'show_pollustop_unit': is_feature_enabled('pollustop_unit'),
    # ... etc
}
```

These can then be used in Word templates with Jinja2 syntax:

```jinja2
{% if show_marvel_system %}
<!-- MARVEL system content here -->
{% endif %}
```

## Current System Status

### ✅ Enabled Systems (Ready for Production)

- Canopy Systems
- RecoAir Systems
- Fire Suppression
- UV-C Systems
- SDU (Supply Diffusion Units)
- Wall Cladding

### ❌ Disabled Systems (Hidden from Users)

- **Kitchen Extract System** (`kitchen_extract_system`)
- **Kitchen Make-Up Air System** (`kitchen_makeup_air_system`)
- **M.A.R.V.E.L. System** (`marvel_system`) - Demand-based control system
- **Cyclocell Cassette Ceiling** (`cyclocell_cassette_ceiling`)
- **Reactaway Unit** (`reactaway_unit`) - UV-C filtration module
- **Dishwasher Extract** (`dishwasher_extract`)
- **Gas Interlocking** (`gas_interlocking`)
- **Pollustop Unit** (`pollustop_unit`)

## Enabling a System

To enable a system for production use:

1. **Update the feature flag** in `src/config/constants.py`:

   ```python
   FEATURE_FLAGS = {
       "marvel_system": True,  # Changed from False to True
       # ... other flags
   }
   ```

2. **Test the system** thoroughly:

   - Create test projects with the system enabled
   - Generate Excel cost sheets
   - Generate Word quotations
   - Verify pricing calculations
   - Check template rendering

3. **Update documentation** as needed:
   - Update this file to move the system from "Disabled" to "Enabled"
   - Update user guides if necessary

## Template Integration

### Word Templates

Word templates can use feature flags to conditionally display content:

```jinja2
{% if show_kitchen_extract_system %}
## Kitchen Extract System

| Item | Description |
|------|-------------|
| 5.01 | Galvanised steel duct, LV Class 'A' |
| 5.02 | Fire rated duct |
| 5.03 | Stainless steel duct |
{% endif %}
```

### Excel Templates

Excel generation automatically respects feature flags for:

- Pricing formulas in JOB TOTAL sheet
- System-specific sheets creation
- Summary calculations

## Benefits of Feature Flags

1. **Gradual Rollout**: Enable systems one at a time as they're ready
2. **Documentation Preservation**: Keep system information in templates without showing to users
3. **Easy Testing**: Enable systems in development/testing environments
4. **Clean User Experience**: Users only see completed, tested systems
5. **Future-Proofing**: Placeholder flags for planned systems

## Best Practices

1. **Default to False**: New systems should start with `False` flag
2. **Descriptive Names**: Use clear, descriptive flag names
3. **Documentation**: Update this file when changing flag status
4. **Testing**: Thoroughly test before enabling in production
5. **Gradual Enablement**: Enable one system at a time to isolate issues

## Troubleshooting

### System Not Showing After Enabling Flag

1. Check that the flag is correctly set to `True`
2. Verify the system is properly integrated in templates
3. Check for any conditional logic that might be hiding the system
4. Restart the Streamlit application to ensure changes are loaded

### Excel Formulas Not Working

1. Verify the feature flag is checked in `update_job_total_sheet()`
2. Check that the pricing summary includes the system
3. Ensure the system has proper pricing data

### Word Template Not Rendering System

1. Check that the feature flag is passed to template context
2. Verify the Jinja2 syntax in the template
3. Ensure the template variable name matches the context variable

## Future Development

When adding new systems:

1. Add a feature flag to `FEATURE_FLAGS` (set to `False`)
2. Implement the system functionality
3. Add conditional logic using `is_feature_enabled()`
4. Update templates with conditional rendering
5. Test thoroughly
6. Enable the flag when ready for production
7. Update this documentation
