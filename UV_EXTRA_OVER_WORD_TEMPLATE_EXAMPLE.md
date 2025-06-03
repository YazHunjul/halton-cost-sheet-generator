# UV Extra Over Word Template Integration Guide

## Overview

UV Extra Over functionality is now fully integrated into the Word document generation system. This guide shows how to use UV Extra Over data in your Word templates.

## Available Data Structure

When UV Extra Over is enabled for an area, the following data is available in your Word templates:

### Project Level

```python
pricing_totals = {
    'total_uv_extra_over_cost': 450.00,  # Total UV Extra Over cost across all areas
    'has_any_uv_extra_over': True,      # True if any area has UV Extra Over
    # ... other pricing data
}
```

### Area Level

```python
area = {
    'has_uv_extra_over': True,      # True if this specific area has UV Extra Over
    'uv_extra_over_cost': 150.00,   # UV Extra Over cost for this area only
    'level_area_combined': 'Level 1 - Main Kitchen',
    # ... other area data
}
```

## Word Template Examples

### 1. Conditional UV Extra Over Section

Only show UV Extra Over schedule if any area has it:

```jinja2
{% if pricing_totals.has_any_uv_extra_over %}

UV EXTRA OVER SCHEDULE

The following areas require UV Extra Over pricing:

{% for area in pricing_totals.areas %}
{% if area.has_uv_extra_over %}
{{ area.level_area_combined | upper }}
UV Extra Over Price: {{ format_currency(area.uv_extra_over_cost) }}

{% endif %}
{% endfor %}

TOTAL UV EXTRA OVER: {{ format_currency(pricing_totals.total_uv_extra_over_cost) }}

{% endif %}
```

### 2. UV Extra Over Table Format

Professional table format for UV Extra Over schedules:

```jinja2
{% if pricing_totals.has_any_uv_extra_over %}

| AREA | UV EXTRA OVER SCHEDULE | PRICE |
|------|------------------------|-------|
{% for area in pricing_totals.areas %}
{% if area.has_uv_extra_over %}
| {{ area.level_area_combined }} | UV canopy upgrade cost | {{ format_currency(area.uv_extra_over_cost) }} |
{% endif %}
{% endfor %}
| | **TOTAL UV EXTRA OVER** | **{{ format_currency(pricing_totals.total_uv_extra_over_cost) }}** |

{% endif %}
```

### 3. Integrated with Main Pricing Schedule

Include UV Extra Over in the main pricing schedule for each area:

```jinja2
{% for area in pricing_totals.areas %}
{{ area.level_area_combined | upper }}

📋 CANOPY SCHEDULE
{% for canopy in area.canopies %}
{{ canopy.reference_number }}    Halton {{ canopy.model }} canopy, ex-works    {{ format_currency(canopy.canopy_price) }}
{% endfor %}
         Delivery & Installation        {{ format_currency(area.delivery_installation_price) }}
         Commissioning                  {{ format_currency(area.commissioning_price) }}
                                         SUB TOTAL: {{ format_currency(area.canopy_schedule_subtotal) }}

{% if area.has_uv_extra_over %}
⚡ UV EXTRA OVER SCHEDULE
         UV canopy upgrade               {{ format_currency(area.uv_extra_over_cost) }}
                                         SUB TOTAL: {{ format_currency(area.uv_extra_over_cost) }}
{% endif %}

{% if area.cladding_total > 0 %}
🏗️ CLADDING SCHEDULE
{% for canopy in area.canopies %}
{% if canopy.has_cladding %}
         Cladding below Item {{ canopy.reference_number }}, supplied and installed    {{ format_currency(canopy.cladding_price) }}
{% endif %}
{% endfor %}
                                         SUB TOTAL: {{ format_currency(area.cladding_total) }}
{% endif %}

💰 {{ area.level_area_combined | upper }} TOTAL (EXCLUDING VAT): {{ format_currency(area.area_subtotal) }}

{% endfor %}
```

### 4. UV Extra Over Summary Only

Simple summary format showing only the UV Extra Over totals:

```jinja2
{% if pricing_totals.has_any_uv_extra_over %}

UV EXTRA OVER SUMMARY
{% for area in pricing_totals.areas %}
{% if area.has_uv_extra_over %}
{{ area.level_area_combined }}: {{ format_currency(area.uv_extra_over_cost) }}
{% endif %}
{% endfor %}

Total UV Extra Over Cost: {{ format_currency(pricing_totals.total_uv_extra_over_cost) }}

{% endif %}
```

### 5. Detailed UV Extra Over Description

Include detailed explanation of what UV Extra Over means:

```jinja2
{% if pricing_totals.has_any_uv_extra_over %}

UV EXTRA OVER EXPLANATION

The following areas include UV-capable canopy systems. The UV Extra Over cost represents the additional investment required for UV functionality compared to standard canopy systems:

{% for area in pricing_totals.areas %}
{% if area.has_uv_extra_over %}
{{ area.level_area_combined }}:
- Standard canopy system (base cost included in main schedule)
- UV canopy upgrade: {{ format_currency(area.uv_extra_over_cost) }}
- Total UV canopy investment: Standard cost + {{ format_currency(area.uv_extra_over_cost) }} upgrade

{% endif %}
{% endfor %}

TOTAL UV UPGRADE INVESTMENT: {{ format_currency(pricing_totals.total_uv_extra_over_cost) }}

This represents the additional cost for UV-capable systems compared to standard equivalents.

{% endif %}
```

## Implementation Notes

### Key Features

1. **Conditional Display**: UV Extra Over sections only appear when `has_any_uv_extra_over` is True
2. **Area-Specific**: Each area can independently have UV Extra Over enabled
3. **Cost Calculation**: Automatically calculates UV canopy cost - non-UV equivalent cost
4. **Currency Formatting**: Use `{{ format_currency(amount) }}` for proper £ formatting
5. **Template Flexibility**: Can be integrated into existing pricing schedules or shown separately

### Data Flow

1. **Excel Generation**: UV Extra Over sheets are created with cost calculations
2. **Excel Reading**: UV Extra Over costs are extracted from UV canopy sheets
3. **Word Generation**: UV Extra Over data is included in template context
4. **Template Rendering**: UV Extra Over sections are conditionally displayed

### Example Usage in Practice

```jinja2
{% for area in pricing_totals.areas %}
    {% if area.has_canopies %}
        {{ area.level_area_combined | upper }} - PRICING SCHEDULE

        Standard Canopy Equipment: {{ format_currency(area.canopy_total) }}
        {% if area.has_uv_extra_over %}
        UV System Upgrade: {{ format_currency(area.uv_extra_over_cost) }}
        {% endif %}
        Delivery & Installation: {{ format_currency(area.delivery_installation_price) }}
        Commissioning: {{ format_currency(area.commissioning_price) }}

        AREA TOTAL: {{ format_currency(area.area_subtotal) }}
    {% endif %}
{% endfor %}
```

## Testing and Verification

✅ **Excel Generation**: UV Extra Over sheets created with proper naming
✅ **Cost Calculation**: UV vs non-UV price difference calculated correctly  
✅ **Excel Reading**: UV Extra Over costs extracted from Excel files
✅ **Word Integration**: UV Extra Over data available in template context
✅ **Template Context**: Both area-level and project-level UV Extra Over data available

## Future Enhancements

The UV Extra Over system is designed to be extensible for future "extra over" types:

- **SDU Extra Over**: For upgraded SDU systems
- **Fire Suppression Extra Over**: For enhanced fire suppression options
- **Lighting Extra Over**: For premium lighting upgrades
- **Custom Extra Over**: For project-specific upgrades

The template structure supports multiple extra over types using the same conditional logic pattern.
