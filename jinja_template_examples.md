# Jinja Template Examples for HVAC Pricing Schedules

This document shows how to use the pricing data in Jinja2 templates for Word document generation.

## Available Pricing Data

The `pricing_totals` object contains comprehensive pricing information:

```python
pricing_totals = {
    'total_canopies': 3,
    'total_canopy_price': 3500.00,
    'total_fire_suppression_price': 2890.00,
    'total_delivery_installation': 800.00,
    'total_commissioning': 500.00,
    'project_total': 7690.00,
    'areas': [
        {
            'level_name': 'Ground Floor',
            'area_name': 'Main Kitchen',
            'level_area_combined': 'Ground Floor - Main Kitchen',
            'canopy_count': 2,
            'canopy_total': 2700.00,
            'fire_suppression_total': 2890.00,
            'delivery_installation_price': 500.00,
            'commissioning_price': 300.00,
            'canopy_schedule_subtotal': 3500.00,  # canopy_total + delivery + commissioning
            'area_subtotal': 6390.00,  # includes fire suppression
            'canopies': [...]
        }
    ]
}
```

## Template Examples

### 1. Basic Pricing Schedule by Area

```jinja2
{% for area in pricing_totals.areas %}
{{ area.level_area_combined | upper }}

CANOPY SCHEDULE
ITEM     DESCRIPTION                    PRICE
--------------------------------------------------
{% for canopy in area.canopies %}
{{ canopy.reference_number }}    Halton {{ canopy.model }} canopy, ex-works    {{ format_currency(canopy.canopy_price) }}
{% endfor %}
{% if area.delivery_installation_price > 0 %}
         Delivery & Installation        {{ format_currency(area.delivery_installation_price) }}
{% endif %}
{% if area.commissioning_price > 0 %}
         Commissioning                  {{ format_currency(area.commissioning_price) }}
{% endif %}
                                         SUB TOTAL: {{ format_currency(area.canopy_schedule_subtotal) }}

{% if area.cladding_total > 0 %}
CLADDING SCHEDULE
ITEM     DESCRIPTION                    PRICE
--------------------------------------------------
{% for canopy in area.canopies %}
{% if canopy.has_cladding %}
         Cladding below Item {{ canopy.reference_number }}, supplied and installed    {{ format_currency(canopy.cladding_price) }}
{% endif %}
{% endfor %}
                                         SUB TOTAL: {{ format_currency(area.cladding_total) }}
{% endif %}

{% if area.fire_suppression_total > 0 %}
ANSUL FIRE SUPPRESSION SCHEDULE
ITEM     DESCRIPTION                    PRICE
--------------------------------------------------
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_price > 0 %}
{{ canopy.reference_number }}    Ansul R102 system. Supplied, installed & commissioned.    {{ format_currency(canopy.fire_suppression_price) }}
{% endif %}
{% endfor %}
                                         SUB TOTAL: {{ format_currency(area.fire_suppression_total) }}
{% endif %}

{{ area.level_area_combined | upper }} TOTAL (EXCLUDING VAT): {{ format_currency(area.area_subtotal) }}

{% endfor %}
```

### 2. Project Totals Summary

```jinja2
PROJECT TOTALS
============================================================
Total Canopies: {{ pricing_totals.total_canopies }}
Canopy Equipment: {{ format_currency(pricing_totals.total_canopy_price) }}
Cladding: {{ format_currency(pricing_totals.total_cladding_price) }}
Fire Suppression: {{ format_currency(pricing_totals.total_fire_suppression_price) }}
Delivery & Installation: {{ format_currency(pricing_totals.total_delivery_installation) }}
Commissioning: {{ format_currency(pricing_totals.total_commissioning) }}
============================================================
PROJECT TOTAL (EXCLUDING VAT): {{ format_currency(pricing_totals.project_total) }}
============================================================
```

### 3. Conditional Cladding Display

```jinja2
{% for area in pricing_totals.areas %}
{% if area.cladding_total > 0 %}
{{ area.level_area_combined }} - Cladding Required
{% for canopy in area.canopies %}
{% if canopy.has_cladding %}
- {{ canopy.reference_number }}: {{ format_currency(canopy.cladding_price) }}
  Dimensions: {{ canopy.wall_cladding.width }}X{{ canopy.wall_cladding.height }}mm
  Position: {{ canopy.wall_cladding.position | join(', ') }}
{% endif %}
{% endfor %}
Total Cladding: {{ format_currency(area.cladding_total) }}
{% endif %}
{% endfor %}
```

### 4. Conditional Fire Suppression Display

```jinja2
{% for area in pricing_totals.areas %}
{% if area.fire_suppression_total > 0 %}
{{ area.level_area_combined }} - Fire Suppression Required
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_tank_quantity > 0 %}
- {{ canopy.reference_number }}: {{ canopy.fire_suppression_tank_quantity }} tanks
{% endif %}
{% endfor %}
{% endif %}
{% endfor %}
```

### 5. Table Format for Word Documents

```jinja2
{% for area in pricing_totals.areas %}
{{ area.level_area_combined | upper }}

| ITEM | CANOPY SCHEDULE | PRICE |
|------|-----------------|-------|
{% for canopy in area.canopies %}
| {{ canopy.reference_number }} | Halton {{ canopy.model }} canopy, ex-works | {{ format_currency(canopy.canopy_price) }} |
{% endfor %}
{% if area.delivery_installation_price > 0 %}
| | Delivery & Installation | {{ format_currency(area.delivery_installation_price) }} |
{% endif %}
{% if area.commissioning_price > 0 %}
| | Commissioning | {{ format_currency(area.commissioning_price) }} |
{% endif %}
| | **SUB TOTAL** | **{{ format_currency(area.canopy_schedule_subtotal) }}** |

{% if area.cladding_total > 0 %}
| ITEM | CLADDING SCHEDULE | PRICE |
|------|-------------------|-------|
{% for canopy in area.canopies %}
{% if canopy.has_cladding %}
| | Cladding below Item {{ canopy.reference_number }}, supplied and installed | {{ format_currency(canopy.cladding_price) }} |
{% endif %}
{% endfor %}
| | **SUB TOTAL** | **{{ format_currency(area.cladding_total) }}** |
{% endif %}

{% if area.fire_suppression_total > 0 %}
| ITEM | ANSUL FIRE SUPPRESSION SCHEDULE | PRICE |
|------|--------------------------------|-------|
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_price > 0 %}
| {{ canopy.reference_number }} | Ansul R102 system. Supplied, installed & commissioned. | {{ format_currency(canopy.fire_suppression_price) }} |
{% endif %}
{% endfor %}
| | **SUB TOTAL** | **{{ format_currency(area.fire_suppression_total) }}** |
{% endif %}

**{{ area.level_area_combined | upper }} TOTAL (EXCLUDING VAT): {{ format_currency(area.area_subtotal) }}**

{% endfor %}
```

### 6. Compact Summary Format

```jinja2
PRICING SUMMARY
{% for area in pricing_totals.areas %}
{{ area.level_area_combined }}:
  Canopy Schedule: {{ format_currency(area.canopy_schedule_subtotal) }}
  Cladding: {{ format_currency(area.cladding_total) }}
  Fire Suppression: {{ format_currency(area.fire_suppression_total) }}
  Area Total: {{ format_currency(area.area_subtotal) }}
{% endfor %}

Project Total: {{ format_currency(pricing_totals.project_total) }}
```

## Key Features

1. **Canopy Schedule Subtotal**: `area.canopy_schedule_subtotal` includes canopy prices + delivery + commissioning (excludes fire suppression and cladding)
2. **Cladding Schedule**: `area.cladding_total` includes all cladding prices for the area
3. **Fire Suppression Schedule**: `area.fire_suppression_total` includes all fire suppression prices (base unit + commissioning share + delivery share per unit)
4. **Area Subtotal**: `area.area_subtotal` includes everything (canopy schedule + cladding + fire suppression)
5. **Currency Formatting**: Use `{{ format_currency(amount) }}` for proper £ formatting
6. **Conditional Display**:
   - Use `{% if area.fire_suppression_total > 0 %}` to show fire suppression only when applicable
   - Use `{% if area.cladding_total > 0 %}` to show cladding only when applicable
   - Use `{% if canopy.has_cladding %}` to check individual canopy cladding
7. **Project Totals**: Access overall totals via `pricing_totals.total_*` fields
8. **Cladding Details**: Access wall cladding dimensions and position via `canopy.wall_cladding` object
9. **Fire Suppression Pricing**: Each `canopy.fire_suppression_price` includes base unit price + proportional commissioning + proportional delivery costs

## Usage in Word Templates

In your Word template (.docx), use these Jinja2 expressions within `{{ }}` or `{% %}` blocks. The `format_currency` function is available throughout the template for consistent currency formatting.

## 🔬 UV-C Control Schedule Implementation

### Data Sources

- **UV-C Model**: Fixed as "UV-C" (from Excel cell C12 in EBOX sheets)
- **UV-C Price**: Read from Excel cell N9 in EBOX sheets
- **UV-C Selection**: Area-level option `area.options.uvc`

### Features

- **Conditional Display**: UV-C Control Schedule only appears if area has UV-C selected
- **Fixed Model**: Always shows "UV-C" as the model name
- **Area-Level Pricing**: UV-C price is associated with the entire area, not individual canopies
- **Professional Format**: Matches the style of other schedules

### Template Usage

```jinja2
{% for level in levels %}
  {% for area in level.areas %}
    {% if area.options.uvc %}
{{ area.level_area_name | upper }} - UV-C CONTROL SCHEDULE

| ITEM | UV-C CONTROL SCHEDULE | PRICE |
|------|----------------------|-------|
| UV-C | UV-C System, supplied and installed | {{ format_currency(area.uvc_price) }} |
| | **SUB TOTAL** | **{{ format_currency(area.uvc_price) }}** |

    {% endif %}
  {% endfor %}
{% endfor %}
```

### Alternative Text Format

```jinja2
{% for level in levels %}
  {% for area in level.areas %}
    {% if area.options.uvc %}
{{ area.level_area_name | upper }} - UV-C CONTROL SCHEDULE

ITEM     UV-C CONTROL SCHEDULE                    PRICE
----------------------------------------------------------
UV-C     UV-C System, supplied and installed     {{ format_currency(area.uvc_price) }}
                                                  SUB TOTAL: {{ format_currency(area.uvc_price) }}

    {% endif %}
  {% endfor %}
{% endfor %}
```

### Compact Format

```jinja2
{% for level in levels %}
  {% for area in level.areas %}
    {% if area.options.uvc %}
🔬 {{ area.level_area_name }} - UV-C SYSTEM: {{ format_currency(area.uvc_price) }}
    {% endif %}
  {% endfor %}
{% endfor %}
```

### Integration with Other Schedules

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

{% if area.uvc_price > 0 %}
🔬 UV-C CONTROL SCHEDULE
UV-C     UV-C System, supplied and installed     {{ format_currency(area.uvc_price) }}
                                                  SUB TOTAL: {{ format_currency(area.uvc_price) }}
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

{% if area.fire_suppression_total > 0 %}
🔥 ANSUL FIRE SUPPRESSION SCHEDULE
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_price > 0 %}
{{ canopy.reference_number }}    Ansul R102 system. Supplied, installed & commissioned.    {{ format_currency(canopy.fire_suppression_price) }}
{% endif %}
{% endfor %}
SUB TOTAL: {{ format_currency(area.fire_suppression_total) }}
{% endif %}

💰 {{ area.level_area_combined | upper }} TOTAL (EXCLUDING VAT): {{ format_currency(area.area_subtotal) }}

{% endfor %}
```

## Key Features

1. **Canopy Schedule Subtotal**: `area.canopy_schedule_subtotal` includes canopy prices + delivery + commissioning (excludes fire suppression and cladding)
2. **Cladding Schedule**: `area.cladding_total` includes all cladding prices for the area
3. **Fire Suppression Schedule**: `area.fire_suppression_total` includes all fire suppression prices (base unit + commissioning share + delivery share per unit)
4. **UV-C Schedule**: `area.uvc_price` includes UV-C system price for the area
5. **Area Subtotal**: `area.area_subtotal` includes everything (canopy schedule + cladding + fire suppression + UV-C + SDU + RecoAir)
6. **Currency Formatting**: Use `{{ format_currency(amount) }}` for proper £ formatting
7. **Conditional Display**:
   - Use `{% if area.fire_suppression_total > 0 %}` to show fire suppression only when applicable
   - Use `{% if area.cladding_total > 0 %}` to show cladding only when applicable
   - Use `{% if area.uvc_price > 0 %}` to show UV-C only when applicable
   - Use `{% if area.options.uvc %}` to check if UV-C is selected for the area
   - Use `{% if canopy.has_cladding %}` to check individual canopy cladding
