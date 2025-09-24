# HVAC Pricing Schedules Implementation Summary

## Overview

This document summarizes the comprehensive pricing schedule functionality implemented for the HVAC Project Management Tool, including canopy schedules, cladding schedules, and fire suppression schedules.

## 🏗️ Cladding Schedule Implementation

### Data Sources

- **Cladding Price**: Read from Excel cells N19, N20, N21, etc.
- **Cladding Dimensions**: Read from Excel cells P19, P20, P21, etc. (format: "1000X2100")
- **Cladding Position**: Read from Excel cells Q19, Q20, Q21, etc. (format: "rear/left hand")

### Features

- **Conditional Display**: Cladding schedule only appears if area has canopies with cladding
- **Individual Pricing**: Each canopy can have its own cladding price
- **Position Descriptions**: Automatically formats position descriptions (e.g., "rear and left hand walls")
- **Area Totals**: Calculates total cladding cost per area

### Template Usage

```jinja2
{% if area.cladding_total > 0 %}
CLADDING SCHEDULE
{% for canopy in area.canopies %}
{% if canopy.has_cladding %}
Cladding below Item {{ canopy.reference_number }}, supplied and installed: {{ format_currency(canopy.cladding_price) }}
{% endif %}
{% endfor %}
SUB TOTAL: {{ format_currency(area.cladding_total) }}
{% endif %}
```

## 🔥 Fire Suppression Schedule Implementation

### Data Sources

- **Base Unit Price**: Read from Excel cells N12, N29, N46, etc.
- **Commissioning Price**: Read from Excel cell N193 (per fire suppression sheet)
- **Delivery Price**: Read from Excel cell N182 (per fire suppression sheet)
- **Tank Quantity**: Read from Excel cells C17, C34, C51, etc.

### Pricing Logic

1. **Base Price**: Individual unit price from N column
2. **Delivery Distribution**:
   - **Single Unit**: Gets full N182 delivery price
   - **Multiple Units**: N182 ÷ number of fire suppression units in area
3. **Unit Price**: Base Price + Delivery Amount
4. **Subtotal**: Sum of all unit prices (delivery included in each unit)

### Example Calculation

**Area with 2 fire suppression units:**

```
- Unit 1 Base Price: £1,690
- Unit 2 Base Price: £1,200
- Delivery (N182): £800

Delivery per unit: £800 ÷ 2 = £400

Unit 1 Price: £1,690 + £400 = £2,090
Unit 2 Price: £1,200 + £400 = £1,600

Fire Suppression Subtotal: £2,090 + £1,600 = £3,690
```

**Area with 1 fire suppression unit:**

```
- Unit 1 Base Price: £500
- Delivery (N182): £800

Unit 1 Price: £500 + £800 = £1,300

Fire Suppression Subtotal: £1,300
```

### Features

- **Conditional Display**: Fire suppression schedule only appears if area has fire suppression
- **Proportional Delivery**: Delivery costs split equally among units
- **Comprehensive Pricing**: Each unit price includes base cost plus delivery share
- **Area Totals**: Calculates total fire suppression cost per area

### Template Usage

```jinja2
{% if area.fire_suppression_total > 0 %}
ANSUL FIRE SUPPRESSION SCHEDULE
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_price > 0 %}
{{ canopy.reference_number }}    Ansul R102 system. Supplied, installed & commissioned.    {{ format_currency(canopy.fire_suppression_price) }}
{% endif %}
{% endfor %}
SUB TOTAL: {{ format_currency(area.fire_suppression_total) }}
{% endif %}
```

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
{% if area.uvc_price > 0 %}
UV-C CONTROL SCHEDULE
ITEM     UV-C CONTROL SCHEDULE                    PRICE
----------------------------------------------------------
UV-C     UV-C System, supplied and installed     {{ format_currency(area.uvc_price) }}
                                                  SUB TOTAL: {{ format_currency(area.uvc_price) }}
{% endif %}
```

## 📊 Comprehensive Pricing Structure

### Area-Level Totals

Each area provides the following pricing breakdowns:

```python
area_totals = {
    'canopy_total': 0,                    # Sum of individual canopy prices
    'cladding_total': 0,                  # Sum of individual cladding prices
    'fire_suppression_total': 0,          # Sum of individual fire suppression unit prices
    'delivery_installation_price': 0,     # Area delivery cost (P182)
    'commissioning_price': 0,             # Area commissioning cost (N193)
    'uvc_price': 0,                       # UV-C system price (N9 from EBOX sheet)
    'sdu_price': 0,                       # SDU system price (future implementation)
    'recoair_price': 0,                   # RecoAir system price (future implementation)
    'canopy_schedule_subtotal': 0,        # canopy_total + delivery + commissioning
    'area_subtotal': 0                    # Everything combined
}
```

### Project-Level Totals

```python
project_totals = {
    'total_canopy_price': 0,              # All canopy equipment
    'total_cladding_price': 0,            # All cladding
    'total_fire_suppression_price': 0,    # All fire suppression units
    'total_delivery_installation': 0,     # All delivery costs
    'total_commissioning': 0,             # All commissioning costs
    'total_uvc_price': 0,                 # All UV-C systems
    'total_sdu_price': 0,                 # All SDU systems (future)
    'total_recoair_price': 0,             # All RecoAir systems (future)
    'project_total': 0                    # Grand total
}
```

## 🎯 Display Logic

### Schedule Display Rules

1. **Canopy Schedule**: Always displayed if area has canopies
2. **Cladding Schedule**: Only displayed if `area.cladding_total > 0`
3. **Fire Suppression Schedule**: Only displayed if `area.fire_suppression_total > 0`
4. **UV-C Schedule**: Only displayed if `area.uvc_price > 0` or `area.options.uvc`

### Pricing Hierarchy

```
📋 CANOPY SCHEDULE
   - Individual canopy prices
   - Delivery & Installation
   - Commissioning
   SUB TOTAL: Canopy Schedule Subtotal

🔬 UV-C CONTROL SCHEDULE (if applicable)
   - UV-C system price
   SUB TOTAL: UV-C Total

🏗️ CLADDING SCHEDULE (if applicable)
   - Individual cladding prices
   SUB TOTAL: Cladding Total

🔥 FIRE SUPPRESSION SCHEDULE (if applicable)
   - Individual fire suppression unit prices (base + delivery share)
   SUB TOTAL: Fire Suppression Total

💰 AREA TOTAL: All schedules combined
```

## 🔧 Implementation Files

### Core Files Modified

- `src/utils/excel.py`: Excel reading with cladding and fire suppression pricing
- `src/utils/word.py`: Word document generation with all pricing schedules
- `test_pricing_schedule_demo.py`: Comprehensive demo with all schedules
- `jinja_template_examples.md`: Template examples for all schedules

### Test Scripts

- `test_cladding_pricing.py`: Cladding-specific testing
- `test_fire_suppression_pricing.py`: Fire suppression-specific testing
- `test_pricing_extraction.py`: General pricing extraction testing

## 📝 Template Examples

### Complete Area Schedule

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

## ✅ Key Benefits

1. **Comprehensive Pricing**: All costs properly allocated and displayed
2. **Conditional Display**: Only relevant schedules shown
3. **Accurate Calculations**: Fire suppression delivery costs properly split among units
4. **Professional Formatting**: Currency formatting and clear structure
5. **Template Flexibility**: Easy to customize display format
6. **Data Integrity**: Prices extracted from correct Excel cells
7. **Scalable Structure**: Supports multiple areas and canopies
8. **Area-Level Options**: UV-C, SDU, and RecoAir options supported at area level

## 🎉 Status: Complete

All pricing schedule functionality has been successfully implemented and tested:

- ✅ Cladding schedule with individual pricing
- ✅ Fire suppression schedule with delivery cost distribution
- ✅ UV-C Control Schedule with area-level pricing
- ✅ Comprehensive pricing totals including all options
- ✅ Word template integration with conditional display
- ✅ Excel reading and writing for all schedule types
