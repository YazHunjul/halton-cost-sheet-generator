# RecoAir Jinja Template Code

This document contains the exact Jinja code needed for the RecoAir Word template to properly iterate through multiple levels and areas.

## 1. Unit Specifications Table (Technical Details)

```jinja2
{% for schedule in recoair_pricing_schedules %}
{% for unit in schedule.units %}
| {{ unit.reference_number }} | {{ unit.model }} | {{ unit.length }} x {{ unit.width }} x {{ unit.height }} | {{ unit.extract_volume }} | {{ unit.p_drop }} | {{ unit.motor }} | {{ unit.weight }} | {{ unit.location }} |
{% endfor %}
{% endfor %}
```

### Alternative: Using Both Pricing Items and Unit Specs

```jinja2
{% for schedule in recoair_pricing_schedules %}
{% for unit in schedule.units %}
| {{ unit.reference_number }} | {{ unit.model }} | {{ unit.length }}\|{{ unit.width }}\|{{ unit.height }} | {{ unit.extract_volume }} | {{ unit.p_drop }} | {{ unit.motor }} | {{ unit.weight }} | {{ unit.location }} |
{% endfor %}
{% endfor %}
```

## 2. Individual Level Pricing Tables

```jinja2
{% for schedule in recoair_pricing_schedules %}

### {{ schedule.level_area_combined | upper }}

**Unit Schedule:**
| ITEM | UNIT SCHEDULE | PRICE |
|------|---------------|-------|
{% for item in schedule.recoair_items %}
| {{ item.reference_number }} | {{ item.model }}, Ex-Works | {{ format_currency(item.price) }} |
{% endfor %}
| | Delivery and Installation | {{ format_currency(schedule.delivery_installation_price) }} |
| | Commissioning | {{ format_currency(schedule.commissioning_price) }} |
| SUB TOTAL | SUB TOTAL | {{ format_currency(schedule.area_subtotal) }} |

{% if schedule.has_flat_pack %}
**Additional Items:**
| ITEM | ADDITIONAL ITEMS | PRICE |
|------|------------------|-------|
{% for item in schedule.recoair_items %}
| {{ item.reference_number }} | Flat Pack Reassemble On Site | {{ format_currency(schedule.flat_pack_price) }} |
{% endfor %}
| SUB TOTAL | SUB TOTAL | {{ format_currency(schedule.flat_pack_price) }} |
{% endif %}

**{{ schedule.level_area_combined | upper }} (EXCLUDING VAT): {{ format_currency(schedule.area_total_with_flat_pack) }}**

---

{% endfor %}
```

## 3. Job Totals Section

```jinja2
**TOTAL (EXCLUDING VAT): {{ format_currency(recoair_job_totals.job_total) }}**

### Project Summary:
- Total Areas: {{ recoair_job_totals.total_areas }}
- Total Units: {{ recoair_job_totals.total_units }}
- Total Units Price: {{ format_currency(recoair_job_totals.total_units_price) }}
- Total Delivery: {{ format_currency(recoair_job_totals.total_delivery_price) }}
- Total Commissioning: {{ format_currency(recoair_job_totals.total_commissioning_price) }}
- Total Flat Pack: {{ format_currency(recoair_job_totals.total_flat_pack_price) }}
```

## 4. Complete Template Structure

```jinja2
# RecoAir Quotation Document

## Project Information
- **Project:** {{ project_name }}
- **Customer:** {{ customer }}
- **Date:** {{ date }}
- **Reference:** {{ reference_variable }}

## Unit Specifications

| ITEM REF. | MODEL | DIMENSIONS (mm) | EXT. VOL | P. DROP | MOTOR KW/PH | WEIGHT | LOCATION |
|-----------|-------|-----------------|----------|---------|-------------|--------|----------|
|           |       | L \| W \| D     | (m³/s)   | (Pa)    | (KW/PH)     | (Kgs)  | INT/EXT  |
{% for schedule in recoair_pricing_schedules %}
{% for unit in schedule.units %}
| {{ unit.reference_number }} | {{ unit.model }} | {{ unit.length }} \| {{ unit.width }} \| {{ unit.height }} | {{ unit.extract_volume }} | {{ unit.p_drop }} | {{ unit.motor }} | {{ unit.weight }} | {{ unit.location }} |
{% endfor %}
{% endfor %}

## Pricing Schedules

{% for schedule in recoair_pricing_schedules %}

### {{ schedule.level_area_combined | upper }}

#### Unit Schedule
| ITEM | UNIT SCHEDULE | PRICE |
|------|---------------|-------|
{% for item in schedule.recoair_items %}
| {{ item.reference_number }} | {{ item.model }}, Ex-Works | {{ format_currency(item.price) }} |
{% endfor %}
| | Delivery and Installation | {{ format_currency(schedule.delivery_installation_price) }} |
| | Commissioning | {{ format_currency(schedule.commissioning_price) }} |
| **SUB TOTAL** | **SUB TOTAL** | **{{ format_currency(schedule.area_subtotal) }}** |

{% if schedule.has_flat_pack %}
#### Additional Items
| ITEM | ADDITIONAL ITEMS | PRICE |
|------|------------------|-------|
{% for item in schedule.recoair_items %}
| {{ item.reference_number }} | {{ schedule.flat_pack_description }} | {{ format_currency(schedule.flat_pack_price) }} |
{% endfor %}
| **SUB TOTAL** | **SUB TOTAL** | **{{ format_currency(schedule.flat_pack_price) }}** |
{% endif %}

#### {{ schedule.level_area_combined | upper }} (EXCLUDING VAT)
**Total: {{ format_currency(schedule.area_total_with_flat_pack) }}**

---

{% endfor %}

## Project Total

**TOTAL (EXCLUDING VAT): {{ format_currency(recoair_job_totals.job_total) }}**
```

## 5. Alternative: Using Levels Structure

If you prefer to iterate through the levels structure instead:

```jinja2
{% for level in levels %}
{% for area in level.areas %}
{% if area.options.recoair %}

### {{ level.level_name | upper }} - {{ area.name | upper }}

#### RecoAir Units
{% for unit in area.recoair_units %}
| {{ unit.item_reference }} | {{ unit.model }} | {{ unit.width }} x {{ unit.length }} x {{ unit.height }}mm | {{ unit.extract_volume }}m³/s | {{ unit.p_drop }}Pa | {{ unit.motor }}kW | {{ unit.weight }}kg | {{ unit.location }} |
{% endfor %}

#### Pricing
- **Unit Price:** {{ format_currency(area.recoair_price) }}
- **Commissioning:** {{ format_currency(area.recoair_commissioning_price) }}

---

{% endif %}
{% endfor %}
{% endfor %}
```

## 6. Key Variables Available

### RecoAir Pricing Schedules (`recoair_pricing_schedules`)

Each schedule contains:

- `level_name` - Level name (e.g., "Level 1", "Second Level")
- `area_name` - Area name (e.g., "Main Kitchen")
- `level_area_combined` - Combined name (e.g., "Level 1 - Main Kitchen")
- `recoair_items` - List of RecoAir units
- `units_total` - Total price of units
- `delivery_installation_price` - Delivery cost
- `commissioning_price` - Commissioning cost
- `flat_pack_price` - Flat pack cost
- `area_subtotal` - Subtotal excluding flat pack
- `area_total_with_flat_pack` - Total including flat pack
- `unit_count` - Number of units
- `has_flat_pack` - Boolean for flat pack availability

### RecoAir Job Totals (`recoair_job_totals`)

- `total_units_price` - Sum of all unit prices
- `total_delivery_price` - Sum of all delivery costs
- `total_commissioning_price` - Sum of all commissioning costs
- `total_flat_pack_price` - Sum of all flat pack costs
- `job_total` - Grand total
- `total_areas` - Number of areas
- `total_units` - Number of units

### Individual RecoAir Items (`recoair_items`)

Each item in `recoair_items` contains basic pricing information:

- `reference_number` - Item reference (e.g., "1.01", "2.01")
- `model` - RecoAir model (e.g., "RAH3.5", "RAH2.0")
- `price` - Unit price
- `delivery_price` - Delivery price for this unit

### Full Unit Specifications (`units`)

Each unit in `schedule.units` contains complete technical specifications:

**Basic Information:**

- `reference_number` - Item reference (e.g., "1.01", "2.01")
- `model` - Transformed RecoAir model (e.g., "RAH3.5", "RAH2.0")
- `model_original` - Original model from Excel
- `price` - Final unit price (including N29 addition)
- `delivery_price` - Delivery price for this unit
- `quantity` - Quantity selected

**Dimensions (mm):**

- `length` - Unit length in mm
- `width` - Unit width in mm
- `height` - Unit height in mm

**Technical Specifications:**

- `extract_volume` - Extract volume in m³/s
- `extract_volume_raw` - Raw extract volume string from Excel
- `p_drop` - Pressure drop in Pa
- `motor` - Motor power in kW/PH
- `weight` - Unit weight in kg
- `location` - Installation location ("INTERNAL" or "EXTERNAL")

**Pricing Details:**

- `base_unit_price` - Base unit price from Excel
- `n29_addition` - Additional cost from N29 cell

## 7. Debugging Template

To debug what data is available, you can temporarily add:

```jinja2
<!-- DEBUG: Available Data -->
<!-- Levels: {{ levels|length }} -->
<!-- RecoAir Schedules: {{ recoair_pricing_schedules|length }} -->
<!-- RecoAir Areas: {{ recoair_areas|length }} -->

{% for schedule in recoair_pricing_schedules %}
<!-- Schedule {{ loop.index }}: {{ schedule.level_name }} - {{ schedule.area_name }} ({{ schedule.unit_count }} units) -->
{% endfor %}
```

## 8. Word Table Format

For Word document tables, use this structure:

```jinja2
{% for schedule in recoair_pricing_schedules %}

{{ schedule.level_area_combined | upper }}

{% for item in schedule.recoair_items %}
{{ item.reference_number }}	{{ item.model }}, Ex-Works	{{ format_currency(item.price) }}
{% endfor %}
	Delivery and Installation	{{ format_currency(schedule.delivery_installation_price) }}
	Commissioning	{{ format_currency(schedule.commissioning_price) }}
SUB TOTAL	SUB TOTAL	{{ format_currency(schedule.area_subtotal) }}

{% if schedule.has_flat_pack %}
{% for item in schedule.recoair_items %}
{{ item.reference_number }}	Flat Pack Reassemble On Site	{{ format_currency(schedule.flat_pack_price) }}
{% endfor %}
SUB TOTAL	SUB TOTAL	{{ format_currency(schedule.flat_pack_price) }}
{% endif %}

£	{{ schedule.level_area_combined | upper }} (EXCLUDING VAT)	{{ format_currency(schedule.area_total_with_flat_pack) }}

{% endfor %}

£	TOTAL (EXCLUDING VAT)	{{ format_currency(recoair_job_totals.job_total) }}
```

## 9. Complete Unit Specifications Table

For the technical specifications table with all unit details:

```jinja2
{% for schedule in recoair_pricing_schedules %}

{{ schedule.level_area_combined | upper }}

{% for unit in schedule.units %}
{{ unit.reference_number }}	{{ unit.model }}	{{ unit.length }}x{{ unit.width }}x{{ unit.height }}	{{ unit.extract_volume }}	{{ unit.p_drop }}	{{ unit.motor }}	{{ unit.weight }}	{{ unit.location }}
{% endfor %}

{% endfor %}
```

## 10. Mixed Pricing and Specifications

To show both pricing and specifications in one table:

```jinja2
{% for schedule in recoair_pricing_schedules %}

{{ schedule.level_area_combined | upper }}

{% for unit in schedule.units %}
{{ unit.reference_number }}	{{ unit.model }} ({{ unit.length }}x{{ unit.width }}x{{ unit.height }}mm, {{ unit.extract_volume }}m³/s)	{{ format_currency(unit.price) }}
{% endfor %}
	Delivery and Installation	{{ format_currency(schedule.delivery_installation_price) }}
	Commissioning	{{ format_currency(schedule.commissioning_price) }}
SUB TOTAL	SUB TOTAL	{{ format_currency(schedule.area_subtotal) }}

{% endfor %}
```

This code will properly iterate through all levels and areas, generating separate pricing tables for each level with full unit specifications as confirmed by our debugging.
