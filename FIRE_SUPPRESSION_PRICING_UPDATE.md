# Fire Suppression Pricing Update

## Summary of Changes

The fire suppression pricing logic has been updated to read delivery costs from **N182** instead of P182, and delivery costs are now **split evenly among units** and **included in each unit's individual price** rather than shown as a separate line item. **Commissioning costs are not split** - only delivery costs are distributed among units.

## Key Changes Made

### 1. Excel Reading (`src/utils/excel.py`)

- **Delivery Source**: Changed from `sheet['P182']` to `sheet['N182']`
- **Cost Distribution**: Only delivery cost split equally among all fire suppression units in the area
- **Unit Pricing**: Each unit price now includes: Base Price + Delivery Share (no commissioning share)
- **Removed**: Separate area-level fire suppression delivery tracking

### 2. Word Processing (`src/utils/word.py`)

- **Removed**: `fire_suppression_delivery_price` from area totals structure
- **Removed**: `total_fire_suppression_delivery` from project totals structure
- **Simplified**: Area and project subtotal calculations (delivery now included in unit prices)

### 3. Template Examples

- **Updated**: All Jinja template examples to remove separate delivery line items
- **Simplified**: Fire suppression schedules now show only unit prices and subtotal

### 4. Documentation

- **Updated**: `PRICING_SCHEDULES_SUMMARY.md` with new pricing logic
- **Updated**: `jinja_template_examples.md` with corrected templates
- **Updated**: Demo script with accurate sample data

## New Pricing Logic

### Before (P182 + Separate Line Item)

```
Unit 1: £1,690 (base) + £400 (commissioning) = £2,090
Unit 2: £1,200 (base) + £400 (commissioning) = £1,600
Delivery of Ansul Components: £800 (separate line)
Subtotal: £2,090 + £1,600 + £800 = £4,490
```

### After (N182 + Smart Distribution)

```
Example 1 - Multiple Units (2 units):
Unit 1: £1,690 (base) + £400 (delivery share: £800÷2) = £2,090
Unit 2: £1,200 (base) + £400 (delivery share: £800÷2) = £1,600
Subtotal: £2,090 + £1,600 = £3,690

Example 2 - Single Unit (1 unit):
Unit 1: £500 (base) + £800 (full delivery price) = £1,300
Subtotal: £1,300
```

## Template Usage

### Fire Suppression Schedule

```jinja2
{% if area.fire_suppression_total > 0 %}
🔥 ANSUL FIRE SUPPRESSION SCHEDULE
{% for canopy in area.canopies %}
{% if canopy.fire_suppression_price > 0 %}
{{ canopy.reference_number }}    Ansul R102 system. Supplied, installed & commissioned.    {{ format_currency(canopy.fire_suppression_price) }}
{% endif %}
{% endfor %}
SUB TOTAL: {{ format_currency(area.fire_suppression_total) }}
{% endif %}
```

## Benefits of This Approach

1. **Simplified Display**: No separate delivery line item needed
2. **Cleaner Templates**: Fewer conditional checks required
3. **Accurate Unit Pricing**: Each unit shows its complete cost including delivery
4. **Smart Distribution**: Single units get full delivery cost, multiple units split it fairly
5. **Single Source**: All fire suppression costs consolidated in unit prices

## Files Modified

- ✅ `src/utils/excel.py` - Updated delivery source and cost splitting
- ✅ `src/utils/word.py` - Removed separate delivery tracking
- ✅ `test_pricing_schedule_demo.py` - Updated sample data
- ✅ `jinja_template_examples.md` - Updated template examples
- ✅ `PRICING_SCHEDULES_SUMMARY.md` - Updated documentation

## Verification

The demo script confirms the changes work correctly:

- Fire suppression units now include delivery costs in individual prices
- Area and project totals calculate correctly
- Template context preparation works without errors
- All pricing schedules display properly

## Status: ✅ Complete

The fire suppression pricing update has been successfully implemented and tested. The system now reads delivery costs from N182, splits them evenly among units, and includes them in individual unit prices for a cleaner, more intuitive pricing display.
