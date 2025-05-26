# Feature Flag Examples for Word Templates

This file shows examples of how to use feature flags in Word templates using Jinja2 syntax.

## Basic Feature Flag Usage

### Kitchen Extract System Example

```jinja2
{% if show_kitchen_extract_system %}
## Kitchen Extract System

| ITEM | DESCRIPTION |
|------|-------------|
| 5.01 | Galvanised steel duct, LV Class 'A' |
| 5.02 | Fire rated duct (see note below) |
| 5.03 | Stainless steel duct |
| 5.04 | High velocity extract cowl |
| 5.05 | Extract fan |
| 5.06 | Extract fan controls package |
| 5.07 | Gas Interlocking System |
| 5.08 | Gas solenoid valve |
| 5.09 | Volume control damper |
| 5.10 | Melinex lined attenuation |
| 5.11 | Access doors |
| 5.12 | Exhaust louvre |
| 5.13 | Powder coat duct to RAL(*) |
| 5.14 | PST ## Pollustop Unit |
| 5.15 | Dishwasher extract fan / controls |
| 5.16 | |

**Note**: Fire rated ductwork, non-coated, galvanised, to BS 476 Part 24 1987 for the following criteria:-
**Stability** 120 minutes, **Integrity** 120 minutes, **Insulation** Nil
{% endif %}
```

### Kitchen Make-Up Air System Example

```jinja2
{% if show_kitchen_makeup_air_system %}
## Kitchen Make-Up Air System

| ITEM | DESCRIPTION |
|------|-------------|
| 6.01 | Galvanised steel duct, LV Class 'A' |
| 6.02 | Fire damper |
| 6.03 | Stainless steel duct |
| 6.04 | Intake cowl |
| 6.05 | Make-up air fan |
| 6.06 | Make-up air fan controls package |
| 6.07 | Ductwork insulation (lagging) |
| 6.08 | Filter plenum inc. set of filters |
| 6.09 | Heater battery – LPHW / electric |
| 6.10 | Volume control damper |
| 6.11 | Attenuation |
| 6.12 | Access doors |
| 6.13 | Make-up air louvre |
| 6.14 | |
| 6.15 | |
| 6.16 | |
| 6.17 | |
| 6.18 | |
{% endif %}
```

### M.A.R.V.E.L. System Example

```jinja2
{% if show_marvel_system %}
## M.A.R.V.E.L. System (DCKV)

The Halton M.A.R.V.E.L. system is a **demand based** control system specifically designed for kitchen canopies and ventilated ceilings. The system has the ability to:

• Detect and identify the status of the cooking appliance under the canopy i.e. 'Off', 'Heating Up' or 'Cooking' via overhead mounted infrared sensors.
• Adjust the airflow from the exhaust ducts via in-line duct mounted motorised volume control dampers.
• Continuously regulate the kitchen extract airflow and system pressure. The associated supply air fan(s) can also be regulated to guarantee a suitable airflow balance within the kitchen.
• It is a flexible system that can be reprogrammed to suit changes in the general kitchen appliance layout.

It is imperative that a full M.A.R.V.E.L. System (extract & supply systems) is quoted for and designed using the most recently issued kitchen layout and site ductwork layout drawings.

**Important Notes:**
• We have not at this stage, considered the control method of the supply air system as there is insufficient information available.
• We have not allowed to run cables of any sort between the canopies/ ventilated ceiling controls and the extract/supply fan VFD's
{% endif %}
```

### Cyclocell Cassette Ceiling Example

```jinja2
{% if show_cyclocell_cassette_ceiling %}
## Cyclocell Cassette Ceiling System

The Cyclocell Cassette Ceiling System is designed around a standard 600mm x 600mm concealed suspended ceiling grill. Each cassette has a 600mm x 600mm outer frame, which is supported by the grid framework. The fixed rigid framework grid provides the platform for which the cassette variants can be installed to suit the kitchen appliance layout in the kitchen workspace below.

| ITEM NUMBER | VENTILATED CEILING LOCATION | CEILING COVERAGE (m²) |
|-------------|----------------------------|----------------------|
| 8.01 | Ventilated ceiling to Main Kitchen – see drawing XXXX | 0 m² |
| 8.02 | | |
| 8.03 | | |

**Important Note**: All ductwork above the ceiling, up to the perimeter of the kitchen, has been included for within the quotation.
**Important Note**: No allowance has been made for any of the kitchen extract or supply air ductwork within the ceiling void above the ventilated ceiling.
{% endif %}
```

### Reactaway Unit Example

```jinja2
{% if show_reactaway_unit %}
## Reactaway Unit

A "Reactaway Unit" is an in-line duct mounted UV-C filtration module, manufactured from catering grade stainless steel complete with flanged inlet and outlet connection spigots. The Reactaway unit is used as an alternative to Halton's canopy mounted UV-C system, or as a retro-fitted product to reduce grease deposits and odour emissions from the kitchen extraction system prior to termination to atmosphere.

The unit is supplied with access doors, shut-off safety features, removable UV-C cassettes, ballast box and a unit mounted control package. Remote controls are available on request.

| ITEM REF. | MODEL REF. | DIMENSIONS (mm) | EXT.VOL. | P. DROP | NOTES | WEIGHT | LOCATION |
|-----------|------------|-----------------|----------|---------|-------|--------|----------|
|           |            | L | W | D | (m³/s) | (Pa) |       | (Kgs) | INT / EXT |
| 9.01 | PST-00 | 0 | 0 | 0 | | | | | |
{% endif %}
```

### Multiple System Conditional Logic

```jinja2
{% if show_kitchen_extract_system or show_kitchen_makeup_air_system %}
## Kitchen Ventilation Systems

{% if show_kitchen_extract_system %}
### Extract System
[Extract system content here]
{% endif %}

{% if show_kitchen_makeup_air_system %}
### Make-Up Air System
[Make-up air system content here]
{% endif %}

{% endif %}
```

### Pricing Conditional Logic

```jinja2
{% if show_marvel_system and pricing_totals.marvel_total > 0 %}
| M.A.R.V.E.L. System | {{ format_currency(pricing_totals.marvel_total) }} |
{% endif %}

{% if show_reactaway_unit and pricing_totals.reactaway_total > 0 %}
| Reactaway Unit | {{ format_currency(pricing_totals.reactaway_total) }} |
{% endif %}

{% if show_pollustop_unit and pricing_totals.pollustop_total > 0 %}
| Pollustop Unit | {{ format_currency(pricing_totals.pollustop_total) }} |
{% endif %}
```

## Available Feature Flag Variables in Templates

The following variables are available in Word templates:

- `show_kitchen_extract_system` - Kitchen Extract System
- `show_kitchen_makeup_air_system` - Kitchen Make-Up Air System
- `show_marvel_system` - M.A.R.V.E.L. System (DCKV)
- `show_cyclocell_cassette_ceiling` - Cyclocell Cassette Ceiling
- `show_reactaway_unit` - Reactaway Unit
- `show_dishwasher_extract` - Dishwasher Extract
- `show_gas_interlocking` - Gas Interlocking
- `show_pollustop_unit` - Pollustop Unit

## Best Practices

1. **Always use conditional blocks** for systems that might not be enabled
2. **Check for data existence** as well as feature flags when showing pricing
3. **Provide fallback content** when systems are disabled
4. **Use descriptive comments** in templates to explain conditional logic
5. **Test templates** with both enabled and disabled flags

## Testing Templates

To test templates with different feature flag combinations:

1. Enable/disable flags in `src/config/constants.py`
2. Generate test documents
3. Verify content appears/disappears correctly
4. Check that pricing calculations are accurate
5. Ensure no broken references or empty sections
