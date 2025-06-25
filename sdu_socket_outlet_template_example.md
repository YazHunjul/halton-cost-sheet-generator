# SDU Socket Outlet Template Examples

This document shows how to access and display single-phase and three-phase socket outlet counts for SDU systems in Word templates.

## 📊 Data Structure

The SDU socket outlet counts are available in the `sdu_areas` variable. Each SDU area contains:

```jinja2
{
  'level_name': 'Ground Floor',
  'area_name': 'Kitchen Area',
  'level_area_combined': 'Ground Floor - Kitchen Area',
  'electrical_services': {
    'single_phase_switched_spur': 5,    # Total 1-PH socket outlets from D40-D47 + D49-D56
    'three_phase_socket_outlet': 3,     # Total 3-PH socket outlets from D40-D47 + D49-D56
    'distribution_board': 1,
    'switched_socket_outlet': 0,
    'emergency_knock_off': 0,
    'ring_main_inc_2no_sso': 0
  },
  // ... other services and pricing data
}
```

## 🏗️ Basic Socket Outlet Display

### Simple Summary

```jinja2
{% if sdu_areas %}
## SDU Socket Outlet Summary

{% for sdu in sdu_areas %}
### {{ sdu.level_area_combined }}
- **Single Phase Socket Outlets**: {{ sdu.electrical_services.single_phase_switched_spur }}
- **Three Phase Socket Outlets**: {{ sdu.electrical_services.three_phase_socket_outlet }}
{% endfor %}

**Total Single Phase**: {{ sdu_areas|sum(attribute='electrical_services.single_phase_switched_spur') }}
**Total Three Phase**: {{ sdu_areas|sum(attribute='electrical_services.three_phase_socket_outlet') }}
{% endif %}
```

### Table Format

```jinja2
{% if sdu_areas %}
## SDU Socket Outlet Schedule

| AREA | SINGLE PHASE | THREE PHASE |
|------|-------------|-------------|
{% for sdu in sdu_areas %}
| {{ sdu.level_area_combined }} | {{ sdu.electrical_services.single_phase_switched_spur }} | {{ sdu.electrical_services.three_phase_socket_outlet }} |
{% endfor %}
| **TOTAL** | **{{ sdu_areas|sum(attribute='electrical_services.single_phase_switched_spur') }}** | **{{ sdu_areas|sum(attribute='electrical_services.three_phase_socket_outlet') }}** |
{% endif %}
```

## 🔌 Detailed Electrical Services Schedule

### Complete Electrical Services Table

```jinja2
{% for sdu in sdu_areas %}
## {{ sdu.level_area_combined | upper }} - ELECTRICAL SERVICES

| ITEM | ELECTRICAL SERVICES | QUANTITY |
|------|-------------------|----------|
{% if sdu.electrical_services.distribution_board > 0 %}
| E1 | Distribution Board | {{ sdu.electrical_services.distribution_board }} |
{% endif %}
{% if sdu.electrical_services.single_phase_switched_spur > 0 %}
| E2 | Single Phase Socket Outlets | {{ sdu.electrical_services.single_phase_switched_spur }} |
{% endif %}
{% if sdu.electrical_services.three_phase_socket_outlet > 0 %}
| E3 | Three Phase Socket Outlets | {{ sdu.electrical_services.three_phase_socket_outlet }} |
{% endif %}
{% if sdu.electrical_services.switched_socket_outlet > 0 %}
| E4 | Switched Socket Outlets | {{ sdu.electrical_services.switched_socket_outlet }} |
{% endif %}
{% if sdu.electrical_services.emergency_knock_off > 0 %}
| E5 | Emergency Knock-Off | {{ sdu.electrical_services.emergency_knock_off }} |
{% endif %}
{% if sdu.electrical_services.ring_main_inc_2no_sso > 0 %}
| E6 | Ring Main inc. 2no SSO | {{ sdu.electrical_services.ring_main_inc_2no_sso }} |
{% endif %}

{% endfor %}
```

## 🏢 Multi-Area Projects

### Socket Outlet Summary by Level

```jinja2
{% if sdu_areas %}
## SDU Socket Outlet Summary by Level

{% set levels_dict = {} %}
{% for sdu in sdu_areas %}
  {% set level = sdu.level_name %}
  {% if level not in levels_dict %}
    {% set _ = levels_dict.update({level: {'areas': [], 'total_1ph': 0, 'total_3ph': 0}}) %}
  {% endif %}
  {% set _ = levels_dict[level]['areas'].append(sdu) %}
  {% set _ = levels_dict[level].update({
    'total_1ph': levels_dict[level]['total_1ph'] + sdu.electrical_services.single_phase_switched_spur,
    'total_3ph': levels_dict[level]['total_3ph'] + sdu.electrical_services.three_phase_socket_outlet
  }) %}
{% endfor %}

{% for level_name, level_data in levels_dict.items() %}
### {{ level_name }}

| AREA | 1-PH OUTLETS | 3-PH OUTLETS |
|------|-------------|-------------|
{% for sdu in level_data.areas %}
| {{ sdu.area_name }} | {{ sdu.electrical_services.single_phase_switched_spur }} | {{ sdu.electrical_services.three_phase_socket_outlet }} |
{% endfor %}
| **{{ level_name }} TOTAL** | **{{ level_data.total_1ph }}** | **{{ level_data.total_3ph }}** |

{% endfor %}

**PROJECT TOTALS:**
- **Single Phase Outlets**: {{ sdu_areas|sum(attribute='electrical_services.single_phase_switched_spur') }}
- **Three Phase Outlets**: {{ sdu_areas|sum(attribute='electrical_services.three_phase_socket_outlet') }}
{% endif %}
```

## 🔍 Conditional Display

### Only Show Areas with Socket Outlets

```jinja2
{% if sdu_areas %}
## SDU Areas with Socket Outlets

{% for sdu in sdu_areas %}
  {% set has_outlets = sdu.electrical_services.single_phase_switched_spur > 0 or sdu.electrical_services.three_phase_socket_outlet > 0 %}
  {% if has_outlets %}
### {{ sdu.level_area_combined }}

**Socket Outlets Required:**
    {% if sdu.electrical_services.single_phase_switched_spur > 0 %}
- {{ sdu.electrical_services.single_phase_switched_spur }}x Single Phase ISO/Outlets
    {% endif %}
    {% if sdu.electrical_services.three_phase_socket_outlet > 0 %}
- {{ sdu.electrical_services.three_phase_socket_outlet }}x Three Phase ISO/Outlets
    {% endif %}

  {% endif %}
{% endfor %}
{% endif %}
```

## 📋 Integration with Main SDU Schedule

### Combined SDU Services Schedule

```jinja2
{% for sdu in sdu_areas %}
## {{ sdu.level_area_combined | upper }} - SDU SERVICES SCHEDULE

### Electrical Services
{% if sdu.electrical_services.single_phase_switched_spur > 0 or sdu.electrical_services.three_phase_socket_outlet > 0 %}
| ITEM | ELECTRICAL SERVICES | QTY |
|------|-------------------|-----|
{% if sdu.electrical_services.single_phase_switched_spur > 0 %}
| E1 | Single Phase Socket Outlets (16A/32A) | {{ sdu.electrical_services.single_phase_switched_spur }} |
{% endif %}
{% if sdu.electrical_services.three_phase_socket_outlet > 0 %}
| E2 | Three Phase Socket Outlets (16A/32A/63A/125A) | {{ sdu.electrical_services.three_phase_socket_outlet }} |
{% endif %}
{% endif %}

### Gas Services
{% if sdu.gas_services.gas_manifold > 0 or sdu.gas_services.gas_solenoid_valve > 0 %}
| ITEM | GAS SERVICES | QTY |
|------|-------------|-----|
{% if sdu.gas_services.gas_manifold > 0 %}
| G1 | Gas Manifold | {{ sdu.gas_services.gas_manifold }} |
{% endif %}
{% if sdu.gas_services.gas_solenoid_valve > 0 %}
| G2 | Gas Solenoid Valve | {{ sdu.gas_services.gas_solenoid_valve }} |
{% endif %}
{% endif %}

### Water Services
{% if sdu.water_services.cws_manifold_22mm > 0 or sdu.water_services.hws_manifold > 0 %}
| ITEM | WATER SERVICES | QTY |
|------|---------------|-----|
{% if sdu.water_services.cws_manifold_22mm > 0 %}
| W1 | CWS Manifold 22mm | {{ sdu.water_services.cws_manifold_22mm }} |
{% endif %}
{% if sdu.water_services.hws_manifold > 0 %}
| W2 | HWS Manifold | {{ sdu.water_services.hws_manifold }} |
{% endif %}
{% endif %}

**{{ sdu.level_area_combined }} SDU TOTAL**: {{ format_currency(sdu.sdu_price) }}

{% endfor %}
```

## 📊 Project Totals

### Socket Outlet Project Summary

```jinja2
{% if sdu_areas %}
---

## SDU PROJECT SUMMARY

**Total SDU Areas**: {{ sdu_areas|length }}

**Socket Outlet Totals:**
- Single Phase Outlets: {{ sdu_areas|sum(attribute='electrical_services.single_phase_switched_spur') }}
- Three Phase Outlets: {{ sdu_areas|sum(attribute='electrical_services.three_phase_socket_outlet') }}
- Total Outlets: {{ (sdu_areas|sum(attribute='electrical_services.single_phase_switched_spur')) + (sdu_areas|sum(attribute='electrical_services.three_phase_socket_outlet')) }}

**SDU Investment Total**: {{ format_currency(sdu_areas|sum(attribute='sdu_price')) }}
{% endif %}
```

## 🔧 Usage Notes

### Key Points:

1. **Data Source**: Socket outlet counts come from Excel cells D40-D47 and D49-D56 based on dropdown selections
2. **Automatic Counting**: System automatically counts 1-PH and 3-PH variants and sums their quantities
3. **Template Access**: Use `sdu.electrical_services.single_phase_switched_spur` and `sdu.electrical_services.three_phase_socket_outlet`
4. **Conditional Display**: Only show sections when outlets are actually required
5. **Totaling**: Use Jinja filters to sum across multiple areas

### Example Values:

- **Single Phase**: Includes 16 AMP 1-PH and 32 AMP 1-PH outlets (both MCB and NO MCB)
- **Three Phase**: Includes 16 AMP 3-PH, 32 AMP 3-PH, 63 AMP 3-PH, and 125 AMP 3-PH outlets (both MCB and NO MCB)
