# Corrected SDU Template for Word Document

## Complete SDU Specification Section

```jinja
{%if has_sdu%}
Services Distribution Unit
The services distribution units (SDU) are manufactured from 1.2mm thick, grade 304 stainless steel from an all folded and welded construction. The SDU is to be manufactured in suitably sized sections to suit site access requirements. The SDU is to be supplied with factory fitted mechanical and electrical services as listed below.
{% set processed_areas = [] %}
{% for sdu in sdu_areas %}
{% if sdu.level_area_combined not in processed_areas %}
{% set _ = processed_areas.append(sdu.level_area_combined) %}

Item {{ sdu.canopy_reference }}
{{ sdu.level_area_combined | upper}} SDU
    Electrical Services         Gas Services         Water Services
    Distribution Board    {{sdu.electrical_services.distribution_board}}        Gas manifold    {{ sdu.gas_services.gas_manifold }}        22mm CWS manifold    {{ sdu.water_services.cws_manifold_22mm }}
    Single Phase switched spur    {{ sdu.electrical_services.single_phase_switched_spur}}        15mm Connection    {{ sdu.gas_services.gas_connection_15mm }}        15mm CWS manifold    {{ sdu.water_services.cws_manifold_15mm }}
    Three Phase socket outlet    {{sdu.electrical_services.three_phase_socket_outlet }}        20mm Connection    {{ sdu.gas_services.gas_connection_20mm }}        22/15mm HWS manifold    {{ sdu.water_services.hws_manifold }}
    Switched socket outlet    {{ sdu.electrical_services.switched_socket_outlet }}        25mm Connection    {{ sdu.gas_services.gas_connection_25mm }}        15mm CWS/HWS outlet    {{ sdu.water_services.water_connection_15mm }}
    Emergency knock-off    {{sdu.electrical_services.emergency_knock_off}}        32mm Connection    {{ sdu.gas_services.gas_connection_32mm }}        22mm CWS/HWS outlet    {{ sdu.water_services.water_connection_22mm }}
    Ring main inc. 2no SSO    {{ sdu.electrical_services.ring_main_inc_2no_sso }}        Gas solenoid valve    {{ sdu.gas_services.gas_solenoid_valve }}        28mm CWS/HWS outlet    {{ sdu.water_services.water_connection_28mm }}
    
    SDU size {{sdu.sdu_length}}mm long x 300mm wide x 1200 high – 2no full / half height risers, 1no horizontal raceway.
Potrack {{sdu.potrack}} Salamander Support {{sdu.salamander_support}}
{% endif %}
{%endfor%}
All Halton Foodservice electrical installations shall be wired and tested to conform with the latest BS7671:2018 IET Wiring Regulations
{%endif%}
```

## Debug Template to Check SDU Data

If SDUs are still not appearing, add this debug section temporarily to see what data is available:

```jinja
<!-- DEBUG SDU DATA -->
{%if has_sdu%}
DEBUG: has_sdu = TRUE
Total SDU areas: {{total_sdu_areas}}
SDU areas list length: {{sdu_areas|length}}

{% for sdu in sdu_areas %}
SDU {{loop.index}}:
- Canopy Reference: {{sdu.canopy_reference}}
- Level/Area: {{sdu.level_area_combined}}
- Has SDU: {{sdu.has_sdu}}
- Electrical DB: {{sdu.electrical_services.distribution_board}}
{% endfor %}
{%else%}
DEBUG: has_sdu = FALSE
{%endif%}
<!-- END DEBUG -->
```

## Common Issues and Solutions

### 1. SDU Section Not Appearing At All
**Cause**: Missing `{%if has_sdu%}` wrapper
**Solution**: Wrap entire SDU section with the conditional check

### 2. Header Appears But No Data
**Cause**: `sdu_areas` list is empty
**Check**: 
- Ensure canopies have `options.sdu` set to `True` in Excel
- Verify `collect_sdu_data` is finding SDU canopies

### 3. Hardcoded Values Instead of Data
**Cause**: Template using literal values instead of variables
**Solution**: Replace all hardcoded values with template variables:
- `0` → `{{sdu.electrical_services.emergency_knock_off}}`
- `XXXXmm` → `{{sdu.sdu_length}}mm`
- `xxxxxxxxxxxxx` → `{{sdu.potrack}}`
- `xxxxxxxxxxxxxx` → `{{sdu.salamander_support}}`

## Complete Context Variables Available

For each SDU in `sdu_areas`:
```
sdu.level_name
sdu.area_name
sdu.canopy_reference
sdu.level_area_combined
sdu.has_sdu
sdu.sdu_price
sdu.sdu_length
sdu.potrack
sdu.salamander_support
sdu.pricing.final_carcass_price
sdu.pricing.final_electrical_price
sdu.pricing.live_site_test_price
sdu.pricing.has_live_test
sdu.pricing.total_price
sdu.electrical_services.distribution_board
sdu.electrical_services.single_phase_switched_spur
sdu.electrical_services.three_phase_socket_outlet
sdu.electrical_services.switched_socket_outlet
sdu.electrical_services.emergency_knock_off
sdu.electrical_services.ring_main_inc_2no_sso
sdu.gas_services.gas_manifold
sdu.gas_services.gas_connection_15mm
sdu.gas_services.gas_connection_20mm
sdu.gas_services.gas_connection_25mm
sdu.gas_services.gas_connection_32mm
sdu.gas_services.gas_solenoid_valve
sdu.water_services.cws_manifold_22mm
sdu.water_services.cws_manifold_15mm
sdu.water_services.hws_manifold
sdu.water_services.water_connection_15mm
sdu.water_services.water_connection_22mm
sdu.water_services.water_connection_28mm
```

## Pricing Section Template

```jinja
{% for area in pricing_totals.areas %}
{{ area.level_area_combined | upper}}
{# ... other pricing sections ... #}

{% if area.has_sdu%}
{% set area_sdus = sdu_areas | selectattr('level_area_combined', 'equalto', area.level_area_combined) | list %}
{% if area_sdus %}
{% set first_sdu = area_sdus[0] %}
    ITEM    SERVICE DISTRIBUTION UNIT SCHEDULE    PRICE
        Service Distribution Unit, {{first_sdu.sdu_length}}mm long, supplied & installed (carcass only).    {{format_currency(first_sdu.pricing.final_carcass_price)}}
        Electrical & mechanical services, supply and install    {{format_currency(first_sdu.pricing.final_electrical_price)}}
    {%tr if first_sdu.pricing.has_live_test %}        
        Extra over cost for separate 'Live Site Test'    {{format_currency(first_sdu.pricing.live_site_test_price)}}
    {%tr endif%}        
        Local shunt trip for electrical shut off (WIRING CONNECTION BY OTHERS)    
    SUB TOTAL    {{format_currency(area.sdu_subtotal)}}
{% endif %}
{%endif%}

{# ... continue with other sections ... #}
{% endfor %}
```