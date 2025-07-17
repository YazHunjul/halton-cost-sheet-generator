# SDU Template - Show All SDUs Individually

## Fixed SDU Specification Section

This template will show ALL SDUs, not just one per area. Each SDU (canopy) will be listed separately.

```jinja
{%if has_sdu%}
Services Distribution Unit
The services distribution units (SDU) are manufactured from 1.2mm thick, grade 304 stainless steel from an all folded and welded construction. The SDU is to be manufactured in suitably sized sections to suit site access requirements. The SDU is to be supplied with factory fitted mechanical and electrical services as listed below.
{% for sdu in sdu_areas %}

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
{%endfor%}
All Halton Foodservice electrical installations shall be wired and tested to conform with the latest BS7671:2018 IET Wiring Regulations
{%endif%}
```

## Key Changes:

1. **Removed the `processed_areas` logic** - This was causing only the first SDU per area to be shown
2. **Removed the conditional check** `{% if sdu.level_area_combined not in processed_areas %}`
3. **Each SDU is now shown individually** - If you have 3 canopies with SDU in the same area, all 3 will appear

## Alternative: Group by Area but Show All Canopies

If you want to group by area but still show all SDUs, use this template:

```jinja
{%if has_sdu%}
Services Distribution Unit
The services distribution units (SDU) are manufactured from 1.2mm thick, grade 304 stainless steel from an all folded and welded construction. The SDU is to be manufactured in suitably sized sections to suit site access requirements. The SDU is to be supplied with factory fitted mechanical and electrical services as listed below.
{% set areas_dict = {} %}
{% for sdu in sdu_areas %}
    {% if sdu.level_area_combined not in areas_dict %}
        {% set _ = areas_dict.update({sdu.level_area_combined: []}) %}
    {% endif %}
    {% set _ = areas_dict[sdu.level_area_combined].append(sdu) %}
{% endfor %}

{% for area_name, area_sdus in areas_dict.items() %}
{{ area_name | upper }}
{% for sdu in area_sdus %}
Item {{ sdu.canopy_reference }}
    Electrical Services         Gas Services         Water Services
    Distribution Board    {{sdu.electrical_services.distribution_board}}        Gas manifold    {{ sdu.gas_services.gas_manifold }}        22mm CWS manifold    {{ sdu.water_services.cws_manifold_22mm }}
    Single Phase switched spur    {{ sdu.electrical_services.single_phase_switched_spur}}        15mm Connection    {{ sdu.gas_services.gas_connection_15mm }}        15mm CWS manifold    {{ sdu.water_services.cws_manifold_15mm }}
    Three Phase socket outlet    {{sdu.electrical_services.three_phase_socket_outlet }}        20mm Connection    {{ sdu.gas_services.gas_connection_20mm }}        22/15mm HWS manifold    {{ sdu.water_services.hws_manifold }}
    Switched socket outlet    {{ sdu.electrical_services.switched_socket_outlet }}        25mm Connection    {{ sdu.gas_services.gas_connection_25mm }}        15mm CWS/HWS outlet    {{ sdu.water_services.water_connection_15mm }}
    Emergency knock-off    {{sdu.electrical_services.emergency_knock_off}}        32mm Connection    {{ sdu.gas_services.gas_connection_32mm }}        22mm CWS/HWS outlet    {{ sdu.water_services.water_connection_22mm }}
    Ring main inc. 2no SSO    {{ sdu.electrical_services.ring_main_inc_2no_sso }}        Gas solenoid valve    {{ sdu.gas_services.gas_solenoid_valve }}        28mm CWS/HWS outlet    {{ sdu.water_services.water_connection_28mm }}
    
    SDU size {{sdu.sdu_length}}mm long x 300mm wide x 1200 high – 2no full / half height risers, 1no horizontal raceway.
Potrack {{sdu.potrack}} Salamander Support {{sdu.salamander_support}}

{% endfor %}
{% endfor %}
All Halton Foodservice electrical installations shall be wired and tested to conform with the latest BS7671:2018 IET Wiring Regulations
{%endif%}
```

## Updated Pricing Section

For the pricing section, you'll also need to show all SDUs:

```jinja
{% for area in pricing_totals.areas %}
{{ area.level_area_combined | upper}}
{# ... other pricing sections ... #}

{% if area.has_sdu%}
{% set area_sdus = sdu_areas | selectattr('level_area_combined', 'equalto', area.level_area_combined) | list %}
{% if area_sdus %}
    ITEM    SERVICE DISTRIBUTION UNIT SCHEDULE    PRICE
{% for sdu in area_sdus %}
    {{sdu.canopy_reference}}    Service Distribution Unit, {{sdu.sdu_length}}mm long, supplied & installed (carcass only).    {{format_currency(sdu.pricing.final_carcass_price)}}
        Electrical & mechanical services, supply and install    {{format_currency(sdu.pricing.final_electrical_price)}}
    {%tr if sdu.pricing.has_live_test %}        
        Extra over cost for separate 'Live Site Test'    {{format_currency(sdu.pricing.live_site_test_price)}}
    {%tr endif%}        
        Local shunt trip for electrical shut off (WIRING CONNECTION BY OTHERS)    
{% endfor %}
    SUB TOTAL    {{format_currency(area.sdu_subtotal)}}
{% endif %}
{%endif%}

{# ... continue with other sections ... #}
{% endfor %}
```

## Why This Fixes the Issue:

1. **Original problem**: The `processed_areas` list was tracking areas that had been shown, and skipping any SDU from an area that was already processed
2. **Solution**: Remove this check so every SDU is shown
3. **Result**: If you have 5 canopies with SDU across 2 areas, all 5 will now appear in the document

## Example Output:

If you have:
- Canopy C001 with SDU in "Level 1 - Kitchen"
- Canopy C002 with SDU in "Level 1 - Kitchen"
- Canopy C003 with SDU in "Level 2 - Prep Area"

The document will now show:

```
Item C001
LEVEL 1 - KITCHEN SDU
[services table]

Item C002
LEVEL 1 - KITCHEN SDU
[services table]

Item C003
LEVEL 2 - PREP AREA SDU
[services table]
```

Instead of just showing C001 and C003 (skipping C002).