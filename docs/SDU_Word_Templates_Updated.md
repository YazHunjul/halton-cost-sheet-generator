# Updated SDU Word Templates with Item Numbers

## SDU Specification Section - Updated Template

Replace your current SDU specification section with this updated version that uses the SDU item number:

```jinja
{%if has_sdu%}
Services Distribution Unit
The services distribution units (SDU) are manufactured from 1.2mm thick, grade 304 stainless steel from an all folded and welded construction. The SDU is to be manufactured in suitably sized sections to suit site access requirements. The SDU is to be supplied with factory fitted mechanical and electrical services as listed below.
{% for sdu in sdu_areas %}

Item {{ sdu.sdu_item_number if sdu.sdu_item_number else sdu.canopy_reference }}
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
{% endfor %}
All Halton Foodservice electrical installations shall be wired and tested to conform with the latest BS7671:2018 IET Wiring Regulations
{%endif%}
```

## SDU Pricing Section - Updated Template

Replace your SDU pricing section with this updated version:

```jinja
{% if area.has_sdu%}
{% set area_sdus = sdu_areas | selectattr('level_area_combined', 'equalto', area.level_area_combined) | list %}
{% if area_sdus %}
    ITEM    SERVICE DISTRIBUTION UNIT SCHEDULE    PRICE
{% for sdu in area_sdus %}
    {{sdu.sdu_item_number if sdu.sdu_item_number else sdu.canopy_reference}}    Service Distribution Unit, {{sdu.sdu_length}}mm long, supplied & installed (carcass only).    {{format_currency(sdu.pricing.final_carcass_price)}}
        Electrical & mechanical services, supply and install    {{format_currency(sdu.pricing.final_electrical_price)}}
    {%tr if sdu.pricing.has_live_test %}        
        Extra over cost for separate 'Live Site Test'    {{format_currency(sdu.pricing.live_site_test_price)}}
    {%tr endif%}        
        Local shunt trip for electrical shut off (WIRING CONNECTION BY OTHERS)    
{% endfor %}
    SUB TOTAL    {{format_currency(area.sdu_subtotal)}}
{% endif %}
{%endif%}
```

## Key Changes:

1. **Item Number Display**: 
   - Uses `{{ sdu.sdu_item_number if sdu.sdu_item_number else sdu.canopy_reference }}`
   - This displays the SDU item number if provided, otherwise falls back to the canopy reference

2. **In Pricing Section**:
   - Same logic for displaying item numbers in the pricing table

## How the SDU Item Number Works:

1. **Input**: When you check SDU for a canopy, an "SDU Item Number" text field appears
2. **Storage**: The item number is saved with the canopy data
3. **Excel**: The item number is written to cell B12 in the SDU sheet
4. **Word**: The item number is displayed in both the specification and pricing sections

## Example Output:

If you enter "9.01" as the SDU item number for canopy C001:

**Specification Section:**
```
Item 9.01
LEVEL 1 - KITCHEN SDU
[services table]
```

**Pricing Section:**
```
9.01    Service Distribution Unit, 3000mm long, supplied & installed (carcass only).    £3,000.00
```

## Notes:

- If no SDU item number is provided, the system will use the canopy reference (e.g., "C001")
- The SDU item number is preserved when saving/loading Excel files
- You can edit the SDU item number at any time in the app