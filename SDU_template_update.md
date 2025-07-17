# SDU Template Update for Word Document

## Key Changes Made:

1. **Fixed the `area_name` error** by adding the `area_name` key to both `area_data` and `enhanced_area` dictionaries in word.py

2. **Added default fields to SDU data**:
   - `sdu_length` (default: 'XXXX')
   - `potrack` (default: 'xxxxxxxxxxxxx')
   - `salamander_support` (default: 'xxxxxxxxxxxxxx')
   - All electrical, gas, and water services initialized with proper defaults

## Updated Jinja Template Code:

### For the SDU Specification Section:

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
All Halton Foodservice electrical installations shall be wired and tested to conform with the latest BS7671:2018 IET Wiring Regulations{%endif%}
```

### For the Pricing Section:

```jinja
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
```

## Data Structure:

The `sdu_areas` list contains objects with this structure:

```javascript
{
    'level_name': 'Ground Floor',
    'area_name': 'Kitchen',
    'area_number': 1,
    'canopy_reference': '1.01',
    'level_area_combined': 'Ground Floor - Kitchen',
    'has_sdu': true,
    'sdu_price': 5000,
    'sdu_length': 'XXXX',  // Will be updated from Excel if available
    'potrack': 'xxxxxxxxxxxxx',
    'salamander_support': 'xxxxxxxxxxxxxx',
    'pricing': {
        'final_carcass_price': 3000,
        'final_electrical_price': 2000,
        'live_site_test_price': 500,
        'has_live_test': false,
        'total_price': 5000
    },
    'electrical_services': {
        'distribution_board': 1,
        'single_phase_switched_spur': 2,
        'three_phase_socket_outlet': 1,
        'switched_socket_outlet': 4,
        'emergency_knock_off': 0,
        'ring_main_inc_2no_sso': 2
    },
    'gas_services': {
        'gas_manifold': 1,
        'gas_connection_15mm': 2,
        'gas_connection_20mm': 1,
        'gas_connection_25mm': 0,
        'gas_connection_32mm': 0,
        'gas_solenoid_valve': 1
    },
    'water_services': {
        'cws_manifold_22mm': 1,
        'cws_manifold_15mm': 1,
        'hws_manifold': 1,
        'water_connection_15mm': 4,
        'water_connection_22mm': 2,
        'water_connection_28mm': 0
    }
}
```

## Notes:

1. The template groups SDUs by area using the `processed_areas` list to avoid duplicates
2. Default values are provided for all fields to prevent template errors
3. The pricing section looks for SDUs in the current area being processed
4. All monetary values use the `format_currency` filter
5. The `{%tr%}` tags are used for conditional table rows

## Testing:

To test the changes:
1. Ensure you have a project with SDU options enabled
2. Generate the Word document
3. Check that SDU sections appear correctly in both specification and pricing sections
4. Verify that all service counts and prices are displayed properly