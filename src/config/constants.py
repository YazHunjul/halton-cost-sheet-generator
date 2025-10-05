"""
Constants and configuration values for the Halton Cost Sheet Generator.
"""
from .business_data import (
    SALES_CONTACTS,
    ESTIMATORS,
    VALID_CANOPY_MODELS,
    DELIVERY_LOCATIONS,
    get_company_addresses
)

# Company addresses - loaded dynamically with caching
# Use get_company_addresses(refresh=True) to force reload
def COMPANY_ADDRESSES():
    """Get active company addresses (cached)."""
    return get_company_addresses(active_only=True)

# For backward compatibility, also create a direct reference
COMPANY_ADDRESSES = get_company_addresses(active_only=True)

# Feature flags for systems not yet implemented
# Set to True when ready to display to users
FEATURE_FLAGS = {
    # Kitchen systems
    "kitchen_extract_system": False,
    "kitchen_makeup_air_system": False,
    
    # Advanced control systems
    "marvel_system": False,  # M.A.R.V.E.L. System (DCKV)
    
    # Ceiling systems
    "cyclocell_cassette_ceiling": True,  # Changed from False to True
    
    # Additional equipment
    "reactaway_unit": False,
    
    # Future systems (placeholders)
    "dishwasher_extract": False,
    "gas_interlocking": False,
    "pollustop_unit": False,
}

def is_feature_enabled(feature_name: str) -> bool:
    """
    Check if a feature is enabled.
    
    Args:
        feature_name: Name of the feature to check
        
    Returns:
        bool: True if feature is enabled, False otherwise
    """
    return FEATURE_FLAGS.get(feature_name, False)

# Project types
PROJECT_TYPES = [
    "Canopy Project",
    "RecoAir Project"
]

# Template file paths
TEMPLATES = {
    "excel": {
        "cost_sheet": "templates/excel/cost_sheet_template.xlsx"
    },
    "word": {
        "canopy_quotation": "templates/word/canopy_quotation_template.docx",
        "recoair_quotation": "templates/word/recoair_quotation_template.docx"
    }
}

# Session state keys
class SessionKeys:
    PROJECT_DATA = "project_data"
    CURRENT_STEP = "current_step"
    PROJECT_TYPE = "project_type" 