"""
Constants and configuration values for the Halton Cost Sheet Generator.
"""
from .business_data import (
    SALES_CONTACTS,
    ESTIMATORS,
    COMPANY_ADDRESSES,
    VALID_CANOPY_MODELS,
    DELIVERY_LOCATIONS
)

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