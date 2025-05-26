"""
Form components for the Halton Cost Sheet Generator.
"""
import streamlit as st
from typing import Dict, Any
from datetime import datetime

from config.constants import (
    ESTIMATORS,
    COMPANY_ADDRESSES,
    SALES_CONTACTS,
    DELIVERY_LOCATIONS,
    SessionKeys
)

def general_project_form() -> Dict[str, Any]:
    """
    Renders the general project information form and returns the collected data.
    
    Returns:
        Dict[str, Any]: Dictionary containing the form data
    """
    # Initialize form data in session state if not exists
    if "general_form_data" not in st.session_state:
        st.session_state.general_form_data = {}
    
    # Company selection mode (outside form for immediate reactivity)
    company_mode = st.radio(
        "Company Selection *",
        options=["Select from list", "Enter custom company"],
        index=0 if st.session_state.general_form_data.get("company_mode", "Select from list") == "Select from list" else 1,
        key="company_mode_input",
        help="Choose whether to select from predefined companies or enter a custom company"
    )
    
    st.markdown("---")  # Visual separator
    
    with st.form("general_project_info", clear_on_submit=False):
        col1, col2 = st.columns(2)
        
        with col1:
            project_name = st.text_input(
                "Project Name *",
                value=st.session_state.general_form_data.get("project_name", ""),
                key="project_name_input",
                help="Enter the name of the project"
            )
            
            project_number = st.text_input(
                "Project Number *",
                value=st.session_state.general_form_data.get("project_number", ""),
                key="project_number_input",
                help="Enter the unique project number"
            )
            
            date = st.date_input(
                "Date *",
                value=datetime.strptime(st.session_state.general_form_data.get("date", datetime.now().strftime("%Y-%m-%d")), "%Y-%m-%d").date(),
                key="date_input",
                help="Select the project date"
            )
            
            customer = st.text_input(
                "Customer *",
                value=st.session_state.general_form_data.get("customer", ""),
                key="customer_input",
                help="Enter the customer name"
            )
            
            if company_mode == "Select from list":
                company = st.selectbox(
                    "Company *",
                    options=list(COMPANY_ADDRESSES.keys()),
                    index=list(COMPANY_ADDRESSES.keys()).index(st.session_state.general_form_data.get("company", list(COMPANY_ADDRESSES.keys())[0])) if st.session_state.general_form_data.get("company") in COMPANY_ADDRESSES else 0,
                    key="company_select",
                    help="Select the company from the predefined list"
                )
                custom_company_name = ""
                custom_company_address = ""
            else:
                company = ""
                custom_company_name = st.text_input(
                    "Custom Company Name *",
                    value=st.session_state.general_form_data.get("custom_company_name", ""),
                    key="custom_company_name_input",
                    help="Enter the custom company name"
                )
                custom_company_address = st.text_area(
                    "Custom Company Address *",
                    value=st.session_state.general_form_data.get("custom_company_address", ""),
                    key="custom_company_address_input",
                    help="Enter the full company address (use line breaks for multiple lines)",
                    height=100
                )
            
            project_location = st.text_input(
                "Project Location *",
                value=st.session_state.general_form_data.get("project_location", ""),
                key="project_location_input",
                help="Enter the project location"
            )
            
            delivery_location = st.selectbox(
                "Delivery Location *",
                options=DELIVERY_LOCATIONS,
                index=DELIVERY_LOCATIONS.index(st.session_state.general_form_data.get("delivery_location", DELIVERY_LOCATIONS[0])) if st.session_state.general_form_data.get("delivery_location") in DELIVERY_LOCATIONS else 0,
                key="delivery_location_input",
                help="Select the delivery location"
            )
        
        with col2:
            sales_contact = st.selectbox(
                "Sales Contact *",
                options=list(SALES_CONTACTS.keys()),
                index=list(SALES_CONTACTS.keys()).index(st.session_state.general_form_data.get("sales_contact", list(SALES_CONTACTS.keys())[0])) if st.session_state.general_form_data.get("sales_contact") in SALES_CONTACTS else 0,
                key="sales_contact_input",
                help="Select the sales contact"
            )
            
            estimator = st.selectbox(
                "Estimator *",
                options=list(ESTIMATORS.keys()),
                index=list(ESTIMATORS.keys()).index(st.session_state.general_form_data.get("estimator", list(ESTIMATORS.keys())[0])) if st.session_state.general_form_data.get("estimator") in ESTIMATORS else 0,
                key="estimator_select",
                help="Select the estimator for this project"
            )
        
        submitted = st.form_submit_button("Save & Continue")
        
        if submitted:
            # Validate required fields based on company mode
            if company_mode == "Select from list":
                required_fields = [
                    project_name,
                    project_number,
                    date,
                    customer,
                    company,
                    project_location,
                    delivery_location,
                    sales_contact,
                    estimator
                ]
            else:  # Custom company
                required_fields = [
                    project_name,
                    project_number,
                    date,
                    customer,
                    custom_company_name,
                    custom_company_address,
                    project_location,
                    delivery_location,
                    sales_contact,
                    estimator
                ]
            
            if not all(required_fields):
                st.error("Please fill in all required fields marked with *")
                return None
            
            # Determine company name and address based on mode
            if company_mode == "Select from list":
                final_company_name = company
                final_company_address = COMPANY_ADDRESSES[company]
            else:  # Custom company
                final_company_name = custom_company_name
                final_company_address = custom_company_address
            
            project_data = {
                "project_name": project_name,
                "project_number": project_number,
                "date": date.strftime("%Y-%m-%d"),
                "customer": customer,
                "company": final_company_name,
                "project_location": project_location,
                "delivery_location": delivery_location,
                "location": project_location,  # Keep 'location' for backward compatibility (maps to project_location)
                "address": final_company_address,  # Store the full address
                "sales_contact": sales_contact,
                "estimator": estimator,
                
                # Store the selection mode and custom fields for form persistence
                "company_mode": company_mode,
                "custom_company_name": custom_company_name,
                "custom_company_address": custom_company_address
            }
            
            # Store in session state for form persistence
            st.session_state.general_form_data = project_data
            return project_data
            
    # Return the current form data even if not submitted
    if st.session_state.general_form_data:
        return st.session_state.general_form_data
    return None 