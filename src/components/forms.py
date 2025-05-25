"""
Form components for the HVAC Project Management Tool.
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
            
            company = st.selectbox(
                "Company *",
                options=list(COMPANY_ADDRESSES.keys()),
                index=list(COMPANY_ADDRESSES.keys()).index(st.session_state.general_form_data.get("company", list(COMPANY_ADDRESSES.keys())[0])) if st.session_state.general_form_data.get("company") in COMPANY_ADDRESSES else 0,
                key="company_input",
                help="Select the company"
            )
            
            location = st.selectbox(
                "Location *",
                options=DELIVERY_LOCATIONS,
                index=DELIVERY_LOCATIONS.index(st.session_state.general_form_data.get("location", DELIVERY_LOCATIONS[0])) if st.session_state.general_form_data.get("location") in DELIVERY_LOCATIONS else 0,
                key="location_input",
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
            
            cost_sheet = st.text_input(
                "Cost Sheet Reference",
                value=st.session_state.general_form_data.get("cost_sheet", ""),
                key="cost_sheet_input",
                help="Enter the cost sheet reference number (optional)"
            )
        
        submitted = st.form_submit_button("Save & Continue")
        
        if submitted:
            required_fields = [
                project_name,
                project_number,
                date,
                customer,
                company,
                location,
                sales_contact,
                estimator
            ]
            
            if not all(required_fields):
                st.error("Please fill in all required fields marked with *")
                return None
            
            project_data = {
                "project_name": project_name,
                "project_number": project_number,
                "date": date.strftime("%Y-%m-%d"),
                "customer": customer,
                "company": company,
                "location": location,
                "address": COMPANY_ADDRESSES[company],  # Store the full address
                "sales_contact": sales_contact,
                "estimator": estimator,
                "cost_sheet": cost_sheet
            }
            
            # Store in session state for form persistence
            st.session_state.general_form_data = project_data
            return project_data
            
    # Return the current form data even if not submitted
    if st.session_state.general_form_data:
        return st.session_state.general_form_data
    return None 