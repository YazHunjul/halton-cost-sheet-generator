"""
Admin company management page.
Manage companies and their addresses.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
from pages.auth_page import require_admin_access, get_current_user
from config.supabase_config import get_supabase_client


def show_add_company_form():
    """Form to add a new company."""
    st.markdown("### + Add New Company")

    with st.form("add_company_form", clear_on_submit=True):
        company_name = st.text_input(
            "Company Name *",
            placeholder="e.g., ABC Catering Equipment Ltd",
            help="Enter the full company name"
        )

        address = st.text_area(
            "Address *",
            placeholder="Line 1\nLine 2\nCity\nPostcode",
            height=150,
            help="Enter the full address (use new lines to separate address lines)"
        )

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            submit_button = st.form_submit_button(
                "+ Add Company",
                type="primary",
                use_container_width=True
            )

        with col2:
            if st.form_submit_button("↻ Clear Form", use_container_width=True):
                st.rerun()

    if submit_button:
        # Validation
        if not company_name or not address:
            st.error("Please fill in both company name and address.")
            return

        # Get current user
        current_user = get_current_user()

        try:
            # Add company to database
            client = get_supabase_client(use_service_role=True)

            company_data = {
                "name": company_name.strip(),
                "address": address.strip(),
                "is_active": True,
                "created_by": current_user["id"]
            }

            response = client.table("companies").insert(company_data).execute()

            if response.data:
                # Refresh the company cache so it appears in dropdowns immediately
                import config.business_data as bd
                bd._companies_cache = None  # Clear cache
                bd.COMPANY_ADDRESSES = bd.get_companies_from_database()  # Reload

                st.success(f"Company '{company_name}' added successfully!")
                st.info("Company will appear in dropdowns immediately!")
                st.balloons()

                # Small delay then rerun to show in list
                import time
                time.sleep(1)
                st.rerun()
            else:
                st.error("Failed to add company.")

        except Exception as e:
            error_msg = str(e)
            if "duplicate key" in error_msg.lower() or "unique" in error_msg.lower():
                st.error(f"Company '{company_name}' already exists!")
            else:
                st.error(f"Error adding company: {error_msg}")


def show_company_list():
    """Display list of all companies."""
    st.markdown("### ■ Company Management")

    try:
        client = get_supabase_client()

        # Fetch all companies
        response = client.table("companies").select("*").order("name").execute()

        if not response.data:
            st.info("No companies found. Add your first company above!")
            return

        # Create DataFrame
        companies_df = pd.DataFrame(response.data)

        # Separate active and inactive companies
        active_companies = companies_df[companies_df["is_active"] == True]
        inactive_companies = companies_df[companies_df["is_active"] == False]

        # Display summary metrics
        col1, col2, col3 = st.columns(3)

        with col1:
            total_companies = len(companies_df)
            st.metric("Total Companies", total_companies)

        with col2:
            active_count = len(active_companies)
            st.metric("Active", active_count, delta=None, delta_color="normal")

        with col3:
            inactive_count = len(inactive_companies)
            st.metric("Deactivated", inactive_count, delta=None, delta_color="inverse")

        st.markdown("---")

        # Search bar for both sections
        search_term = st.text_input(
            "◇ Search Companies",
            placeholder="Search by name or address...",
            help="Search across both active and deactivated companies"
        )

        # --- ACTIVE COMPANIES SECTION ---
        st.markdown("### ✓ Active Companies")
        st.markdown(f"*{len(active_companies)} companies currently active and visible in dropdowns*")

        # Apply search filter to active companies
        filtered_active = active_companies.copy()
        if search_term:
            mask = (
                filtered_active["name"].str.contains(search_term, case=False, na=False) |
                filtered_active["address"].str.contains(search_term, case=False, na=False)
            )
            filtered_active = filtered_active[mask]

        if len(filtered_active) == 0:
            st.info("No active companies found matching your search.")
        else:
            st.markdown(f"**Showing {len(filtered_active)} active companies**")

        # Display active companies
        for idx, company in filtered_active.iterrows():
            status_icon = "[ACTIVE]" if company["is_active"] else "[INACTIVE]"

            with st.expander(
                f"{status_icon} {company['name']}",
                expanded=False
            ):
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.markdown("**Address:**")
                    # Display address with proper line breaks
                    address_lines = company["address"].split('\n')
                    for line in address_lines:
                        st.text(line)

                with col2:
                    st.write("**Status:**", "Active" if company["is_active"] else "Inactive")

                    created_date = datetime.fromisoformat(company["created_at"].replace("Z", "+00:00"))
                    st.write("**Added:**", created_date.strftime("%Y-%m-%d"))

                st.markdown("---")

                # Edit section
                st.markdown("##### Edit Company")

                with st.form(f"edit_company_{company['id']}", clear_on_submit=False):
                    new_name = st.text_input(
                        "Company Name",
                        value=company["name"],
                        key=f"name_{company['id']}"
                    )

                    new_address = st.text_area(
                        "Address",
                        value=company["address"],
                        height=150,
                        key=f"address_{company['id']}"
                    )

                    col1, col2, col3 = st.columns(3)

                    with col1:
                        update_button = st.form_submit_button(
                            "⬆ Update",
                            use_container_width=True
                        )

                    with col2:
                        # Only show deactivate button in active section
                        deactivate_button = st.form_submit_button(
                            "✕ Deactivate",
                            use_container_width=True
                        )

                    with col3:
                        delete_button = st.form_submit_button(
                            "⊗ Delete",
                            use_container_width=True
                        )

                # Handle update
                if update_button:
                    if not new_name or not new_address:
                        st.error("Name and address cannot be empty!")
                    else:
                        try:
                            current_user = get_current_user()
                            client = get_supabase_client(use_service_role=True)

                            update_data = {
                                "name": new_name.strip(),
                                "address": new_address.strip(),
                                "updated_by": current_user["id"]
                            }

                            response = client.table("companies").update(
                                update_data
                            ).eq("id", company["id"]).execute()

                            if response.data:
                                # Refresh cache
                                import config.business_data as bd
                                bd._companies_cache = None
                                bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                                st.success(f"Company updated successfully!")
                                st.rerun()
                            else:
                                st.error("Failed to update company.")

                        except Exception as e:
                            error_msg = str(e)
                            if "duplicate key" in error_msg.lower():
                                st.error(f"Company name already exists!")
                            else:
                                st.error(f"Error: {error_msg}")

                # Handle deactivate (only for active companies section)
                if deactivate_button:
                    try:
                        client = get_supabase_client(use_service_role=True)
                        response = client.table("companies").update(
                            {"is_active": False}
                        ).eq("id", company["id"]).execute()

                        if response.data:
                            # Refresh cache
                            import config.business_data as bd
                            bd._companies_cache = None
                            bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                            st.success(f"Company deactivated and removed from dropdowns!")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {str(e)}")

                # Handle delete
                if delete_button:
                    try:
                        client = get_supabase_client(use_service_role=True)
                        response = client.table("companies").delete().eq(
                            "id", company["id"]
                        ).execute()

                        # Refresh cache
                        import config.business_data as bd
                        bd._companies_cache = None
                        bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                        st.success(f"Company '{company['name']}' deleted!")
                        st.rerun()

                    except Exception as e:
                        st.error(f"Error deleting company: {str(e)}")

        # --- DEACTIVATED COMPANIES SECTION ---
        st.markdown("---")
        st.markdown("### ✕ Deactivated Companies")
        st.markdown(f"*{len(inactive_companies)} companies currently deactivated and hidden from dropdowns*")

        # Apply search filter to inactive companies
        filtered_inactive = inactive_companies.copy()
        if search_term:
            mask = (
                filtered_inactive["name"].str.contains(search_term, case=False, na=False) |
                filtered_inactive["address"].str.contains(search_term, case=False, na=False)
            )
            filtered_inactive = filtered_inactive[mask]

        if len(filtered_inactive) == 0:
            if len(inactive_companies) == 0:
                st.success("No deactivated companies. All companies are active!")
            else:
                st.info("No deactivated companies found matching your search.")
        else:
            st.markdown(f"**Showing {len(filtered_inactive)} deactivated companies**")

        # Display deactivated companies
        for idx, company in filtered_inactive.iterrows():
            with st.expander(
                f"[DEACTIVATED] {company['name']}",
                expanded=False
            ):
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.markdown("**Address:**")
                    address_lines = company["address"].split('\n')
                    for line in address_lines:
                        st.text(line)

                with col2:
                    st.write("**Status:**", "Deactivated")
                    created_date = datetime.fromisoformat(company["created_at"].replace("Z", "+00:00"))
                    st.write("**Added:**", created_date.strftime("%Y-%m-%d"))

                st.markdown("---")

                # Edit section
                st.markdown("##### Edit Company")

                with st.form(f"edit_company_{company['id']}", clear_on_submit=False):
                    new_name = st.text_input(
                        "Company Name",
                        value=company["name"],
                        key=f"edit_name_{company['id']}"
                    )

                    new_address = st.text_area(
                        "Address",
                        value=company["address"],
                        height=100,
                        key=f"edit_address_{company['id']}"
                    )

                    update_button = st.form_submit_button("⬆ Update", type="primary")

                if update_button:
                    if not new_name or not new_address:
                        st.error("Both name and address are required!")
                    else:
                        try:
                            client = get_supabase_client(use_service_role=True)
                            response = client.table("companies").update({
                                "name": new_name.strip(),
                                "address": new_address.strip()
                            }).eq("id", company["id"]).execute()

                            if response.data:
                                # Refresh cache
                                import config.business_data as bd
                                bd._companies_cache = None
                                bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                                st.success(f"Company updated successfully!")
                                st.rerun()
                        except Exception as e:
                            error_msg = str(e)
                            if "duplicate key" in error_msg.lower():
                                st.error(f"Company name already exists!")
                            else:
                                st.error(f"Error: {error_msg}")

                # Actions section
                st.markdown("##### Actions")

                col1, col2 = st.columns(2)

                with col1:
                    # Activate button (since this is inactive section)
                    activate_button = st.button(
                        "✓ Reactivate",
                        key=f"activate_{company['id']}",
                        use_container_width=True,
                        type="primary"
                    )

                with col2:
                    delete_button = st.button(
                        "⊗ Delete Permanently",
                        key=f"delete_{company['id']}",
                        use_container_width=True,
                        type="secondary"
                    )

                # Handle activate
                if activate_button:
                    try:
                        client = get_supabase_client(use_service_role=True)
                        response = client.table("companies").update(
                            {"is_active": True}
                        ).eq("id", company["id"]).execute()

                        if response.data:
                            # Refresh cache
                            import config.business_data as bd
                            bd._companies_cache = None
                            bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                            st.success(f"Company reactivated! It will now appear in dropdowns.")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {str(e)}")

                # Handle delete
                if delete_button:
                    try:
                        client = get_supabase_client(use_service_role=True)
                        response = client.table("companies").delete().eq(
                            "id", company["id"]
                        ).execute()

                        # Refresh cache
                        import config.business_data as bd
                        bd._companies_cache = None
                        bd.COMPANY_ADDRESSES = bd.get_companies_from_database()

                        st.success(f"Company '{company['name']}' permanently deleted!")
                        st.rerun()

                    except Exception as e:
                        st.error(f"Error deleting company: {str(e)}")

    except Exception as e:
        st.error(f"Error loading companies: {str(e)}")


def admin_companies_page():
    """Main admin companies management page."""

    # Require admin access
    require_admin_access()

    # Page header
    st.title("■ Company Management")
    st.markdown("Manage companies and their addresses for project quotations.")
    st.markdown("---")

    # Create tabs
    tab1, tab2 = st.tabs(["▸ All Companies", "+ Add Company"])

    with tab1:
        show_company_list()

    with tab2:
        show_add_company_form()
