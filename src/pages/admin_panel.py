"""
Unified admin panel with all management options as sub-tabs.
"""
import streamlit as st
from pages.auth_page import require_admin_access


def admin_panel_page():
    """Main admin panel with all management options."""

    # Require admin access
    require_admin_access()

    # Page header
    st.title("◆ Admin Panel")
    st.markdown("Manage all system settings and data")
    st.markdown("---")

    # Create tabs for different management sections
    tab1, tab2, tab3 = st.tabs([
        "▸ Users",
        "▸ Companies",
        "▸ Templates"
    ])

    # User Management Tab
    with tab1:
        from pages.admin_users import show_user_list, show_pending_invitations, show_create_invitation_form

        user_subtabs = st.tabs(["All Users", "Pending Invitations", "Invite User"])

        with user_subtabs[0]:
            show_user_list()

        with user_subtabs[1]:
            show_pending_invitations()

        with user_subtabs[2]:
            show_create_invitation_form()

    # Company Management Tab
    with tab2:
        from pages.admin_companies import show_company_list, show_add_company_form

        company_subtabs = st.tabs(["All Companies", "Add Company"])

        with company_subtabs[0]:
            show_company_list()

        with company_subtabs[1]:
            show_add_company_form()

    # Template Management Tab
    with tab3:
        from pages.admin_templates import show_template_card, show_backups_list, TEMPLATES
        from utils.template_storage import sync_local_templates_to_storage

        # First-time setup option
        with st.expander("⚙ First-Time Setup"):
            st.markdown("### Sync Local Templates to Supabase Storage")
            st.info("💡 If this is your first time using template management, click below to upload existing templates from the local filesystem to Supabase Storage.")

            if st.button("⬆ Sync Local Templates to Storage", type="primary"):
                with st.spinner("Syncing templates..."):
                    success, message = sync_local_templates_to_storage()

                if success:
                    st.success("✓ Templates synced successfully!")
                    st.code(message)
                else:
                    st.error(f"✕ Error: {message}")

        st.markdown("---")

        template_subtabs = st.tabs(["Templates", "Backups"])

        with template_subtabs[0]:
            st.markdown("### Current Templates")
            st.info("💡 **Note:** Templates are stored in Supabase Storage. When you upload a new template, the current version is automatically backed up.")
            st.markdown("---")

            # Show each template
            for template_key in TEMPLATES.keys():
                show_template_card(template_key)

        with template_subtabs[1]:
            show_backups_list()
