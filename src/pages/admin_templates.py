"""
Admin template management page.
Allows admins to view and upload Word document templates to Supabase Storage.
"""
import streamlit as st
from datetime import datetime
from pages.auth_page import require_admin_access
from utils.template_storage import (
    upload_template_to_storage,
    download_template_from_storage,
    get_template_metadata,
    list_template_backups,
    sync_local_templates_to_storage,
    TEMPLATE_FILES
)

# Template definitions with descriptions
TEMPLATES = {
    "canopy_quotation": {
        "name": "Canopy Quotation Template",
        "filename": TEMPLATE_FILES["canopy_quotation"],
        "description": "Template used for generating canopy project quotations"
    },
    "recoair_quotation": {
        "name": "RecoAir Quotation Template",
        "filename": TEMPLATE_FILES["recoair_quotation"],
        "description": "Template used for generating RecoAir project quotations"
    },
    "ahu_quotation": {
        "name": "AHU Quotation Template",
        "filename": TEMPLATE_FILES["ahu_quotation"],
        "description": "Template used for generating AHU project quotations"
    }
}


def show_template_card(template_key):
    """Display a template card with info and upload option."""
    template_info = TEMPLATES.get(template_key)

    if not template_info:
        st.error(f"Template configuration not found: {template_key}")
        return

    with st.container():
        st.markdown(f"### {template_info['name']}")
        st.caption(template_info['description'])

        col1, col2 = st.columns([2, 1])

        with col1:
            # Get metadata from Supabase Storage
            success, metadata, message = get_template_metadata(template_key)

            if success and metadata:
                st.write("**Status:** ✓ Active")
                st.write(f"**File:** {template_info['filename']}")
                st.write(f"**Size:** {metadata.get('size', 0):,} bytes")

                if metadata.get('updated_at'):
                    updated_at = datetime.fromisoformat(metadata['updated_at'].replace('Z', '+00:00'))
                    st.write(f"**Last Modified:** {updated_at.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                st.warning("**Status:** ✕ Not found in storage")
                st.caption(f"Expected: {template_info['filename']}")

        with col2:
            # Download current template
            if success:
                # Download from Supabase Storage
                dl_success, file_bytes, dl_message = download_template_from_storage(template_key)

                if dl_success:
                    st.download_button(
                        label="⬇ Download",
                        data=file_bytes,
                        file_name=template_info['filename'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"download_{template_key}",
                        use_container_width=True
                    )

        # Upload new template
        st.markdown("---")
        st.markdown("##### Upload New Template")

        uploaded_file = st.file_uploader(
            "Choose a Word document (.docx)",
            type=['docx'],
            key=f"upload_{template_key}",
            help=f"Upload a new template to replace {template_info['filename']}"
        )

        if uploaded_file is not None:
            col1, col2 = st.columns(2)

            with col1:
                st.info(f"**Selected:** {uploaded_file.name}")
                st.write(f"**Size:** {uploaded_file.size:,} bytes")

            with col2:
                if st.button("⬆ Upload & Replace", key=f"confirm_{template_key}", type="primary", use_container_width=True):
                    with st.spinner("Uploading template to Supabase Storage..."):
                        # Get file bytes
                        file_bytes = uploaded_file.getvalue()

                        # Upload to Supabase Storage
                        success, message = upload_template_to_storage(
                            template_key,
                            file_bytes,
                            template_info['filename']
                        )

                    if success:
                        st.success(f"✓ {message}")
                        st.info("💡 The new template is now stored in Supabase and will be used for all future document generations.")
                        st.rerun()
                    else:
                        st.error(f"✕ {message}")

        st.markdown("---")


def show_backups_list():
    """Display list of template backups from Supabase Storage."""
    st.markdown("### ◆ Template Backups")
    st.caption("Backups are stored in Supabase Storage")

    all_backups = []

    for template_key in TEMPLATES.keys():
        success, backups, message = list_template_backups(template_key)

        if success and backups:
            for backup in backups:
                all_backups.append({
                    'template_key': template_key,
                    'name': backup['name'],
                    'created_at': backup.get('created_at', backup.get('updated_at')),
                    'size': backup.get('metadata', {}).get('size', 0),
                    'template_name': TEMPLATES[template_key]['name']
                })

    if not all_backups:
        st.info("No backups found. Backups are created automatically when you upload a new template.")
        return

    # Sort by created_at descending (newest first)
    all_backups.sort(key=lambda x: x.get('created_at', ''), reverse=True)

    st.write(f"**Total Backups:** {len(all_backups)}")
    st.markdown("---")

    # Show last 20 backups
    for backup in all_backups[:20]:
        col1, col2 = st.columns([3, 1])

        with col1:
            st.write(f"**{backup['name']}**")
            created_at = datetime.fromisoformat(backup['created_at'].replace('Z', '+00:00'))
            st.caption(f"{backup['template_name']} | Created: {created_at.strftime('%Y-%m-%d %H:%M:%S')} | Size: {backup['size']:,} bytes")

        with col2:
            # Download backup
            if st.button("⬇ Download", key=f"backup_{backup['template_key']}_{backup['name']}", use_container_width=True):
                try:
                    from config.supabase_config import get_supabase_client
                    client = get_supabase_client(use_service_role=True)
                    storage_path = f"backups/{backup['template_key']}/{backup['name']}"

                    file_bytes = client.storage.from_("templates").download(storage_path)

                    st.download_button(
                        label="💾 Save",
                        data=file_bytes,
                        file_name=backup['name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"save_{backup['template_key']}_{backup['name']}",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error downloading backup: {str(e)}")

        st.markdown("---")


def admin_templates_page():
    """Main admin templates management page."""

    # Require admin access
    require_admin_access()

    # Page header
    st.title("◆ Template Management")
    st.markdown("Manage Word document templates stored in Supabase Storage.")
    st.markdown("---")

    # Option to sync local templates to storage (for first-time setup)
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

    # Create tabs
    tab1, tab2 = st.tabs(["▸ Templates", "▸ Backups"])

    with tab1:
        st.markdown("### Current Templates")
        st.info("💡 **Note:** Templates are stored in Supabase Storage. When you upload a new template, the current version is automatically backed up. You can restore from backups in the Backups tab.")
        st.markdown("---")

        # Show each template
        for template_key in TEMPLATES.keys():
            show_template_card(template_key)

    with tab2:
        show_backups_list()
