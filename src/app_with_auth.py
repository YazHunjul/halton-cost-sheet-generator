"""
Main Streamlit application for the Halton Cost Sheet Generator WITH AUTHENTICATION.
This is the new entry point that requires users to log in before accessing the app.
"""
import streamlit as st
import sys
import os

# Import authentication
from pages.auth_page import authentication_page, require_authentication, get_current_user
from pages.admin_panel import admin_panel_page
from utils.auth import logout_user

# Import original app functions
from app import (
    word_generation_page,
    revision_page,
    single_page_project_builder,
    initialize_session_state
)


def show_user_sidebar():
    """Display user information and logout in sidebar."""
    current_user = get_current_user()

    if current_user:
        st.sidebar.markdown("---")
        st.sidebar.markdown("### ◆ User Info")
        st.sidebar.write(f"**{current_user['first_name']} {current_user['last_name']}**")
        st.sidebar.caption(f"@ {current_user['email']}")

        # Show role with badge
        if current_user['role'] == 'admin':
            st.sidebar.success("★ Administrator")
        else:
            st.sidebar.info("• User")

        # Logout button
        if st.sidebar.button("⎆ Logout", use_container_width=True):
            success, message = logout_user()
            if success:
                st.success(message)
                st.rerun()
            else:
                st.error(message)


def main():
    """Main application entry point with authentication."""

    # Page configuration
    st.set_page_config(
        page_title="Halton Quotation System",
        page_icon="H",
        layout="wide"
    )

    # Initialize session state
    initialize_session_state()

    # Check if user is authenticated
    if not st.session_state.get("authenticated", False):
        # Show login/signup page
        authentication_page()
        return

    # User is authenticated - show main app
    st.title("■ Halton Quotation System")

    # Sidebar navigation
    st.sidebar.title("◆ Navigation")

    # Build navigation options based on user role
    current_user = get_current_user()
    is_admin = current_user and current_user['role'] == 'admin'

    # Base navigation options (available to all authenticated users)
    nav_options = [
        "▸ Single Page Setup",
        "▸ Generate Word Documents",
        "▸ Create Revision"
    ]

    # Add admin-only option
    if is_admin:
        nav_options.insert(0, "★ Admin Panel")

    page = st.sidebar.selectbox(
        "Choose a page:",
        nav_options
    )

    # Show user info in sidebar
    show_user_sidebar()

    # Route to appropriate page
    if page == "★ Admin Panel":
        admin_panel_page()

    elif page == "▸ Single Page Setup":
        require_authentication()
        single_page_project_builder()

    elif page == "▸ Generate Word Documents":
        require_authentication()
        word_generation_page()

    elif page == "▸ Create Revision":
        require_authentication()
        revision_page()


if __name__ == "__main__":
    main()
