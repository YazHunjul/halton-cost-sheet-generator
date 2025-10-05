"""
Authentication page for login, signup, and invitation acceptance.
"""
import streamlit as st
from utils.auth import (
    login_user,
    accept_invitation,
    verify_invitation,
    logout_user
)
from datetime import datetime


def show_login_form():
    """Display login form."""
    st.markdown("###  Login to Halton Quotation System")
    st.markdown("---")

    with st.form("login_form", clear_on_submit=False):
        email = st.text_input(
            "Email Address",
            placeholder="your.email@company.com",
            key="login_email"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter your password",
            key="login_password"
        )

        col1, col2 = st.columns([1, 1])

        with col1:
            login_button = st.form_submit_button(
                "Login",
                type="primary",
                use_container_width=True
            )

        with col2:
            if st.form_submit_button("Have an invitation?", use_container_width=True):
                st.session_state.auth_mode = "signup"
                st.rerun()

    if login_button:
        if not email or not password:
            st.error(" Please enter both email and password.")
            return

        with st.spinner("Logging in..."):
            success, user_data, message = login_user(email, password)

        if success:
            # Store user data in session
            st.session_state.authenticated = True
            st.session_state.user_id = user_data["id"]
            st.session_state.user_email = user_data["email"]
            st.session_state.user_first_name = user_data["first_name"]
            st.session_state.user_last_name = user_data["last_name"]
            st.session_state.user_role = user_data["role"]
            st.session_state.access_token = user_data["access_token"]
            st.session_state.refresh_token = user_data["refresh_token"]

            st.success(f" Welcome back, {user_data['first_name']}!")
            st.balloons()

            # Small delay to show success message
            import time
            time.sleep(1)

            # Redirect to main app
            st.rerun()
        else:
            st.error(f" {message}")


def show_signup_form():
    """Display signup form for invitation acceptance."""
    st.markdown("###  Accept Your Invitation")
    st.markdown("---")

    # Get invitation token from URL or input
    query_params = st.query_params
    invitation_token = query_params.get("token", None)

    if not invitation_token:
        st.info(" Enter the invitation token from your email")

        invitation_token = st.text_input(
            "Invitation Token",
            placeholder="Paste your invitation token here",
            help="You should have received this token in your invitation email"
        )

        if st.button("Verify Token"):
            if invitation_token:
                st.session_state.temp_invitation_token = invitation_token
                st.rerun()
            else:
                st.warning(" Please enter an invitation token")

        if st.button("← Back to Login"):
            st.session_state.auth_mode = "login"
            st.rerun()

        return

    # Verify invitation
    with st.spinner("Verifying invitation..."):
        valid, invitation_data, message = verify_invitation(invitation_token)

    if not valid:
        st.error(f" {message}")
        st.markdown("---")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Try Another Token", use_container_width=True):
                # Clear token from URL
                st.query_params.clear()
                st.rerun()

        with col2:
            if st.button("← Back to Login", use_container_width=True):
                st.session_state.auth_mode = "login"
                st.query_params.clear()
                st.rerun()

        return

    # Show invitation details
    st.success(" Valid invitation found!")

    with st.container():
        st.markdown("#### Your Account Details")
        col1, col2 = st.columns(2)

        with col1:
            st.write("**First Name:**", invitation_data["first_name"])
            st.write("**Last Name:**", invitation_data["last_name"])

        with col2:
            st.write("**Email:**", invitation_data["email"])
            st.write("**Role:**", invitation_data["role"].upper())

        # Show role description
        if invitation_data["role"] == "admin":
            st.info(" **Admin Access**: You will have full access to manage users and system settings.")
        else:
            st.info(" **User Access**: You will be able to create and manage your own projects.")

    st.markdown("---")
    st.markdown("#### Set Your Password")

    with st.form("signup_form", clear_on_submit=False):
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Choose a strong password (min 8 characters)",
            help="Password must be at least 8 characters long"
        )

        confirm_password = st.text_input(
            "Confirm Password",
            type="password",
            placeholder="Re-enter your password"
        )

        terms_accepted = st.checkbox(
            "I accept the terms and conditions",
            help="By checking this box, you agree to use the system responsibly"
        )

        submit_button = st.form_submit_button(
            "Create Account",
            type="primary",
            use_container_width=True
        )

    if submit_button:
        # Validation
        if not password or not confirm_password:
            st.error(" Please fill in all password fields.")
            return

        if len(password) < 8:
            st.error(" Password must be at least 8 characters long.")
            return

        if password != confirm_password:
            st.error(" Passwords do not match.")
            return

        if not terms_accepted:
            st.error(" You must accept the terms and conditions.")
            return

        # Accept invitation and create account
        with st.spinner("Creating your account..."):
            success, message = accept_invitation(invitation_token, password)

        if success:
            st.success(f" {message}")
            st.balloons()
            st.info(" Redirecting to login page...")

            # Clear invitation token from URL
            st.query_params.clear()

            # Small delay to show success message
            import time
            time.sleep(2)

            # Redirect to login
            st.session_state.auth_mode = "login"
            st.rerun()
        else:
            st.error(f" {message}")


def authentication_page():
    """Main authentication page with login and signup."""

    # Check if already authenticated
    if st.session_state.get("authenticated", False):
        # User is already logged in, show logout option
        st.success(f" You are logged in as **{st.session_state.get('user_first_name', 'User')} {st.session_state.get('user_last_name', '')}**")
        st.info(f"**Role:** {st.session_state.get('user_role', 'user').upper()}")

        if st.button(" Logout"):
            success, message = logout_user()
            if success:
                st.success(message)
                st.rerun()
            else:
                st.error(message)

        st.markdown("---")
        st.info(" Use the sidebar to navigate to different pages.")
        return

    # Page configuration
    st.title(" Halton Quotation System")
    st.markdown("### Welcome! Please log in to continue.")

    # Initialize auth mode in session state
    if "auth_mode" not in st.session_state:
        # Check if there's an invitation token in URL
        query_params = st.query_params
        if query_params.get("token"):
            st.session_state.auth_mode = "signup"
        else:
            st.session_state.auth_mode = "login"

    # Create tabs for login and signup
    tab1, tab2 = st.tabs([" Login", " Sign Up with Invitation"])

    with tab1:
        if st.session_state.auth_mode != "login":
            st.session_state.auth_mode = "login"
        show_login_form()

    with tab2:
        if st.session_state.auth_mode != "signup":
            st.session_state.auth_mode = "signup"
        show_signup_form()

    # Footer
    st.markdown("---")
    st.caption(" Secure authentication powered by Supabase")


def require_authentication():
    """
    Check if user is authenticated. If not, show login page and stop execution.
    Use this at the top of pages that require authentication.
    """
    if not st.session_state.get("authenticated", False):
        st.warning(" Please log in to access this page.")
        authentication_page()
        st.stop()


def require_admin_access():
    """
    Check if user has admin role. If not, show error and stop execution.
    Use this at the top of admin-only pages.
    """
    require_authentication()

    if st.session_state.get("user_role") != "admin":
        st.error(" You do not have permission to access this page.")
        st.info("This page is only accessible to administrators.")
        st.stop()


def get_current_user():
    """
    Get current user data from session state.

    Returns:
        dict: User data or None if not authenticated
    """
    if not st.session_state.get("authenticated", False):
        return None

    return {
        "id": st.session_state.get("user_id"),
        "email": st.session_state.get("user_email"),
        "first_name": st.session_state.get("user_first_name"),
        "last_name": st.session_state.get("user_last_name"),
        "role": st.session_state.get("user_role"),
        "access_token": st.session_state.get("access_token"),
        "refresh_token": st.session_state.get("refresh_token")
    }
