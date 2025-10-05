"""
Admin user management page.
Only accessible to users with admin role.
"""
import streamlit as st
import pandas as pd
from datetime import datetime
from pages.auth_page import require_admin_access, get_current_user
from utils.auth import create_invitation, get_user_profile, update_user_profile, delete_user
from config.supabase_config import get_supabase_client


def generate_invitation_link(token: str) -> str:
    """Generate invitation link with token."""
    # Get the current app URL
    # In production, this should be your actual domain
    base_url = "http://localhost:8501"  # Update this for production
    return f"{base_url}?token={token}"


def show_create_invitation_form():
    """Form to create new user invitations."""
    st.markdown("### + Invite New User")

    with st.form("create_invitation_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            first_name = st.text_input(
                "First Name *",
                placeholder="John"
            )
            email = st.text_input(
                "Email Address *",
                placeholder="user@company.com"
            )

        with col2:
            last_name = st.text_input(
                "Last Name *",
                placeholder="Doe"
            )
            role = st.selectbox(
                "User Role *",
                options=["user", "admin"],
                format_func=lambda x: "• Regular User" if x == "user" else "★ Administrator"
            )

        expiry_days = st.slider(
            "Invitation Expiry (days)",
            min_value=1,
            max_value=30,
            value=7,
            help="Number of days until the invitation expires"
        )

        st.markdown("---")
        submit_button = st.form_submit_button(
            "+ Create Invitation",
            type="primary",
            use_container_width=True
        )

    if submit_button:
        # Validation
        if not all([first_name, last_name, email]):
            st.error(" Please fill in all required fields.")
            return

        # Email validation
        if "@" not in email or "." not in email:
            st.error(" Please enter a valid email address.")
            return

        # Get current user
        current_user = get_current_user()

        # Create invitation
        with st.spinner("Creating invitation..."):
            success, token, message = create_invitation(
                email=email,
                first_name=first_name,
                last_name=last_name,
                role=role,
                invited_by_id=current_user["id"],
                expiry_days=expiry_days
            )

        if success:
            st.success(f" {message}")

            # Generate invitation link
            invitation_link = generate_invitation_link(token)

            # Display invitation details
            st.markdown("---")
            st.markdown("####  Invitation Created Successfully!")

            # Show invitation link
            st.text_input(
                "Invitation Link (copy and send to user):",
                value=invitation_link,
                key="invitation_link_display",
                help="Copy this link and send it to the user"
            )

            # Show invitation details
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("User", f"{first_name} {last_name}")
            with col2:
                st.metric("Role", role.upper())
            with col3:
                st.metric("Expires In", f"{expiry_days} days")

            st.info(" **Next Steps:** Copy the invitation link above and send it to the user via email or messaging app.")

        else:
            st.error(f" {message}")


def show_user_list():
    """Display list of all users."""
    st.markdown("### ◆ User Management")

    try:
        # Use service role to bypass RLS and see all users
        client = get_supabase_client(use_service_role=True)

        # Fetch all user profiles
        response = client.table("user_profiles").select("*").order("created_at", desc=True).execute()

        if not response.data:
            st.info("No users found.")
            return

        # Create DataFrame
        users_df = pd.DataFrame(response.data)

        # Display summary metrics
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_users = len(users_df)
            st.metric("Total Users", total_users)

        with col2:
            admin_count = len(users_df[users_df["role"] == "admin"])
            st.metric("Admins", admin_count)

        with col3:
            active_count = len(users_df[users_df["is_active"] == True])
            st.metric("Active", active_count)

        with col4:
            inactive_count = len(users_df[users_df["is_active"] == False])
            st.metric("Inactive", inactive_count)

        st.markdown("---")

        # Filter options
        col1, col2, col3 = st.columns(3)

        with col1:
            role_filter = st.selectbox(
                "Filter by Role",
                options=["All", "admin", "user"],
                format_func=lambda x: x if x == "All" else x.upper()
            )

        with col2:
            status_filter = st.selectbox(
                "Filter by Status",
                options=["All", "Active", "Inactive"]
            )

        with col3:
            search_term = st.text_input(
                "Search",
                placeholder="Name or email..."
            )

        # Apply filters
        filtered_df = users_df.copy()

        if role_filter != "All":
            filtered_df = filtered_df[filtered_df["role"] == role_filter]

        if status_filter == "Active":
            filtered_df = filtered_df[filtered_df["is_active"] == True]
        elif status_filter == "Inactive":
            filtered_df = filtered_df[filtered_df["is_active"] == False]

        if search_term:
            mask = (
                filtered_df["first_name"].str.contains(search_term, case=False, na=False) |
                filtered_df["last_name"].str.contains(search_term, case=False, na=False) |
                filtered_df["email"].str.contains(search_term, case=False, na=False)
            )
            filtered_df = filtered_df[mask]

        st.markdown(f"**Showing {len(filtered_df)} of {len(users_df)} users**")

        # Display users
        for idx, user in filtered_df.iterrows():
            with st.expander(
                f"{'' if user['role'] == 'admin' else ''} {user['first_name']} {user['last_name']} - {user['email']}",
                expanded=False
            ):
                col1, col2 = st.columns(2)

                with col1:
                    st.write("**Email:**", user["email"])
                    st.write("**Role:**", user["role"].upper())
                    st.write("**Status:**", " Active" if user["is_active"] else " Inactive")

                with col2:
                    created_date = datetime.fromisoformat(user["created_at"].replace("Z", "+00:00"))
                    st.write("**Created:**", created_date.strftime("%Y-%m-%d %H:%M"))

                    if user.get("last_login"):
                        last_login = datetime.fromisoformat(user["last_login"].replace("Z", "+00:00"))
                        st.write("**Last Login:**", last_login.strftime("%Y-%m-%d %H:%M"))
                    else:
                        st.write("**Last Login:**", "Never")

                st.markdown("---")

                # Admin actions
                action_col1, action_col2, action_col3 = st.columns(3)

                current_user = get_current_user()

                # Prevent admin from modifying themselves
                if user["id"] != current_user["id"]:
                    with action_col1:
                        if user["is_active"]:
                            if st.button(f"✕ Deactivate", key=f"deactivate_{user['id']}"):
                                # Deactivate user
                                success, message = update_user_profile(
                                    user["id"],
                                    {"is_active": False}
                                )
                                if success:
                                    st.success("User deactivated")
                                    st.rerun()
                                else:
                                    st.error(message)
                        else:
                            if st.button(f"✓ Activate", key=f"activate_{user['id']}"):
                                # Activate user
                                success, message = update_user_profile(
                                    user["id"],
                                    {"is_active": True}
                                )
                                if success:
                                    st.success("User activated")
                                    st.rerun()
                                else:
                                    st.error(message)

                    with action_col2:
                        if st.button(f"⊗ Delete User", key=f"delete_{user['id']}", type="secondary"):
                            # Show confirmation dialog
                            if f"confirm_delete_{user['id']}" not in st.session_state:
                                st.session_state[f"confirm_delete_{user['id']}"] = True
                                st.warning("⚠️ Click 'Confirm Delete' to permanently delete this user.")
                                st.rerun()

                    # Show confirmation button if delete was clicked
                    if st.session_state.get(f"confirm_delete_{user['id']}", False):
                        with action_col3:
                            if st.button(f"✓ Confirm Delete", key=f"confirm_delete_btn_{user['id']}", type="primary"):
                                # Delete user
                                success, message = delete_user(user["id"])
                                if success:
                                    st.success("User deleted successfully")
                                    # Clear confirmation state
                                    del st.session_state[f"confirm_delete_{user['id']}"]
                                    st.rerun()
                                else:
                                    st.error(message)
                                    del st.session_state[f"confirm_delete_{user['id']}"]

                            if st.button(f"✕ Cancel", key=f"cancel_delete_{user['id']}"):
                                # Cancel delete
                                del st.session_state[f"confirm_delete_{user['id']}"]
                                st.rerun()
                else:
                    with action_col1:
                        st.write("**(You)**")

    except Exception as e:
        st.error(f"Error loading users: {str(e)}")


def show_pending_invitations():
    """Display list of pending invitations."""
    st.markdown("### ◆ Pending Invitations")

    try:
        # Use service role to see all invitations
        client = get_supabase_client(use_service_role=True)

        # Fetch pending invitations
        response = client.table("user_invitations").select("*").eq("status", "pending").order("created_at", desc=True).execute()

        if not response.data:
            st.info("No pending invitations.")
            return

        # Display count
        st.metric("Pending Invitations", len(response.data))

        # Display invitations
        for invitation in response.data:
            expires_at = datetime.fromisoformat(invitation["expires_at"].replace("Z", "+00:00"))
            is_expired = datetime.now(expires_at.tzinfo) > expires_at

            status_emoji = "⏰" if is_expired else ""

            with st.expander(
                f"{status_emoji} {invitation['first_name']} {invitation['last_name']} - {invitation['email']}",
                expanded=False
            ):
                col1, col2 = st.columns(2)

                with col1:
                    st.write("**Email:**", invitation["email"])
                    st.write("**Role:**", invitation["role"].upper())
                    created_date = datetime.fromisoformat(invitation["created_at"].replace("Z", "+00:00"))
                    st.write("**Sent:**", created_date.strftime("%Y-%m-%d %H:%M"))

                with col2:
                    st.write("**Status:**", " Expired" if is_expired else "🟢 Active")
                    st.write("**Expires:**", expires_at.strftime("%Y-%m-%d %H:%M"))

                # Show invitation link if not expired
                if not is_expired:
                    invitation_link = generate_invitation_link(invitation["invitation_token"])
                    st.text_input(
                        "Invitation Link:",
                        value=invitation_link,
                        key=f"inv_link_{invitation['id']}",
                        help="Copy and send this link to the user"
                    )

                # Cancel button
                if st.button(f"✕ Cancel Invitation", key=f"cancel_{invitation['id']}"):
                    try:
                        cancel_client = get_supabase_client(use_service_role=True)
                        cancel_client.table("user_invitations").update(
                            {"status": "cancelled"}
                        ).eq("id", invitation["id"]).execute()

                        st.success("Invitation cancelled")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error cancelling invitation: {str(e)}")

    except Exception as e:
        st.error(f"Error loading invitations: {str(e)}")


def admin_users_page():
    """Main admin users management page."""

    # Require admin access
    require_admin_access()

    # Page header
    st.title("◆ User Management")
    st.markdown("Manage users, send invitations, and control access to the system.")
    st.markdown("---")

    # Create tabs
    tab1, tab2, tab3 = st.tabs(["▸ All Users", "▸ Pending Invitations", "+ Invite User"])

    with tab1:
        show_user_list()

    with tab2:
        show_pending_invitations()

    with tab3:
        show_create_invitation_form()
