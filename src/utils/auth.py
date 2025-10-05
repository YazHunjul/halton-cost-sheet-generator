"""
Authentication utilities for Halton Quotation System.
Handles user authentication, registration, and role management using Supabase Auth.
"""
import secrets
import streamlit as st
from typing import Optional, Dict, Tuple
from datetime import datetime, timedelta
from config.supabase_config import get_supabase_client


class AuthError(Exception):
    """Custom exception for authentication errors."""
    pass


def signup_user(
    email: str,
    password: str,
    first_name: str,
    last_name: str,
    role: str = "user"
) -> Tuple[bool, str]:
    """
    Create a new user account.

    Args:
        email: User's email address
        password: User's password
        first_name: User's first name
        last_name: User's last name
        role: User role ('admin' or 'user')

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        client = get_supabase_client()

        # Sign up user with metadata
        response = client.auth.sign_up({
            "email": email,
            "password": password,
            "options": {
                "data": {
                    "first_name": first_name,
                    "last_name": last_name,
                    "role": role
                }
            }
        })

        if response.user:
            return True, "User account created successfully! Please check your email to verify your account."
        else:
            return False, "Failed to create user account."

    except Exception as e:
        return False, f"Signup error: {str(e)}"


def login_user(email: str, password: str) -> Tuple[bool, Optional[Dict], str]:
    """
    Authenticate user and create session.

    Args:
        email: User's email address
        password: User's password

    Returns:
        Tuple of (success: bool, user_data: Optional[Dict], message: str)
    """
    try:
        client = get_supabase_client()

        # Sign in user
        response = client.auth.sign_in_with_password({
            "email": email,
            "password": password
        })

        if response.user:
            # Get user profile with role
            profile_response = client.table("user_profiles").select("*").eq("id", response.user.id).execute()

            if profile_response.data and len(profile_response.data) > 0:
                profile = profile_response.data[0]

                # Check if user is active
                if not profile.get("is_active", True):
                    client.auth.sign_out()
                    return False, None, "Your account has been deactivated. Please contact an administrator."

                # Update last login
                client.rpc("update_last_login", {"user_id": response.user.id}).execute()

                user_data = {
                    "id": response.user.id,
                    "email": response.user.email,
                    "first_name": profile.get("first_name", ""),
                    "last_name": profile.get("last_name", ""),
                    "role": profile.get("role", "user"),
                    "access_token": response.session.access_token,
                    "refresh_token": response.session.refresh_token
                }

                return True, user_data, "Login successful!"
            else:
                return False, None, "User profile not found."
        else:
            return False, None, "Invalid email or password."

    except Exception as e:
        return False, None, f"Login error: {str(e)}"


def logout_user() -> Tuple[bool, str]:
    """
    Log out current user and clear session.

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        client = get_supabase_client()
        client.auth.sign_out()

        # Clear Streamlit session state
        for key in list(st.session_state.keys()):
            if key.startswith("user_") or key == "authenticated":
                del st.session_state[key]

        return True, "Logged out successfully!"

    except Exception as e:
        return False, f"Logout error: {str(e)}"


def create_invitation(
    email: str,
    first_name: str,
    last_name: str,
    role: str,
    invited_by_id: str,
    expiry_days: int = 7
) -> Tuple[bool, Optional[str], str]:
    """
    Create a user invitation.

    Args:
        email: Invitee's email address
        first_name: Invitee's first name
        last_name: Invitee's last name
        role: User role ('admin' or 'user')
        invited_by_id: UUID of the admin creating the invitation
        expiry_days: Number of days until invitation expires

    Returns:
        Tuple of (success: bool, invitation_token: Optional[str], message: str)
    """
    try:
        client = get_supabase_client(use_service_role=True)

        # Generate secure token
        invitation_token = secrets.token_urlsafe(32)

        # Calculate expiry
        expires_at = datetime.now() + timedelta(days=expiry_days)

        # Create invitation record
        invitation_data = {
            "email": email,
            "first_name": first_name,
            "last_name": last_name,
            "role": role,
            "invited_by": invited_by_id,
            "invitation_token": invitation_token,
            "expires_at": expires_at.isoformat(),
            "status": "pending"
        }

        response = client.table("user_invitations").insert(invitation_data).execute()

        if response.data:
            return True, invitation_token, "Invitation created successfully!"
        else:
            return False, None, "Failed to create invitation."

    except Exception as e:
        return False, None, f"Invitation error: {str(e)}"


def verify_invitation(invitation_token: str) -> Tuple[bool, Optional[Dict], str]:
    """
    Verify an invitation token and retrieve invitation details.

    Args:
        invitation_token: The invitation token to verify

    Returns:
        Tuple of (valid: bool, invitation_data: Optional[Dict], message: str)
    """
    try:
        client = get_supabase_client()

        # Fetch invitation
        response = client.table("user_invitations").select("*").eq("invitation_token", invitation_token).execute()

        if not response.data or len(response.data) == 0:
            return False, None, "Invalid invitation token."

        invitation = response.data[0]

        # Check status
        if invitation["status"] != "pending":
            return False, None, f"This invitation has already been {invitation['status']}."

        # Check expiry
        expires_at = datetime.fromisoformat(invitation["expires_at"].replace("Z", "+00:00"))
        if datetime.now(expires_at.tzinfo) > expires_at:
            # Update status to expired
            client.table("user_invitations").update({"status": "expired"}).eq("id", invitation["id"]).execute()
            return False, None, "This invitation has expired."

        return True, invitation, "Invitation is valid."

    except Exception as e:
        return False, None, f"Verification error: {str(e)}"


def accept_invitation(invitation_token: str, password: str) -> Tuple[bool, str]:
    """
    Accept an invitation and create user account.

    Args:
        invitation_token: The invitation token
        password: User's chosen password

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Verify invitation
        valid, invitation, message = verify_invitation(invitation_token)
        if not valid:
            return False, message

        # Create user account
        success, signup_message = signup_user(
            email=invitation["email"],
            password=password,
            first_name=invitation["first_name"],
            last_name=invitation["last_name"],
            role=invitation["role"]
        )

        if success:
            # Mark invitation as accepted
            client = get_supabase_client(use_service_role=True)
            client.table("user_invitations").update({
                "status": "accepted",
                "accepted_at": datetime.now().isoformat()
            }).eq("invitation_token", invitation_token).execute()

            return True, "Account created successfully! You can now log in."
        else:
            return False, signup_message

    except Exception as e:
        return False, f"Accept invitation error: {str(e)}"


def get_user_profile(user_id: str) -> Optional[Dict]:
    """
    Get user profile by user ID.

    Args:
        user_id: User's UUID

    Returns:
        User profile dictionary or None
    """
    try:
        client = get_supabase_client()
        response = client.table("user_profiles").select("*").eq("id", user_id).execute()

        if response.data and len(response.data) > 0:
            return response.data[0]
        return None

    except Exception as e:
        st.error(f"Error fetching user profile: {str(e)}")
        return None


def update_user_profile(user_id: str, updates: Dict) -> Tuple[bool, str]:
    """
    Update user profile.

    Args:
        user_id: User's UUID
        updates: Dictionary of fields to update

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Use service role to bypass RLS for admin operations
        client = get_supabase_client(use_service_role=True)

        # Don't allow changing role through this function (admin only)
        if "role" in updates and "id" in updates:
            del updates["role"]

        response = client.table("user_profiles").update(updates).eq("id", user_id).execute()

        if response.data:
            return True, "Profile updated successfully!"
        else:
            return False, "Failed to update profile."

    except Exception as e:
        return False, f"Update error: {str(e)}"


def delete_user(user_id: str) -> Tuple[bool, str]:
    """
    Delete a user account and profile.

    Args:
        user_id: User's UUID

    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Use service role to bypass RLS
        client = get_supabase_client(use_service_role=True)

        # Delete user profile (this will cascade delete related data)
        profile_response = client.table("user_profiles").delete().eq("id", user_id).execute()

        if profile_response.data:
            # Delete user from auth
            auth_response = client.auth.admin.delete_user(user_id)
            return True, "User deleted successfully!"
        else:
            return False, "Failed to delete user profile."

    except Exception as e:
        return False, f"Delete error: {str(e)}"


def is_admin(user_id: str) -> bool:
    """
    Check if user has admin role.

    Args:
        user_id: User's UUID

    Returns:
        True if user is admin, False otherwise
    """
    profile = get_user_profile(user_id)
    return profile.get("role") == "admin" if profile else False


def require_auth(func):
    """
    Decorator to require authentication for a function.
    Use with Streamlit pages to protect routes.
    """
    def wrapper(*args, **kwargs):
        if not st.session_state.get("authenticated", False):
            st.error("⚠️ You must be logged in to access this page.")
            st.stop()
        return func(*args, **kwargs)
    return wrapper


def require_admin(func):
    """
    Decorator to require admin role for a function.
    Use with Streamlit pages to protect admin routes.
    """
    def wrapper(*args, **kwargs):
        if not st.session_state.get("authenticated", False):
            st.error("⚠️ You must be logged in to access this page.")
            st.stop()

        if st.session_state.get("user_role") != "admin":
            st.error("⚠️ You must be an administrator to access this page.")
            st.stop()

        return func(*args, **kwargs)
    return wrapper
