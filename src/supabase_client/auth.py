"""
Authentication handlers for Supabase.
"""
import logging
from typing import Dict, Any, Optional
from .client import get_supabase

logger = logging.getLogger(__name__)

class AuthenticationManager:
    """Handles authentication with Supabase."""

    @staticmethod
    def sign_up(email: str, password: str, metadata: Dict[str, Any] = None) -> Dict[str, Any]:
        """Sign up a new user."""
        try:
            supabase = get_supabase()
            result = supabase.auth.sign_up({
                "email": email,
                "password": password,
                "options": {
                    "data": metadata or {}
                }
            })
            return result.user
        except Exception as e:
            logger.error(f"Error during sign up: {str(e)}")
            raise

    @staticmethod
    def sign_in(email: str, password: str) -> Dict[str, Any]:
        """Sign in an existing user."""
        try:
            supabase = get_supabase()
            result = supabase.auth.sign_in_with_password({
                "email": email,
                "password": password
            })
            return result.user
        except Exception as e:
            logger.error(f"Error during sign in: {str(e)}")
            raise

    @staticmethod
    def sign_out() -> None:
        """Sign out the current user."""
        try:
            supabase = get_supabase()
            supabase.auth.sign_out()
        except Exception as e:
            logger.error(f"Error during sign out: {str(e)}")
            raise

    @staticmethod
    def get_current_user() -> Optional[Dict[str, Any]]:
        """Get the current authenticated user."""
        try:
            supabase = get_supabase()
            return supabase.auth.get_user().user
        except Exception as e:
            logger.error(f"Error getting current user: {str(e)}")
            return None

    @staticmethod
    def update_user(user_id: str, user_data: Dict[str, Any]) -> Dict[str, Any]:
        """Update user metadata."""
        try:
            supabase = get_supabase()
            result = supabase.auth.admin.update_user_by_id(
                user_id,
                {"user_metadata": user_data}
            )
            return result.user
        except Exception as e:
            logger.error(f"Error updating user: {str(e)}")
            raise

    @staticmethod
    def reset_password(email: str) -> None:
        """Send password reset email."""
        try:
            supabase = get_supabase()
            supabase.auth.reset_password_email(email)
        except Exception as e:
            logger.error(f"Error sending password reset: {str(e)}")
            raise 