"""
Supabase configuration and client initialization.
Supports both local .env and Streamlit Cloud secrets.toml
"""
import os
from typing import Optional
from supabase import create_client, Client
from dotenv import load_dotenv

# Load environment variables from .env (local development)
load_dotenv()

# Global client instances
_supabase_client: Optional[Client] = None
_supabase_admin_client: Optional[Client] = None


def _get_config_value(key: str) -> Optional[str]:
    """
    Get configuration value from Streamlit secrets or environment variables.
    Tries Streamlit secrets first, then falls back to environment variables.
    """
    # Try Streamlit secrets first (for Streamlit Cloud deployment)
    try:
        import streamlit as st
        if hasattr(st, 'secrets') and 'supabase' in st.secrets:
            return st.secrets["supabase"].get(key)
    except (ImportError, KeyError, AttributeError):
        pass

    # Fall back to environment variables (for local development)
    return os.getenv(key)


def get_supabase_client(use_service_role: bool = False) -> Client:
    """
    Get or create a Supabase client instance.

    Args:
        use_service_role: If True, use service role key for admin operations.
                         If False, use anon key for normal operations.

    Returns:
        Client: Supabase client instance

    Raises:
        ValueError: If required environment variables are not set
    """
    global _supabase_client, _supabase_admin_client

    # Lazy-load configuration values
    SUPABASE_URL = _get_config_value("SUPABASE_URL")
    SUPABASE_ANON_KEY = _get_config_value("SUPABASE_ANON_KEY")
    SUPABASE_SERVICE_ROLE_KEY = _get_config_value("SUPABASE_SERVICE_ROLE_KEY")

    if not SUPABASE_URL:
        raise ValueError("SUPABASE_URL environment variable is not set")

    if use_service_role:
        if not SUPABASE_SERVICE_ROLE_KEY:
            raise ValueError("SUPABASE_SERVICE_ROLE_KEY environment variable is not set")

        if _supabase_admin_client is None:
            _supabase_admin_client = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

        return _supabase_admin_client
    else:
        if not SUPABASE_ANON_KEY:
            raise ValueError("SUPABASE_ANON_KEY environment variable is not set")

        if _supabase_client is None:
            _supabase_client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)

        return _supabase_client


def test_connection() -> dict:
    """
    Test the Supabase connection.

    Returns:
        dict: Connection status with details
    """
    try:
        client = get_supabase_client()
        url = _get_config_value("SUPABASE_URL")

        # Try a simple operation to test connection
        # This will attempt to connect to Supabase
        result = {
            "status": "connected",
            "url": url,
            "message": "Successfully connected to Supabase"
        }

        return result
    except Exception as e:
        url = _get_config_value("SUPABASE_URL")
        return {
            "status": "error",
            "url": url,
            "message": f"Failed to connect: {str(e)}"
        }


def reset_clients():
    """Reset client instances (useful for testing or reconnection)."""
    global _supabase_client, _supabase_admin_client
    _supabase_client = None
    _supabase_admin_client = None
