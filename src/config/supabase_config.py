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

# Try to import streamlit for secrets support (Streamlit Cloud)
try:
    import streamlit as st
    # Use Streamlit secrets if available (deployed on Streamlit Cloud)
    if hasattr(st, 'secrets') and 'supabase' in st.secrets:
        SUPABASE_URL = st.secrets["supabase"]["SUPABASE_URL"]
        SUPABASE_ANON_KEY = st.secrets["supabase"]["SUPABASE_ANON_KEY"]
        SUPABASE_SERVICE_ROLE_KEY = st.secrets["supabase"]["SUPABASE_SERVICE_ROLE_KEY"]
    else:
        # Fall back to environment variables (local .env)
        SUPABASE_URL = os.getenv("SUPABASE_URL")
        SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY")
        SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
except ImportError:
    # Streamlit not available, use environment variables
    SUPABASE_URL = os.getenv("SUPABASE_URL")
    SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY")
    SUPABASE_SERVICE_ROLE_KEY = os.getenv("SUPABASE_SERVICE_ROLE_KEY")

# Global client instances
_supabase_client: Optional[Client] = None
_supabase_admin_client: Optional[Client] = None


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

        # Try a simple operation to test connection
        # This will attempt to connect to Supabase
        result = {
            "status": "connected",
            "url": SUPABASE_URL,
            "message": "Successfully connected to Supabase"
        }

        return result
    except Exception as e:
        return {
            "status": "error",
            "url": SUPABASE_URL,
            "message": f"Failed to connect: {str(e)}"
        }


def reset_clients():
    """Reset client instances (useful for testing or reconnection)."""
    global _supabase_client, _supabase_admin_client
    _supabase_client = None
    _supabase_admin_client = None
