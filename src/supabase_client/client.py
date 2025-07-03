"""
Supabase client configuration and initialization.
"""
import os
from dotenv import load_dotenv
from supabase import create_client, Client

# Load environment variables
load_dotenv()

# Initialize Supabase client with direct credentials
# In production, these should be in environment variables
supabase: Client = create_client(
    supabase_url="https://yiqfwblohscmuiyyrkex.supabase.co",
    supabase_key="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlpcWZ3YmxvaHNjbXVpeXlya2V4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTE0Nzk0NTYsImV4cCI6MjA2NzA1NTQ1Nn0.BThsNI5P8gAY1vf2_DuLdiIPat2r9NBspEpJEcYrjWQ"
)

def get_supabase() -> Client:
    """
    Get the Supabase client instance.
    Returns:
        Client: Configured Supabase client
    """
    return supabase 