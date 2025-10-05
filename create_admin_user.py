"""
Script to create the first admin user for Halton Quotation System.
Run this after setting up the database schema in Supabase.
"""
import sys
import os

# Add src directory to path
src_dir = os.path.join(os.path.dirname(__file__), 'src')
sys.path.insert(0, src_dir)

from utils.auth import signup_user
from config.supabase_config import get_supabase_client

def create_admin_account():
    """Create admin account with user input."""
    print("=" * 70)
    print(" CREATE ADMIN ACCOUNT - Halton Quotation System")
    print("=" * 70)
    print()

    # Get user details
    print("Enter admin account details:")
    print()

    first_name = input("First Name: ").strip()
    if not first_name:
        print("❌ First name is required!")
        return

    last_name = input("Last Name: ").strip()
    if not last_name:
        print("❌ Last name is required!")
        return

    email = input("Email Address: ").strip()
    if not email or "@" not in email:
        print("❌ Valid email address is required!")
        return

    password = input("Password (min 8 characters): ").strip()
    if len(password) < 8:
        print("❌ Password must be at least 8 characters!")
        return

    confirm_password = input("Confirm Password: ").strip()
    if password != confirm_password:
        print("❌ Passwords do not match!")
        return

    print()
    print("-" * 70)
    print("Creating admin account...")
    print("-" * 70)

    # Create admin user
    success, message = signup_user(
        email=email,
        password=password,
        first_name=first_name,
        last_name=last_name,
        role="admin"  # Important: Set role to admin
    )

    print()
    if success:
        print("✅ " + message)
        print()
        print("=" * 70)
        print(" ADMIN ACCOUNT CREATED SUCCESSFULLY!")
        print("=" * 70)
        print()
        print("📧 Email:", email)
        print("👤 Name:", f"{first_name} {last_name}")
        print("🔑 Role: ADMINISTRATOR")
        print()
        print("=" * 70)
        print()
        print("NEXT STEPS:")
        print("1. Check your email for verification link (if email confirmation enabled)")
        print("2. Run: streamlit run app.py")
        print("3. Login with your credentials")
        print("4. Start inviting team members!")
        print()
        print("=" * 70)
    else:
        print("❌ " + message)
        print()
        print("TROUBLESHOOTING:")
        print("1. Ensure database schema has been run in Supabase")
        print("2. Check that Supabase Auth is configured")
        print("3. Verify .env file has correct credentials")
        print("4. Try running: python test_supabase_connection.py")


def main():
    """Main function."""
    try:
        # Test connection first
        from config.supabase_config import test_connection

        print("Testing Supabase connection...")
        result = test_connection()

        if result['status'] != 'connected':
            print()
            print("❌ Cannot connect to Supabase!")
            print("   Message:", result['message'])
            print()
            print("Please ensure:")
            print("1. .env file exists with correct credentials")
            print("2. Supabase project is active")
            print("3. Internet connection is working")
            print()
            return

        print("✅ Connected to Supabase!")
        print()

        # Create admin account
        create_admin_account()

    except Exception as e:
        print()
        print("=" * 70)
        print(" ERROR")
        print("=" * 70)
        print(str(e))
        print()
        print("Please check:")
        print("1. Database schema has been run: database/schema.sql")
        print("2. .env file exists with Supabase credentials")
        print("3. Dependencies installed: pip install -r requirements.txt")
        print()
        import traceback
        print("Full error:")
        traceback.print_exc()


if __name__ == "__main__":
    main()
