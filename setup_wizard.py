"""
Setup Wizard for Halton Quotation System Authentication.
Guides you through the complete setup process.
"""
import sys
import os

# Add src directory to path
src_dir = os.path.join(os.path.dirname(__file__), 'src')
sys.path.insert(0, src_dir)

def print_header(title):
    """Print formatted header."""
    print()
    print("=" * 80)
    print(f" {title}")
    print("=" * 80)
    print()

def print_step(number, title):
    """Print step header."""
    print()
    print(f"{'='*80}")
    print(f" STEP {number}: {title}")
    print(f"{'='*80}")
    print()

def test_connection():
    """Test Supabase connection."""
    from config.supabase_config import test_connection as test_conn
    result = test_conn()
    return result['status'] == 'connected', result['message']

def check_database_ready():
    """Check if database tables exist."""
    try:
        from config.supabase_config import get_supabase_client
        client = get_supabase_client(use_service_role=True)

        # Try to query user_profiles table
        response = client.table("user_profiles").select("id").limit(1).execute()
        return True, "Database tables found!"
    except Exception as e:
        error_msg = str(e).lower()
        if "relation" in error_msg and "does not exist" in error_msg:
            return False, "Database tables not found. Please run schema first."
        else:
            return False, f"Error checking database: {str(e)}"

def create_admin_user():
    """Create admin user."""
    from utils.auth import signup_user

    print("Enter admin account details:")
    print()

    first_name = input("  First Name: ").strip()
    last_name = input("  Last Name: ").strip()
    email = input("  Email Address: ").strip()
    password = input("  Password (min 8 characters): ").strip()

    if not all([first_name, last_name, email, password]):
        return False, "All fields are required!"

    if len(password) < 8:
        return False, "Password must be at least 8 characters!"

    success, message = signup_user(
        email=email,
        password=password,
        first_name=first_name,
        last_name=last_name,
        role="admin"
    )

    return success, message

def main():
    """Main setup wizard."""
    print_header("HALTON QUOTATION SYSTEM - SETUP WIZARD")

    print("This wizard will guide you through setting up authentication.")
    print()
    print("You will:")
    print("  1. Test Supabase connection")
    print("  2. Verify database schema")
    print("  3. Create your admin account")
    print()

    input("Press Enter to continue...")

    # Step 1: Test Connection
    print_step(1, "TEST SUPABASE CONNECTION")

    connected, message = test_connection()

    if connected:
        print(f"✅ {message}")
    else:
        print(f"❌ {message}")
        print()
        print("Please fix connection issues before continuing:")
        print("  1. Check .env file exists")
        print("  2. Verify credentials are correct")
        print("  3. Ensure internet connection is working")
        print()
        return

    # Step 2: Check Database
    print_step(2, "VERIFY DATABASE SCHEMA")

    print("Checking if database tables exist...")
    print()

    db_ready, db_message = check_database_ready()

    if db_ready:
        print(f"✅ {db_message}")
    else:
        print(f"❌ {db_message}")
        print()
        print("IMPORTANT: You must run the database schema first!")
        print()
        print("How to do this:")
        print("  1. Open: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo")
        print("  2. Click 'SQL Editor' in sidebar")
        print("  3. Click 'New Query'")
        print("  4. Copy ALL content from: database/schema.sql")
        print("  5. Paste into editor")
        print("  6. Click 'RUN' (or press Ctrl+Enter)")
        print("  7. Wait for success message")
        print("  8. Run this wizard again")
        print()
        return

    # Step 3: Create Admin
    print_step(3, "CREATE ADMIN ACCOUNT")

    success, message = create_admin_user()

    print()
    if success:
        print(f"✅ {message}")
        print()
        print_header("SETUP COMPLETE!")
        print("Your admin account has been created successfully!")
        print()
        print("NEXT STEPS:")
        print("  1. Run: streamlit run app.py")
        print("  2. Login with your credentials")
        print("  3. Navigate to 'User Management' to invite team members")
        print()
        print("=" * 80)
    else:
        print(f"❌ {message}")
        print()
        print("Please try again or create account manually in Supabase Dashboard.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print()
        print("Setup cancelled by user.")
    except Exception as e:
        print()
        print("=" * 80)
        print(" ERROR")
        print("=" * 80)
        print(str(e))
        print()
        import traceback
        traceback.print_exc()
