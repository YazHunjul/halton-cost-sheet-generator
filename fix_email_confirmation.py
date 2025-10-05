"""
Fix email confirmation issue by confirming the user manually.
"""
import sys
import os

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from config.supabase_config import get_supabase_client

print("=" * 70)
print(" FIX EMAIL CONFIRMATION")
print("=" * 70)
print()
print("Confirming admin user email...")
print()

try:
    # Use service role to bypass RLS
    client = get_supabase_client(use_service_role=True)

    # Update the user to confirm email
    # We need to use Supabase Admin API
    email = "yazan@halton.com"

    # Get user by email
    response = client.auth.admin.list_users()

    user_found = False
    for user in response:
        if user.email == email:
            user_found = True
            user_id = user.id

            # Confirm email using admin API
            client.auth.admin.update_user_by_id(
                user_id,
                {"email_confirm": True}
            )

            print(f"✅ Email confirmed for: {email}")
            print()
            print("You can now login without email confirmation!")
            break

    if not user_found:
        print(f"❌ User not found: {email}")
        print()
        print("The user might not have been created yet.")
        print("Try running: python create_my_admin.py")

except Exception as e:
    error_msg = str(e)

    if "Admin API" in error_msg or "admin" in error_msg.lower():
        print("❌ Admin API not accessible with current permissions")
        print()
        print("SOLUTION: Manually confirm email in Supabase Dashboard")
        print()
        print("How to fix:")
        print("1. Go to: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo")
        print("2. Click 'Authentication' → 'Users'")
        print("3. Find user: yazan@halton.com")
        print("4. Click on the user")
        print("5. Look for 'Email Confirmed' status")
        print("6. If not confirmed, you can:")
        print("   - Delete this user and recreate with 'Auto Confirm User' checked")
        print("   - Or wait for actual email confirmation")
        print()
    else:
        print(f"❌ Error: {error_msg}")
        print()
        print("Try the manual solution above.")

print()
print("=" * 70)
print()
print("ALTERNATIVE: Disable Email Confirmation Entirely")
print()
print("1. Go to: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo")
print("2. Click 'Authentication' → 'Providers'")
print("3. Find 'Email' provider")
print("4. Click to edit")
print("5. UNCHECK 'Enable email confirmations'")
print("6. Save")
print("7. Delete and recreate the user with: python create_my_admin.py")
print()
print("=" * 70)
