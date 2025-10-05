"""
Quick script to create admin account for Yazan.
"""
import sys
import os

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from utils.auth import signup_user

print("=" * 70)
print(" CREATING ADMIN ACCOUNT")
print("=" * 70)
print()

# Create admin account
success, message = signup_user(
    email="yazan@halton.com",  # You can change this email
    password="Halton2025!",     # CHANGE THIS PASSWORD AFTER FIRST LOGIN!
    first_name="Yazan",
    last_name="Admin",
    role="admin"
)

print()
if success:
    print("✅", message)
    print()
    print("=" * 70)
    print(" ADMIN ACCOUNT CREATED!")
    print("=" * 70)
    print()
    print("📧 Email:    yazan@halton.com")
    print("🔑 Password: Halton2025!")
    print("👤 Name:     Yazan Admin")
    print("🔐 Role:     ADMINISTRATOR")
    print()
    print("=" * 70)
    print()
    print("⚠️  IMPORTANT: Change your password after first login!")
    print()
    print("NEXT STEPS:")
    print("1. Run: streamlit run app.py")
    print("2. Login with the credentials above")
    print("3. Go to User Management → Invite team members")
    print()
    print("=" * 70)
else:
    print("❌", message)
    print()
    print("This might mean:")
    print("- Email already exists (try logging in)")
    print("- Database schema not set up")
    print("- Supabase connection issue")
    print()
    print("Try running: python test_supabase_connection.py")
