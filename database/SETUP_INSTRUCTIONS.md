# Supabase Authentication Setup Instructions

## Overview
This guide will help you set up user authentication for the Halton Quotation System with two user types: **Admin** and **Regular User**.

---

## Step 1: Run the Database Schema

1. **Open Supabase Dashboard**
   - Go to: https://supabase.com/dashboard
   - Select your project: `rlvtnyotgsgywxkshazo`

2. **Navigate to SQL Editor**
   - Click on the **SQL Editor** icon in the left sidebar
   - Click **New Query**

3. **Execute Schema**
   - Copy the entire contents of `database/schema.sql`
   - Paste into the SQL Editor
   - Click **Run** or press `Ctrl+Enter`

4. **Verify Tables Created**
   - Go to **Table Editor** in the left sidebar
   - You should see these new tables:
     - `user_profiles`
     - `user_invitations`
     - `audit_logs`

---

## Step 2: Configure Supabase Authentication

### Enable Email Authentication

1. **Go to Authentication Settings**
   - Click **Authentication** in left sidebar
   - Click **Providers** tab

2. **Enable Email Provider**
   - Find "Email" provider
   - Toggle it **ON**
   - Configure settings:
     - ✅ Enable Email Confirmations (recommended)
     - ✅ Enable Email Signup
     - Set confirmation email template (optional)

3. **Configure Email Templates** (Optional but Recommended)
   - Click **Email Templates** tab
   - Customize:
     - Confirmation email
     - Invitation email
     - Password reset email
     - Magic link email

4. **Set Site URL**
   - Go to **Settings** → **General**
   - Set **Site URL** to your application URL:
     - Development: `http://localhost:8501`
     - Production: Your Streamlit Cloud URL
   - Add to **Redirect URLs**:
     - `http://localhost:8501`
     - Your production URL

---

## Step 3: Create Your First Admin User

You have **two options** to create the first admin user:

### Option A: Through Supabase Dashboard (Recommended)

1. **Create User in Auth**
   - Go to **Authentication** → **Users**
   - Click **Add User** → **Create New User**
   - Enter:
     - Email: `your-admin@email.com`
     - Password: `SecurePassword123!`
     - ✅ Auto Confirm User

2. **Add User Metadata**
   - After creating user, click on the user
   - Click **User Metadata** tab
   - Add metadata (click **Edit**):
     ```json
     {
       "first_name": "Admin",
       "last_name": "User",
       "role": "admin"
     }
     ```
   - Click **Save**

3. **Verify Profile Created**
   - Go to **Table Editor** → `user_profiles`
   - You should see the admin user profile
   - If not, the trigger will create it on first login

### Option B: Using SQL Editor

```sql
-- Create auth user
INSERT INTO auth.users (
    instance_id,
    id,
    aud,
    role,
    email,
    encrypted_password,
    email_confirmed_at,
    raw_user_meta_data,
    created_at,
    updated_at
) VALUES (
    '00000000-0000-0000-0000-000000000000',
    gen_random_uuid(),
    'authenticated',
    'authenticated',
    'admin@yourcompany.com',
    crypt('your-secure-password', gen_salt('bf')),
    NOW(),
    '{"first_name": "Admin", "last_name": "User", "role": "admin"}'::jsonb,
    NOW(),
    NOW()
);

-- The trigger will automatically create the user_profile
```

---

## Step 4: Test Authentication

### Test Connection Script

Run the test script to verify everything is working:

```bash
# Activate virtual environment
source venv/bin/activate

# Create test script
python -c "
from src.config.supabase_config import test_connection
result = test_connection()
print(f'Status: {result[\"status\"]}')
print(f'Message: {result[\"message\"]}')
"
```

### Test Login (After Creating UI)

1. Start the Streamlit app
2. Navigate to login page
3. Enter admin credentials
4. Verify successful login and admin access

---

## Step 5: Understanding the User System

### User Roles

| Role | Capabilities |
|------|-------------|
| **admin** | • Create/manage users<br>• Send invitations<br>• View all projects<br>• Manage system settings<br>• Access admin panel |
| **user** | • Create projects<br>• View own projects<br>• Generate quotations<br>• Update own profile |

### User Invitation Flow

1. **Admin Creates Invitation**
   - Specifies: email, first name, last name, role
   - System generates secure invitation token
   - Invitation expires in 7 days (configurable)

2. **Invitation Link Generated**
   - Format: `https://your-app.com/accept-invite?token=XXXXX`
   - Admin shares link with invitee

3. **User Accepts Invitation**
   - Clicks invitation link
   - Sets password
   - Account automatically created with pre-assigned role

4. **User Can Login**
   - Uses email and password
   - Access granted based on role

---

## Step 6: Security Features Enabled

### Row Level Security (RLS)

All tables have RLS enabled with policies:

- ✅ **user_profiles**: Users can only view/edit own profile, admins can manage all
- ✅ **user_invitations**: Only admins can create/view invitations
- ✅ **audit_logs**: Only admins can view logs

### Automatic Features

- ✅ **Auto Profile Creation**: Trigger creates profile when user signs up
- ✅ **Last Login Tracking**: Updates on each successful login
- ✅ **Password Hashing**: Supabase Auth handles secure password storage
- ✅ **Email Verification**: Optional but recommended
- ✅ **Token Expiry**: JWT tokens expire and refresh automatically

---

## Step 7: Environment Variables

Ensure your `.env` file has the correct credentials:

```env
SUPABASE_URL=https://rlvtnyotgsgywxkshazo.supabase.co
SUPABASE_ANON_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...
SUPABASE_SERVICE_ROLE_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...
```

**⚠️ Important**:
- `.env` is in `.gitignore` - do NOT commit credentials
- Use `.env.example` as template for other developers

---

## Next Steps

1. ✅ Run database schema
2. ✅ Configure Supabase Auth
3. ✅ Create first admin user
4. 🔄 Build authentication UI (login, signup, invitation pages)
5. 🔄 Integrate auth into existing Streamlit app
6. 🔄 Test complete user flow

---

## Troubleshooting

### Issue: User profile not created after signup

**Solution**: Check if the trigger is active:
```sql
SELECT * FROM pg_trigger WHERE tgname = 'on_auth_user_created';
```

If not found, re-run the schema.

### Issue: RLS blocking operations

**Solution**: Verify policies are active:
```sql
SELECT schemaname, tablename, policyname
FROM pg_policies
WHERE schemaname = 'public';
```

### Issue: Invitation email not sending

**Solution**:
- Check SMTP settings in Supabase dashboard
- Verify email templates are configured
- Check if email confirmations are enabled

---

## API Reference

### Authentication Functions

```python
from utils.auth import (
    signup_user,           # Create new user
    login_user,           # Authenticate user
    logout_user,          # End session
    create_invitation,    # Send invitation (admin only)
    verify_invitation,    # Check invitation validity
    accept_invitation,    # Accept invite and create account
    get_user_profile,     # Get user data
    update_user_profile,  # Update user info
    is_admin,            # Check if user is admin
    require_auth,        # Decorator for protected routes
    require_admin        # Decorator for admin-only routes
)
```

### Example Usage

```python
# Login user
success, user_data, message = login_user("user@email.com", "password")
if success:
    st.session_state.user_id = user_data["id"]
    st.session_state.user_role = user_data["role"]
    st.session_state.authenticated = True

# Create invitation (admin only)
success, token, message = create_invitation(
    email="newuser@email.com",
    first_name="John",
    last_name="Doe",
    role="user",
    invited_by_id=st.session_state.user_id
)

# Generate invitation link
invite_link = f"https://your-app.com/accept-invite?token={token}"
```

---

## Support

For issues or questions:
1. Check Supabase logs: **Logs & Analytics** in dashboard
2. Review table data in **Table Editor**
3. Test queries in **SQL Editor**
4. Check authentication events in **Authentication** → **Logs**
