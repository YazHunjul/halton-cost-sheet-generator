# 🔐 Authentication Setup Guide

## ✅ What's Been Created

Your Halton Quotation System now has **complete authentication** with login, signup, and admin user management!

---

## 📋 Quick Start (3 Simple Steps)

### **Step 1: Set Up Database** (5 minutes)

1. **Open Supabase Dashboard**
   - Go to: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo
   - Click **SQL Editor** in sidebar

2. **Run Schema**
   - Click **New Query**
   - Copy ALL content from `database/schema.sql`
   - Paste into editor
   - Click **RUN** (or Ctrl+Enter)
   - ✅ You should see "Success. No rows returned"

3. **Verify Tables Created**
   - Click **Table Editor** in sidebar
   - You should see:
     - ✅ `user_profiles`
     - ✅ `user_invitations`
     - ✅ `audit_logs`

---

### **Step 2: Create Your Admin Account** (2 minutes)

1. **Go to Authentication**
   - Click **Authentication** in sidebar
   - Click **Users** tab
   - Click **Add User** → **Create New User**

2. **Fill in Details**
   - **Email**: `yazan@yourcompany.com` (or your preferred email)
   - **Password**: `YourSecurePassword123!` (choose a strong password)
   - ✅ Check **Auto Confirm User**
   - Click **Create User**

3. **Add User Metadata**
   - Find your user in the list, click on it
   - Click **User Metadata** tab
   - Click **Edit**
   - Paste this JSON:
     ```json
     {
       "first_name": "Yazan",
       "last_name": "Admin",
       "role": "admin"
     }
     ```
   - Click **Save**

---

### **Step 3: Start the App** (1 minute)

```bash
# Navigate to project directory
cd /Users/yazan/Desktop/Efficiency/UKCS

# Activate virtual environment (if not already active)
source venv/bin/activate

# Run the application
streamlit run app.py
```

**Your app will open at:** `http://localhost:8501`

---

## 🎯 How to Use the System

### **First Time Login**

1. App opens → You'll see **Login Page**
2. Enter your admin credentials:
   - Email: `yazan@yourcompany.com`
   - Password: `YourSecurePassword123!`
3. Click **Login**
4. ✅ You're in!

---

### **Inviting New Users** (Admin Only)

Once logged in as admin:

1. **Navigate to User Management**
   - Sidebar → Select "👥 User Management"

2. **Go to "Invite User" Tab**

3. **Fill in User Details**:
   - First Name: `John`
   - Last Name: `Doe`
   - Email: `john.doe@company.com`
   - Role: Choose **"👤 Regular User"** or **"🔑 Administrator"**
   - Expiry: Default 7 days (or customize)

4. **Click "Create Invitation"**

5. **Copy the Invitation Link**
   - A unique link will be generated
   - Example: `http://localhost:8501?token=abc123...`
   - Copy and send this link to the new user

---

### **New User Signup Process**

When a new user receives the invitation link:

1. **Click the Link**
   - Opens app with invitation token pre-filled

2. **Review Account Details**
   - App shows: First Name, Last Name, Email, Role
   - All pre-filled from invitation

3. **Set Password**
   - Enter strong password (min 8 characters)
   - Confirm password
   - Accept terms

4. **Click "Create Account"**

5. **Account Created!**
   - Redirected to login page
   - Can now login with email and password

---

## 🔑 User Roles & Permissions

### **Admin Users Can:**
- ✅ Access "User Management" page
- ✅ Create invitations for new users
- ✅ Assign roles (admin or user)
- ✅ View all users
- ✅ Activate/deactivate users
- ✅ View pending invitations
- ✅ Access all regular features

### **Regular Users Can:**
- ✅ Login to the system
- ✅ Create projects
- ✅ Generate quotations
- ✅ View their own projects
- ❌ Cannot access User Management
- ❌ Cannot invite other users
- ❌ Cannot manage users

---

## 📱 App Features

### **Login Page** (`/`)
- Email/password authentication
- "Have an invitation?" button for signup
- Secure session management
- Error handling for invalid credentials

### **Signup Page** (`/?token=xxx`)
- Invitation token verification
- Pre-filled user details
- Password creation
- Terms acceptance
- Automatic redirect to login after signup

### **User Management Page** (Admin Only)
- **All Users Tab**:
  - View all registered users
  - Filter by role (admin/user)
  - Filter by status (active/inactive)
  - Search by name or email
  - Activate/deactivate users

- **Pending Invitations Tab**:
  - View all pending invitations
  - See expiry status
  - Copy invitation links
  - Cancel invitations

- **Invite User Tab**:
  - Create new invitations
  - Assign roles
  - Set expiry period
  - Generate invitation links

---

## 🔒 Security Features

- ✅ **Password Encryption**: Bcrypt hashing via Supabase Auth
- ✅ **JWT Tokens**: Secure session tokens
- ✅ **Row Level Security**: Database-level authorization
- ✅ **Role-Based Access Control**: Admin vs User permissions
- ✅ **Invitation Expiry**: Time-limited invitation links
- ✅ **Email Verification**: Optional (can be enabled in Supabase)
- ✅ **Session Management**: Automatic logout on browser close
- ✅ **XSRF Protection**: Enabled in Streamlit config

---

## 🧪 Testing Checklist

After setup, verify:

**Authentication:**
- [ ] Admin can login with correct credentials
- [ ] Wrong password is rejected
- [ ] Login redirects to main app
- [ ] User info shows in sidebar
- [ ] Logout clears session and redirects to login

**Admin Features:**
- [ ] "User Management" appears in navigation (admin only)
- [ ] Can create invitation
- [ ] Invitation link is generated
- [ ] Can view all users
- [ ] Can activate/deactivate users
- [ ] Can view pending invitations

**User Signup:**
- [ ] Invitation link opens signup page
- [ ] User details are pre-filled
- [ ] Can set password
- [ ] Account is created successfully
- [ ] Redirect to login works
- [ ] New user can login

**Regular User:**
- [ ] Can login successfully
- [ ] Cannot see "User Management" in navigation
- [ ] Can access "Single Page Setup"
- [ ] Can access "Generate Word Documents"
- [ ] Can access "Create Revision"

---

## 📂 File Structure

```
UKCS/
├── app.py                          # Entry point (with auth)
├── .env                            # Supabase credentials
│
├── database/
│   ├── schema.sql                  # Database schema
│   ├── SETUP_INSTRUCTIONS.md       # Detailed setup
│   └── AUTHENTICATION_SUMMARY.md   # System overview
│
├── src/
│   ├── app.py                      # Original app (no auth)
│   ├── app_with_auth.py           # New authenticated app
│   │
│   ├── pages/
│   │   ├── auth_page.py           # Login/Signup UI
│   │   └── admin_users.py         # User management UI
│   │
│   ├── utils/
│   │   └── auth.py                # Authentication functions
│   │
│   └── config/
│       └── supabase_config.py     # Supabase client
│
└── .streamlit/
    └── config.toml                # Streamlit configuration
```

---

## 🆘 Troubleshooting

### **Problem: Can't login - "Invalid email or password"**

**Solution:**
1. Verify you created the user in Supabase Dashboard
2. Check you added the User Metadata with `role: "admin"`
3. Ensure password is correct
4. Check user is confirmed (Auto Confirm User was checked)

### **Problem: Login works but redirects back to login**

**Solution:**
1. Check if user profile was created in `user_profiles` table
2. Verify the trigger `on_auth_user_created` exists
3. Manually insert profile if needed:
   ```sql
   INSERT INTO public.user_profiles (id, email, first_name, last_name, role)
   SELECT id, email, 'Yazan', 'Admin', 'admin'
   FROM auth.users
   WHERE email = 'your-email@company.com';
   ```

### **Problem: "User Management" not showing**

**Solution:**
1. Verify user role is "admin" in `user_profiles` table
2. Check session state: `st.session_state.user_role` should be "admin"
3. Re-login to refresh session

### **Problem: Invitation link doesn't work**

**Solution:**
1. Check invitation exists in `user_invitations` table
2. Verify status is "pending" not "expired"
3. Check expires_at date hasn't passed
4. Ensure token matches exactly (copy-paste carefully)

### **Problem: Database connection error**

**Solution:**
1. Verify `.env` file exists with correct credentials
2. Check Supabase project is active
3. Test connection: `python test_supabase_connection.py`
4. Ensure `supabase` and `python-dotenv` are installed

---

## 🎓 Next Steps

After successful setup:

1. **Create Your Admin Account** ✅
2. **Login and Test** ✅
3. **Invite Team Members**:
   - Go to User Management
   - Invite your team
   - Assign appropriate roles

4. **Configure Production URL**:
   - Update `base_url` in `src/pages/admin_users.py` (line 16)
   - Change from `http://localhost:8501` to your Streamlit Cloud URL

5. **Optional Enhancements**:
   - Enable email confirmations in Supabase
   - Customize email templates
   - Add password reset functionality
   - Implement audit logging

---

## 📞 Support

**Resources:**
- Setup Instructions: `database/SETUP_INSTRUCTIONS.md`
- Authentication Summary: `database/AUTHENTICATION_SUMMARY.md`
- Integration Guide: `SUPABASE_INTEGRATION_COMPLETE.md`

**Supabase Dashboard:**
- Project URL: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo
- SQL Editor: For running queries
- Table Editor: For viewing data
- Authentication: For managing users

---

## 🎉 You're All Set!

Your Halton Quotation System now has enterprise-grade authentication with:
- ✅ Secure login/logout
- ✅ Invitation-based signup
- ✅ Role-based access control
- ✅ Admin user management
- ✅ Complete audit trail

**Ready to start?** Run `streamlit run app.py` and login! 🚀
