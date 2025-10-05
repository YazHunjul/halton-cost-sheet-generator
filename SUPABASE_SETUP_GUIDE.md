# Supabase Project Setup Guide

Complete guide for setting up your new Supabase project for the Halton Quotation System.

## Project Details

- **URL**: https://vxvjncrolwyvvykalirw.supabase.co
- **Region**: Closer to your client
- **Status**: Ready for setup

## Setup Steps

### Step 1: Run Database Schema

1. Go to your Supabase Dashboard: https://vxvjncrolwyvvykalirw.supabase.co
2. Navigate to **SQL Editor** in the left sidebar
3. Click **+ New Query**
4. Copy and paste the contents of `database/schema.sql`
5. Click **Run** or press `Ctrl+Enter`
6. Wait for success message

### Step 2: Add Companies Schema

1. In SQL Editor, click **+ New Query**
2. Copy and paste the contents of `database/companies_schema.sql`
3. Click **Run**
4. This will create the companies table AND import all 75+ companies automatically

### Step 3: Create Storage Bucket for Templates

1. Navigate to **Storage** in the left sidebar
2. Click **Create a new bucket**
3. Name: `templates`
4. Set to **Private** (not public)
5. Click **Create bucket**

### Step 4: Set Storage Bucket Policies

1. Click on the `templates` bucket
2. Go to **Policies** tab
3. Click **New Policy**
4. Create a policy for authenticated users to read:

```sql
-- Policy name: Allow authenticated users to read templates
-- Allowed operation: SELECT
-- Target roles: authenticated

CREATE POLICY "Authenticated users can read templates"
ON storage.objects
FOR SELECT
TO authenticated
USING (bucket_id = 'templates');
```

5. Create a policy for service role to do everything:

```sql
-- Policy name: Service role full access
-- Allowed operation: ALL
-- Target roles: service_role

CREATE POLICY "Service role full access to templates"
ON storage.objects
FOR ALL
TO service_role
USING (bucket_id = 'templates');
```

### Step 5: Create First Admin User

**Option A: Via Supabase Dashboard (Recommended)**

1. Navigate to **Authentication** → **Users**
2. Click **Add user**
3. Choose **Create new user**
4. Enter email and password
5. Click **Create user**
6. Copy the user's UUID
7. Go to **SQL Editor** and run:

```sql
-- Replace 'USER_UUID_HERE' with the actual UUID from step 6
UPDATE public.user_profiles
SET role = 'admin'
WHERE id = 'USER_UUID_HERE';
```

**Option B: Via SQL (All-in-one)**

1. Go to **SQL Editor**
2. Run this query (replace email and password):

```sql
-- This creates a user in auth.users
-- The trigger will automatically create the user_profiles entry
-- Then we update the role to admin

DO $$
DECLARE
    user_id UUID;
BEGIN
    -- Note: You'll need to use the Supabase Dashboard to create the auth user
    -- This is because direct insertion into auth.users requires special permissions
    RAISE NOTICE 'Please create the user via Dashboard → Authentication → Users first';
END $$;
```

### Step 6: Upload Word Templates

**Option A: Via Admin Panel (After first login)**

1. Log in to the application with your admin account
2. Navigate to **★ Admin Panel** → **Templates**
3. Click **First-Time Setup** button
4. This will sync all local templates to Supabase Storage

**Option B: Manual Upload via Supabase Dashboard**

1. Navigate to **Storage** → `templates` bucket
2. Create three folders:
   - `canopy_quotation`
   - `recoair_quotation`
   - `ahu_quotation`
3. Upload templates:
   - `templates/word/Halton Quote Feb 2024.docx` → `canopy_quotation/` folder
   - `templates/word/Halton RECO Quotation Jan 2025 (2).docx` → `recoair_quotation/` folder
   - `templates/word/Halton AHU quote JAN2020.docx` → `ahu_quotation/` folder

### Step 7: Enable Email Confirmations (Optional)

1. Navigate to **Authentication** → **Settings**
2. Under **Email Auth**, toggle:
   - **Enable email confirmations**: OFF (for development) or ON (for production)
   - **Secure email change**: ON (recommended)
3. Configure SMTP settings if using email confirmations

### Step 8: Test the Setup

1. Restart your Streamlit app: `streamlit run src/app_with_auth.py`
2. Log in with your admin account
3. Test the following:
   - ✓ Login works
   - ✓ Admin Panel is accessible
   - ✓ User Management shows your admin user
   - ✓ Company Management shows all 75+ companies
   - ✓ Template Management can download/upload templates
   - ✓ Create a test quotation to verify templates work

## Verification Checklist

- [ ] Database schema created successfully
- [ ] Companies schema created and 75+ companies imported
- [ ] Storage bucket `templates` created and set to private
- [ ] Storage policies created (authenticated read, service_role full access)
- [ ] First admin user created and role set to 'admin'
- [ ] Word templates uploaded to Storage
- [ ] Application connects to new Supabase project
- [ ] Admin can log in and access Admin Panel
- [ ] All companies visible in Company Management
- [ ] Templates can be downloaded and uploaded
- [ ] Quotations generate successfully with templates

## Troubleshooting

### Issue: "No users found" in Admin Panel

**Solution**: Make sure you're using the service_role key correctly in `.env`:

```env
SUPABASE_SERVICE_ROLE_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ4dmpuY3JvbHd5dnZ5a2FsaXJ3Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1OTU1NDcyMiwiZXhwIjoyMDc1MTMwNzIyfQ.xGUwD0H5vLaDTEGEDK_PmCyldtxQ33T-Kq0zwWSH1pQ
```

### Issue: "Failed to download template"

**Solution**:
1. Verify storage bucket exists and is named `templates`
2. Check storage policies are set correctly
3. Upload templates via Admin Panel → Templates → First-Time Setup

### Issue: "Failed to create user profile"

**Solution**:
1. Verify the trigger `on_auth_user_created` exists:

```sql
SELECT tgname FROM pg_trigger WHERE tgname = 'on_auth_user_created';
```

2. If missing, re-run the schema.sql file

### Issue: Companies not showing

**Solution**:
1. Verify companies were imported:

```sql
SELECT COUNT(*) FROM public.companies;
```

2. If count is 0, re-run `database/companies_schema.sql`

### Issue: Can't upload/delete users

**Solution**: Service role key must be set correctly. Verify in `.env`:

```bash
cat .env | grep SERVICE_ROLE
```

## Important Notes

1. **Service Role Key Security**: Never commit the service_role key to version control. Keep `.env` in `.gitignore`.

2. **RLS Policies**: The database uses Row Level Security. Service role bypasses RLS for admin operations.

3. **Template Backups**: When uploading a new template, the old one is automatically backed up with timestamp.

4. **Email Verification**: For production, enable email confirmations in Supabase Auth settings.

5. **Region**: Your new project is in a region closer to your client for better performance.

## Support

If you encounter issues:
1. Check the Supabase Dashboard logs
2. Check application console for error messages
3. Verify all SQL queries executed successfully
4. Ensure all environment variables are set correctly in `.env`

## Next Steps After Setup

1. Create additional admin users via Admin Panel → Users → Invite User
2. Test creating quotations end-to-end
3. Upload any custom Word templates
4. Configure company list as needed
5. Set up backups and monitoring in Supabase Dashboard
