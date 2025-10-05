# 🔧 Fix Infinite Recursion Error in RLS Policies

## ❌ Problem
You're getting: `infinite recursion detected in policy for relation "user_profiles"`

This happens because the RLS policies were checking `user_profiles.role` while querying `user_profiles`, creating a circular reference.

---

## ✅ Solution: Run the Fixed Schema

### **Step 1: Run Fixed Schema in Supabase**

1. **Open Supabase Dashboard**
   - Go to: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo

2. **Open SQL Editor**
   - Click **SQL Editor** in left sidebar
   - Click **New Query**

3. **Run the Fixed Schema**
   - Copy **ALL** content from: `database/schema_fixed.sql`
   - Paste into SQL Editor
   - Click **RUN** (or Ctrl+Enter)
   - Wait for "Success. No rows returned" message

This will:
- Drop all the old problematic policies
- Create new simplified policies without recursion
- Keep all your existing data intact

---

### **Step 2: Recreate Admin User (if needed)**

If the user was created but couldn't login:

```bash
source venv/bin/activate
python create_my_admin.py
```

Or create manually in Supabase Dashboard:
- **Authentication** → **Users** → **Add User**
- Email: `yazan@halton.com`
- Password: `Halton2025!`
- ✅ Auto Confirm User
- User Metadata:
  ```json
  {
    "first_name": "Yazan",
    "last_name": "Admin",
    "role": "admin"
  }
  ```

---

### **Step 3: Test Login**

```bash
streamlit run app.py
```

Login with:
- Email: `yazan@halton.com`
- Password: `Halton2025!`

✅ Should work now!

---

## 📝 What Changed?

### **Old Schema (Problematic):**
```sql
-- This caused infinite recursion:
CREATE POLICY "Admins can view all profiles"
    ON public.user_profiles
    FOR SELECT
    USING (
        EXISTS (
            SELECT 1 FROM public.user_profiles  -- ← Querying same table!
            WHERE id = auth.uid() AND role = 'admin'
        )
    );
```

### **New Schema (Fixed):**
```sql
-- Simplified policies without recursion:
CREATE POLICY "Users can view own profile"
    ON public.user_profiles
    FOR SELECT
    USING (auth.uid() = id);  -- ← Simple, no recursion

-- Service role has full access (for admin operations via backend)
CREATE POLICY "Service role full access"
    ON public.user_profiles
    FOR ALL
    USING (current_user = 'service_role');
```

---

## 🔐 How Admin Access Works Now

Instead of checking admin role at the database level (which caused recursion), we:

1. **Database Level**: All authenticated users can view their own profile
2. **Application Level**: Admin checks happen in the application code
3. **Service Role**: Backend operations use service role key for admin tasks

This is actually **more flexible** and **easier to manage**!

---

## ✅ Verify Fix Applied

After running the fixed schema, verify with:

```sql
-- Check policies exist
SELECT schemaname, tablename, policyname
FROM pg_policies
WHERE schemaname = 'public'
ORDER BY tablename, policyname;
```

You should see:
- `user_profiles` → `Users can view own profile`
- `user_profiles` → `Users can update own profile`
- `user_profiles` → `Service role full access`
- `user_profiles` → `Allow insert for new users`

---

## 🆘 Still Having Issues?

If you still get errors:

1. **Check Supabase Logs**
   - Dashboard → **Logs & Analytics** → **Database Logs**

2. **Verify Tables Exist**
   - **Table Editor** → Check `user_profiles`, `user_invitations`, `audit_logs`

3. **Test Connection**
   ```bash
   python test_supabase_connection.py
   ```

4. **Check User Exists**
   ```sql
   SELECT id, email, email_confirmed_at, raw_user_meta_data
   FROM auth.users
   WHERE email = 'yazan@halton.com';
   ```

---

## 📞 Summary

**Quick Fix:**
1. Run `database/schema_fixed.sql` in Supabase SQL Editor
2. Create/verify admin user exists
3. Login and test

**Time**: 3 minutes

**Result**: No more recursion error! ✅
