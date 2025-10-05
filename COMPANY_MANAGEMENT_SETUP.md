# 🏢 Company Management Setup Guide

## Overview

You can now manage companies and their addresses from the admin dashboard! This replaces the hardcoded company list with a database-driven system.

---

## 📋 Setup Steps

### **Step 1: Run Company Database Schema**

1. **Open Supabase Dashboard**
   - Go to: https://supabase.com/dashboard/project/rlvtnyotgsgywxkshazo

2. **Open SQL Editor**
   - Click **SQL Editor** in sidebar
   - Click **New Query**

3. **Run Schema**
   - Copy **ALL** content from: `database/companies_schema.sql`
   - Paste into SQL Editor
   - Click **RUN**
   - ✅ You should see "Success" and "Companies table created and populated successfully!"

This will:
- Create the `companies` table
- Import all your existing companies (73 companies)
- Set up proper security policies

---

### **Step 2: Restart Your App**

```bash
streamlit run app.py
```

The app will now load companies from the database!

---

## 🎯 How to Use

### **Access Company Management** (Admin Only)

1. Login as admin
2. Sidebar → Select **"🏢 Company Management"**

### **View All Companies**

Tab: **"📋 All Companies"**
- View all companies in alphabetical order
- See company name and full address
- Filter by status (Active/Inactive)
- Search by name or address

### **Add New Company**

Tab: **"➕ Add Company"**

1. Fill in:
   - **Company Name**: Full company name
   - **Address**: Multi-line address (use Enter for new lines)

2. Click **"💾 Add Company"**

3. ✅ Company is immediately available in dropdowns!

### **Edit Existing Company**

1. Go to **"All Companies"** tab
2. Click on any company to expand
3. Edit:
   - Company name
   - Address
4. Click **"💾 Update"**

### **Deactivate/Activate Company**

- **Deactivate**: Hides from dropdowns but keeps in database
- **Activate**: Makes visible in dropdowns again
- Click **"🚫 Deactivate"** or **"✅ Activate"**

### **Delete Company**

⚠️ Permanently removes from database
- Click **"🗑️ Delete"**
- Cannot be undone!

---

## 🔄 How It Works

### **Dynamic Loading**

Your app now:
1. Loads companies from database first
2. Falls back to hardcoded list if database unavailable
3. Only shows **active** companies in dropdowns
4. Updates immediately when you add/edit companies

### **Database Structure**

```sql
companies table:
├── id (UUID)
├── name (unique)
├── address
├── is_active (true/false)
├── created_at
├── updated_at
├── created_by (user who created)
└── updated_by (user who last updated)
```

---

## ✅ Benefits

### **Before** (Hardcoded)
- ❌ Had to edit code to add companies
- ❌ Required redeployment
- ❌ No change tracking
- ❌ All users saw all companies

### **After** (Database)
- ✅ Add companies via UI
- ✅ Instant updates
- ✅ Track who created/updated
- ✅ Show/hide companies easily
- ✅ Admin-only management

---

## 🧪 Test It

1. **Login as admin**
2. **Navigate**: Sidebar → "🏢 Company Management"
3. **Add test company**:
   - Name: "Test Company Ltd"
   - Address: "123 Test Street\nTest City\nTE1 2ST"
   - Click "Add Company"
4. **Verify in project setup**:
   - Go to "Single Page Setup"
   - Company dropdown should include "Test Company Ltd"
5. **Edit the company**:
   - Go back to Company Management
   - Change address
   - Verify changes appear in dropdown
6. **Deactivate**:
   - Deactivate the test company
   - Check dropdown - should not appear
7. **Delete**:
   - Delete the test company
   - Removed from database

---

## 🔒 Security

- ✅ Only **admins** can add/edit/delete companies
- ✅ **All users** can view active companies (for dropdowns)
- ✅ Row Level Security enforced
- ✅ Audit trail (created_by, updated_by, timestamps)

---

## 📊 Current Status

After running the schema:
- **73 companies** imported from your existing list
- All marked as **active**
- Ready to use immediately
- Fully backward compatible

---

## 🛠️ Troubleshooting

### **Companies not showing in dropdown**

**Check**:
1. Companies table exists in Supabase
2. Company is marked as `is_active = true`
3. App has been restarted
4. No database connection errors in logs

**Fix**:
```python
# Test database connection
python -c "
import sys
sys.path.insert(0, 'src')
from config.business_data import get_companies_from_database
companies = get_companies_from_database()
print(f'Loaded {len(companies)} companies')
print(list(companies.keys())[:5])
"
```

### **Can't add company - already exists**

- Company names must be unique
- Check if company already exists in database
- Try a slightly different name

### **Changes not appearing**

- App loads companies once at startup
- Restart app to reload from database
- Or use Streamlit's "Rerun" button

---

## 💡 Next Steps

You can now:
1. ✅ Add new companies as you get new clients
2. ✅ Update addresses when companies move
3. ✅ Deactivate old companies without deleting them
4. ✅ Keep clean, up-to-date company list
5. ✅ No code changes needed!

---

## 📞 Quick Reference

**Add Company**: Admin → Company Management → Add Company tab
**Edit Company**: Admin → Company Management → All Companies → Click company → Edit
**Deactivate**: Click company → 🚫 Deactivate button
**Delete**: Click company → 🗑️ Delete button (⚠️ permanent!)

---

**Ready to use!** Run the schema and start managing your companies from the admin dashboard! 🚀
