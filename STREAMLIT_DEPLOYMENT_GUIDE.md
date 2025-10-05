# Streamlit Cloud Deployment Guide

Complete guide for deploying the Halton Quotation System to Streamlit Cloud.

## Prerequisites

- ✅ Supabase project set up (https://vxvjncrolwyvvykalirw.supabase.co)
- ✅ Database schema created
- ✅ Admin user created (Yazanhunjul5@gmail.com)
- ✅ GitHub repository with your code

## Step 1: Prepare Your Repository

### 1.1 Ensure `.gitignore` is Correct

Make sure these files are in `.gitignore`:

```
.env
.streamlit/secrets.toml
venv/
__pycache__/
*.pyc
```

### 1.2 Verify `requirements.txt`

Your `requirements.txt` should include all dependencies:

```
streamlit
supabase
python-dotenv
pandas
openpyxl
python-docx
docxtpl
Pillow
```

### 1.3 Push to GitHub

```bash
git add .
git commit -m "Prepare for Streamlit Cloud deployment"
git push origin main
```

## Step 2: Deploy to Streamlit Cloud

### 2.1 Go to Streamlit Cloud

1. Visit https://share.streamlit.io/
2. Sign in with your GitHub account
3. Click **"New app"**

### 2.2 Configure App Settings

- **Repository**: Select your GitHub repository
- **Branch**: `main` (or your default branch)
- **Main file path**: `src/app_with_auth.py`
- **App URL**: Choose your custom URL (e.g., `halton-quotation-system`)

### 2.3 Add Secrets (CRITICAL STEP)

1. Click **"Advanced settings"**
2. In the **"Secrets"** section, paste this:

```toml
[supabase]
SUPABASE_URL = "https://vxvjncrolwyvvykalirw.supabase.co"
SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ4dmpuY3JvbHd5dnZ5a2FsaXJ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTk1NTQ3MjIsImV4cCI6MjA3NTEzMDcyMn0.FOSEwPRuAj9FX2TBZ3UVhCa4OqEVDheWCUBYkUMIKa4"
SUPABASE_SERVICE_ROLE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ4dmpuY3JvbHd5dnZ5a2FsaXJ3Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1OTU1NDcyMiwiZXhwIjoyMDc1MTMwNzIyfQ.xGUwD0H5vLaDTEGEDK_PmCyldtxQ33T-Kq0zwWSH1pQ"
```

3. Click **"Deploy"**

## Step 3: Verify Deployment

### 3.1 Check App Status

- Streamlit Cloud will show deployment logs
- Wait for "Your app is live!" message
- App URL will be: `https://your-app-name.streamlit.app`

### 3.2 Test the Application

1. Visit your app URL
2. Log in with: **Yazanhunjul5@gmail.com** / **Admin123!@#**
3. Verify:
   - ✓ Login works
   - ✓ Admin Panel accessible
   - ✓ Companies load
   - ✓ Can create quotations

## How It Works

### Local Development
```
.env file → supabase_config.py → Supabase
```

### Streamlit Cloud Deployment
```
Streamlit Secrets → supabase_config.py → Supabase
```

The `supabase_config.py` file automatically detects which environment it's in:

```python
# 1. Check if running on Streamlit Cloud
if hasattr(st, 'secrets') and 'supabase' in st.secrets:
    # Use Streamlit secrets
    SUPABASE_URL = st.secrets["supabase"]["SUPABASE_URL"]
else:
    # Use .env file (local)
    SUPABASE_URL = os.getenv("SUPABASE_URL")
```

## Template Storage on Streamlit Cloud

### Important: Streamlit Cloud has Ephemeral Filesystem

Files uploaded to the filesystem **will be lost** when the app restarts. That's why we use **Supabase Storage** for templates.

### How Templates Work on Streamlit Cloud:

1. **Templates stored in Supabase Storage bucket** (`templates`)
2. When generating a quotation:
   - App downloads template from Supabase to temporary folder
   - Generates Word document
   - Provides download link to user
3. When app restarts, temporary files are cleared
4. Templates remain safe in Supabase Storage

### First-Time Setup After Deployment:

1. Log in as admin
2. Go to **Admin Panel** → **Templates**
3. Click **"First-Time Setup"** to sync templates to Supabase
4. This uploads all three templates (Canopy, RecoAir, AHU) to Supabase Storage

## Troubleshooting

### Issue: "SUPABASE_URL environment variable is not set"

**Solution**: Make sure you added secrets in Streamlit Cloud:
1. Go to app settings (⋮ menu)
2. Click "Secrets"
3. Paste the TOML configuration above
4. Save and reboot app

### Issue: "Failed to connect to Supabase"

**Solution**: Verify your secrets are correct:
- URL should be: `https://vxvjncrolwyvvykalirw.supabase.co`
- Keys should match your Supabase project

### Issue: "No templates found"

**Solution**:
1. Go to Admin Panel → Templates
2. Click "First-Time Setup"
3. This uploads templates to Supabase Storage

### Issue: "Login error: infinite recursion"

**Solution**: Make sure you ran `database/rls_fix_final.sql` in Supabase SQL Editor

## Security Best Practices

### ✅ DO:
- ✓ Keep secrets in Streamlit Cloud secrets manager
- ✓ Use `.gitignore` to exclude `.env` and `secrets.toml`
- ✓ Use RLS (Row Level Security) in Supabase
- ✓ Use service_role key only for admin operations
- ✓ Change default admin password after first login

### ❌ DON'T:
- ✗ Commit `.env` to GitHub
- ✗ Share service_role key publicly
- ✗ Disable RLS in production
- ✗ Store sensitive data in filesystem

## Updating the Deployed App

### Method 1: Automatic (Recommended)
1. Push changes to GitHub: `git push origin main`
2. Streamlit Cloud auto-detects changes
3. App automatically redeploys

### Method 2: Manual
1. Go to Streamlit Cloud dashboard
2. Click on your app
3. Click "Reboot app"

## Monitoring and Logs

### View Logs
1. Go to Streamlit Cloud dashboard
2. Click on your app
3. Click "Manage app" → "Logs"

### Common Log Messages
- "Successfully connected to Supabase" → ✓ Working
- "Environment variable not set" → ❌ Check secrets
- "Infinite recursion" → ❌ Run RLS fix

## App Settings

### Memory and Resources
- Streamlit Cloud free tier: 1GB RAM
- If app runs out of memory, consider:
  - Optimizing data loading
  - Using pagination for large datasets
  - Caching with `@st.cache_data`

### Custom Domain (Optional)
1. Go to app settings
2. Click "Custom domain"
3. Add your domain and follow DNS instructions

## Support

### Streamlit Cloud Issues
- Docs: https://docs.streamlit.io/streamlit-community-cloud
- Forum: https://discuss.streamlit.io/

### Supabase Issues
- Docs: https://supabase.com/docs
- Support: https://supabase.com/support

## Quick Reference

| Item | Value |
|------|-------|
| **App File** | `src/app_with_auth.py` |
| **Supabase URL** | https://vxvjncrolwyvvykalirw.supabase.co |
| **Admin Email** | Yazanhunjul5@gmail.com |
| **Admin Password** | Admin123!@# (change after login) |
| **Template Bucket** | `templates` (Supabase Storage) |
| **Database Region** | Closer to client |

## Next Steps After Deployment

1. ✓ Change admin password
2. ✓ Upload templates via First-Time Setup
3. ✓ Test creating a quotation end-to-end
4. ✓ Invite additional users via Admin Panel
5. ✓ Share app URL with team
6. ✓ Monitor usage and logs
