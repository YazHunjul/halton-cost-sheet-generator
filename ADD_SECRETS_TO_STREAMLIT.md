# How to Add Secrets to Streamlit Cloud

## Step-by-Step Guide

### 1. Go to Your App Settings

1. Visit https://share.streamlit.io/
2. Click on your deployed app
3. Click the **"⋮"** menu (three dots)
4. Select **"Settings"**

### 2. Add Secrets

1. In the settings page, find the **"Secrets"** section
2. Click in the text box
3. **Copy and paste this EXACTLY**:

```toml
[supabase]
SUPABASE_URL = "https://vxvjncrolwyvvykalirw.supabase.co"
SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ4dmpuY3JvbHd5dnZ5a2FsaXJ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTk1NTQ3MjIsImV4cCI6MjA3NTEzMDcyMn0.FOSEwPRuAj9FX2TBZ3UVhCa4OqEVDheWCUBYkUMIKa4"
SUPABASE_SERVICE_ROLE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZ4dmpuY3JvbHd5dnZ5a2FsaXJ3Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1OTU1NDcyMiwiZXhwIjoyMDc1MTMwNzIyfQ.xGUwD0H5vLaDTEGEDK_PmCyldtxQ33T-Kq0zwWSH1pQ"
```

4. Click **"Save"**

### 3. Reboot App

1. After saving secrets, click **"Reboot app"** button
2. Wait for app to restart (30-60 seconds)
3. Visit your app URL again

### 4. Test Login

- Email: `Yazanhunjul5@gmail.com`
- Password: `Admin123!@#`

## Important Notes

⚠️ **Make sure to:**
- Copy the entire block including `[supabase]`
- Don't modify the quotes or spacing
- Save before rebooting

✅ **After adding secrets:**
- The app will have access to your Supabase credentials
- Login should work
- Admin panel should be accessible

## Troubleshooting

### If you still see "SUPABASE_URL environment variable is not set":

1. Double-check secrets are saved correctly
2. Make sure `[supabase]` section header is included
3. Reboot the app again
4. Check app logs for any errors

### How to Check App Logs:

1. Go to your app in Streamlit Cloud
2. Click "⋮" menu → "Manage app"
3. Look at the logs section
4. Look for errors related to Supabase or secrets

## Quick Checklist

- [ ] Opened app settings in Streamlit Cloud
- [ ] Pasted secrets in TOML format
- [ ] Included `[supabase]` header
- [ ] Saved secrets
- [ ] Rebooted app
- [ ] Tested login
- [ ] Login successful!

## Visual Guide

```
Streamlit Cloud Dashboard
    ↓
Your App
    ↓
⋮ Menu → Settings
    ↓
Secrets Section
    ↓
Paste TOML secrets
    ↓
Save
    ↓
Reboot app
    ↓
✅ App now has access to Supabase!
```
