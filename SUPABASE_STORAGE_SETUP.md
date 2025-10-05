# Supabase Storage Setup for Templates

## Overview

Templates are now stored in **Supabase Storage** (a cloud storage bucket) instead of the local filesystem. This ensures templates persist when hosting on platforms like Streamlit Cloud, Heroku, etc.

## Setup Steps

### 1. Create Storage Bucket in Supabase

1. Go to your Supabase dashboard: https://supabase.com/dashboard
2. Select your project
3. Click on **Storage** in the left sidebar
4. Click **Create a new bucket**
5. Name it: `templates`
6. Set it as **Private** (not public)
7. Click **Create bucket**

### 2. Upload Initial Templates

Once the bucket is created, you have two options:

#### Option A: Use the Admin Interface (Recommended)

1. Log into your app as an admin
2. Go to **Template Management**
3. Expand the **First-Time Setup** section
4. Click **Sync Local Templates to Storage**
5. This will upload all existing templates from `templates/word/` to Supabase Storage

#### Option B: Manual Upload via Supabase Dashboard

1. In Supabase Storage, click on the `templates` bucket
2. Create folders:
   - `canopy_quotation/`
   - `recoair_quotation/`
   - `ahu_quotation/`
3. Upload templates to their respective folders:
   - Upload `Halton Quote Feb 2024.docx` to `canopy_quotation/`
   - Upload `Halton RECO Quotation Jan 2025 (2).docx` to `recoair_quotation/`
   - Upload `Halton AHU quote JAN2020.docx` to `ahu_quotation/`

### 3. Set Storage Permissions (RLS)

The bucket should have these policies:

```sql
-- Allow authenticated users to read templates
CREATE POLICY "Authenticated users can read templates"
ON storage.objects FOR SELECT
TO authenticated
USING (bucket_id = 'templates');

-- Allow service role full access (for admin operations)
-- This is automatically set when using service_role client
```

## How It Works

### Template Upload
1. Admin uploads new template via Template Management page
2. Old template is automatically backed up to `backups/{template_key}/`
3. New template is saved to `{template_key}/{filename}`
4. Template is immediately available for document generation

### Document Generation
1. When generating a Word document, the system checks if template exists locally
2. If not found locally, it downloads from Supabase Storage
3. Template is cached locally (in ephemeral filesystem)
4. Document is generated using the template
5. Next generation re-uses cached template (or downloads fresh if cache expired)

### Template Structure in Supabase

```
templates/  (bucket)
├── canopy_quotation/
│   └── Halton Quote Feb 2024.docx
├── recoair_quotation/
│   └── Halton RECO Quotation Jan 2025 (2).docx
├── ahu_quotation/
│   └── Halton AHU quote JAN2020.docx
└── backups/
    ├── canopy_quotation/
    │   ├── Halton Quote Feb 2024_backup_20250105_143022.docx
    │   └── Halton Quote Feb 2024_backup_20250104_091530.docx
    ├── recoair_quotation/
    │   └── ...
    └── ahu_quotation/
        └── ...
```

## Benefits

✓ **Persistent Storage**: Templates survive app restarts and redeployments
✓ **Automatic Backups**: Old versions saved automatically when uploading new templates
✓ **Version Control**: Access to all historical template versions
✓ **Cloud-Based**: Works with any hosting platform (Streamlit Cloud, Heroku, AWS, etc.)
✓ **Admin Control**: Admins can update templates without code changes or filesystem access
✓ **Scalable**: Supabase Storage handles file serving efficiently

## Troubleshooting

### Templates Not Found

If you get "Template not found in storage" errors:

1. Check that the `templates` bucket exists in Supabase
2. Verify templates are uploaded to the correct folders
3. Run the "Sync Local Templates to Storage" function from Template Management

### Permission Errors

If you get permission errors:

1. Ensure you're using the service_role key (stored in `.env`)
2. Check that RLS policies allow authenticated users to read from `templates` bucket

### Upload Failures

If template uploads fail:

1. Check file size (should be under 50MB for Supabase free tier)
2. Ensure file is a valid .docx format
3. Check Supabase project storage quota
