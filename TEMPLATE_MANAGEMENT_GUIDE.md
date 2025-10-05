# Template Management Guide

## How Templates Work

### Storage Location
All Word document templates are stored in the filesystem at:
```
templates/word/
├── Halton Quote Feb 2024.docx          (Canopy quotations)
├── Halton RECO Quotation Jan 2025 (2).docx  (RecoAir quotations)
├── Halton AHU quote JAN2020.docx       (AHU quotations)
└── backups/                             (Automatic backups)
```

### How Template Upload Works

1. **Download Current Template**
   - Go to Template Management page
   - Click "Download" button for any template
   - This downloads the current active template

2. **Edit Template**
   - Open the downloaded `.docx` file in Microsoft Word
   - Edit the template using Word's features
   - Keep all the `{{ variable_name }}` placeholders intact
   - Save your changes

3. **Upload New Template**
   - Go back to Template Management page
   - Click "Choose a Word document"
   - Select your edited file
   - Click "Upload & Replace"

4. **What Happens Behind the Scenes**
   - Old template is automatically backed up to `templates/word/backups/`
   - Backup includes timestamp: `Halton Quote Feb 2024_backup_20250105_143022.docx`
   - Your new template replaces the old file with the SAME filename
   - Next time anyone generates a Word document, it uses YOUR new template

### Why It Works

The system uses **file replacement** rather than database storage:
- Template path in code: `templates/word/Halton Quote Feb 2024.docx`
- When you upload, it replaces this exact file
- Code loads from the same path, gets your new content
- No code changes needed - just file replacement

### Restoring from Backup

If you need to restore an old version:
1. Go to "Backups" tab in Template Management
2. Download the backup you want to restore
3. Upload it as a new template (same process as above)

### Template Variables

Your template should contain variables like:
- `{{ company_name }}` - Company name
- `{{ project_number }}` - Project number
- `{{ canopies }}` - List of canopies
- Many more - see existing templates for full list

**IMPORTANT:** Do NOT remove or rename these variables - the system fills them in automatically.

## Summary

✓ Templates are saved as **real .docx files** on the server
✓ Upload **replaces** the file with same name
✓ Changes are **immediate** - next generation uses new template
✓ Old versions are **automatically backed up**
✓ You can **restore any backup** anytime
