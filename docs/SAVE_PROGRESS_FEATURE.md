# Save Progress Feature

The Halton Cost Sheet Generator now includes a save progress feature that allows users to save their form data and restore it later using a shareable link.

## How It Works

1. **Fill in the Form**: Enter your project information, structure, and canopy details in the Single Page Setup.

2. **Save Your Progress**: 
   - Scroll down to find the "Save Your Progress" section
   - Click the "Generate Save Link" button
   - A unique link will be generated containing all your form data

3. **Copy or Save the Link**:
   - Use the "Copy to Clipboard" button to copy the link
   - Or click "Open in New Tab" to test the link
   - Save the link somewhere safe (bookmark, email, notes, etc.)

4. **Restore Your Progress**:
   - Simply open the saved link in your browser
   - All form fields will be automatically populated with your saved data
   - You'll see a success message confirming the data was loaded

## Technical Details

### Data Storage
- Form data is compressed and encoded into the URL itself
- No data is stored on the server - everything is in the link
- The link contains all form fields, project structure, and canopy configurations

### Compression
- Uses zlib compression to minimize URL length
- Base64 encoding ensures URL safety
- Typical compression ratio: 60-80% reduction in size

### Security
- Data is only accessible to those who have the link
- No sensitive data is transmitted to any server
- All processing happens in the browser

### Supported Fields
The following data is saved and restored:
- Project information (name, number, customer, etc.)
- Company and address details
- Sales contact and estimator
- Complete project structure (levels, areas, canopies)
- All canopy configurations and options
- Template selection

### Browser Compatibility
- Works with all modern browsers
- Clipboard copy feature requires HTTPS or localhost
- Maximum URL length varies by browser (typically 2000-65000 characters)

## Limitations

1. **URL Length**: Very large projects with many canopies may exceed browser URL limits
2. **Browser Bookmarks**: Some browsers may truncate very long bookmarked URLs
3. **Data Changes**: If the form structure changes significantly in future updates, old links may not work perfectly

## Best Practices

1. **Save Regularly**: Generate new save links as you make progress
2. **Multiple Backups**: Save links in multiple places for important projects
3. **Test Links**: Click "Open in New Tab" to verify the link works before relying on it
4. **Share Carefully**: Anyone with the link can access and modify the form data

## Implementation Notes

The feature is implemented in:
- `/src/utils/state_manager.py` - Core save/load functionality
- `/src/app.py` - Integration with the Single Page Setup

Key functions:
- `compress_state()` - Compresses form data for URL storage
- `decompress_state()` - Restores form data from URL
- `extract_form_state()` - Gathers all relevant form fields
- `restore_form_state()` - Populates form fields from saved data
- `generate_save_link()` - Creates the shareable URL
- `load_from_url()` - Loads data when opening a saved link