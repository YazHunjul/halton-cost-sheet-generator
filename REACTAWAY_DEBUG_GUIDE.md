# Reactaway Debug Output Guide

## 🔍 Debug Logging Added

Comprehensive debug logging has been added to help diagnose if Reactaway sheets are being created.

## 📝 What to Look For in Console Output

When you generate an Excel file, you'll see debug messages in the console/terminal. Here's what each message means:

### 1. Template Sheet Detection
```
📋 DEBUG: Found 16 REACTAWAY template sheets: ['REACTAWAY  (9)', 'REACTAWAY  (10)', ...]
```
**Meaning**: The system found REACTAWAY sheets in the template
**Expected**: Should find 16 sheets
**Problem if**: Shows 0 sheets - template is missing REACTAWAY sheets

### 2. Area Option Check
```
🔍 DEBUG: Area 'Kitchen' - Reactaway option check:
   Area options: {'uvc': False, 'recoair': False, 'reactaway': True, ...}
   has_reactaway = True
   ✅ REACTAWAY DETECTED - Will create sheet for this area
   Available reactaway_sheets: 16
```
**Meaning**: The system is checking each area for Reactaway option
**Expected**: If you checked Reactaway box, should see `has_reactaway = True`
**Problem if**: Shows `has_reactaway = False` when you checked the box

### 3. Sheet Creation (If Reactaway is Enabled)
```
🟢 DEBUG: Creating REACTAWAY sheet for area 'Kitchen'
   Using template sheet: REACTAWAY  (9)
   Renaming to: REACTAWAY - Ground Floor (1)
   Sheet state set to: visible
   Tab color set
   Metadata written
   Title written to B1: Ground Floor - Kitchen - REACTAWAY SYSTEM
   ✅ REACTAWAY sheet 'REACTAWAY - Ground Floor (1)' created successfully
```
**Meaning**: REACTAWAY sheet is being created and configured
**Expected**: See this message for each area with Reactaway enabled
**Problem if**: Don't see this message when Reactaway is checked

### 4. Sheet Creation Skipped (If Reactaway NOT Enabled)
```
⚪ DEBUG: Reactaway NOT enabled for area 'Kitchen' - skipping
```
**Meaning**: This area doesn't have Reactaway option checked
**Expected**: Normal if you didn't check Reactaway for this area
**Problem if**: See this when you DID check Reactaway

### 5. Unused Sheet Cleanup
```
🗑️  Removing 222 unused system template sheets...
   DEBUG: Unused REACTAWAY sheets to delete: ['REACTAWAY  (10)', 'REACTAWAY  (11)', ...]
   Deleted unused REACTAWAY: REACTAWAY  (10)
   Deleted unused REACTAWAY: REACTAWAY  (11)
```
**Meaning**: Removing unused template REACTAWAY sheets
**Expected**: Should delete all REACTAWAY sheets EXCEPT the ones you created
**Problem if**: Deletes your created sheets (would show "REACTAWAY - Ground Floor")

### 6. Visibility Check
```
🔒 DEBUG: Checking sheet visibility. REACTAWAY sheets should start with 'REACTAWAY -'
   ✅ Keeping REACTAWAY sheet visible: REACTAWAY - Ground Floor (1)
```
**Meaning**: System is ensuring your REACTAWAY sheet stays visible
**Expected**: See one line for each area with Reactaway
**Problem if**: Shows "Hiding" instead of "Keeping"

### 7. Final Status
```
📊 DEBUG: Final REACTAWAY sheet status: 1 visible REACTAWAY sheets
```
**Meaning**: Summary of how many REACTAWAY sheets are in the final Excel
**Expected**: Number should match how many areas have Reactaway enabled
**Problem if**: Shows 0 when you enabled Reactaway

## 🐛 Debugging Scenarios

### Scenario 1: Checkbox Checked but No Sheet Created

**Expected Output**:
```
📋 DEBUG: Found 16 REACTAWAY template sheets
🔍 DEBUG: Area 'Kitchen' - Reactaway option check:
   has_reactaway = True
   ✅ REACTAWAY DETECTED
🟢 DEBUG: Creating REACTAWAY sheet
   ✅ REACTAWAY sheet created successfully
📊 DEBUG: Final REACTAWAY sheet status: 1 visible REACTAWAY sheets
```

**If you see**:
```
🔍 DEBUG: Area 'Kitchen' - Reactaway option check:
   has_reactaway = False
⚪ DEBUG: Reactaway NOT enabled for area 'Kitchen' - skipping
```

**Problem**: The checkbox value isn't being saved to session state
**Solution**:
1. Check the box
2. Add a canopy or click another field to trigger state update
3. Verify "Detailed Structure" shows "Reactaway: Yes"
4. Then generate Excel

### Scenario 2: Sheet Created but Hidden

**Expected Output**:
```
🟢 DEBUG: Creating REACTAWAY sheet
   Sheet state set to: visible
   ✅ REACTAWAY sheet created successfully
🔒 DEBUG: Checking sheet visibility
   ✅ Keeping REACTAWAY sheet visible: REACTAWAY - Ground Floor (1)
📊 DEBUG: Final REACTAWAY sheet status: 1 visible REACTAWAY sheets
```

**If you see**:
```
🟢 DEBUG: Creating REACTAWAY sheet
   ✅ REACTAWAY sheet created successfully
🔒 DEBUG: Checking sheet visibility
   🔒 Hiding unused REACTAWAY template: REACTAWAY - Ground Floor (1)
📊 DEBUG: Final REACTAWAY sheet status: 0 visible REACTAWAY sheets
```

**Problem**: Sheet name doesn't match visibility whitelist
**Solution**: Check code - sheet name should start with "REACTAWAY -"

### Scenario 3: No REACTAWAY Template Sheets

**If you see**:
```
📋 DEBUG: Found 0 REACTAWAY template sheets: []
```

**Problem**: Template file doesn't have REACTAWAY sheets
**Solution**:
- Verify using correct template: "COST SHEET R19.2 SEPT2025ss.xlsx"
- Check template actually has REACTAWAY sheets
- Re-download template if needed

### Scenario 4: Template Sheets Not Unhidden

**Expected at start**:
```
Found template sheet: 'REACTAWAY  (9)' - Current state: hidden
✅ Unhidden template sheet: REACTAWAY  (9)
```

**If sheets stay hidden**: Code has issue with unhiding logic

## 📊 How to View Console Output

### In Streamlit Cloud:
1. Go to app dashboard
2. Click "Manage app" → "Logs"
3. Look for debug messages when generating Excel

### Running Locally:
1. Start app in terminal: `streamlit run app.py`
2. Generate Excel in browser
3. Watch terminal for debug output
4. Copy relevant lines to share for debugging

### In Test Script:
```bash
python test_reactaway.py 2>&1 | grep DEBUG
```

## ✅ Success Pattern

When everything works correctly, you'll see this sequence:

```
1. 📋 DEBUG: Found 16 REACTAWAY template sheets
2. 🔍 DEBUG: Area 'Your Area' - Reactaway option check:
      has_reactaway = True
      ✅ REACTAWAY DETECTED
3. 🟢 DEBUG: Creating REACTAWAY sheet
      Using template sheet: REACTAWAY  (9)
      Renaming to: REACTAWAY - Ground Floor (1)
      Sheet state set to: visible
      ✅ REACTAWAY sheet created successfully
4. 🔒 DEBUG: Checking sheet visibility
      ✅ Keeping REACTAWAY sheet visible: REACTAWAY - Ground Floor (1)
5. 📊 DEBUG: Final REACTAWAY sheet status: 1 visible REACTAWAY sheets
```

## 🎯 Quick Checklist

When debugging, verify each step:

- [ ] Step 1: Template has REACTAWAY sheets (📋 shows count > 0)
- [ ] Step 2: Area option detected (🔍 shows has_reactaway = True)
- [ ] Step 3: Sheet created (🟢 shows creation messages)
- [ ] Step 4: Sheet stays visible (✅ Keeping, not 🔒 Hiding)
- [ ] Step 5: Final count > 0 (📊 shows visible sheets)

## 📞 Reporting Issues

If the feature isn't working, include:

1. **Console output** showing the DEBUG messages
2. **Screenshot** of checkbox being checked
3. **Screenshot** of "Detailed Structure" showing options
4. **Excel file** (if generated) so we can inspect it

The debug output will show exactly where the process is breaking.
