# Reactaway Troubleshooting Guide

## ✅ Feature Status
The Reactaway feature has been fully implemented and tested. The code is working correctly in automated tests.

## 🔍 How to Verify It's Working in the App

### Step 1: Check the Checkbox is Present

1. Start the app: `streamlit run app.py`
2. Navigate to "Single Page Setup" or create a new project
3. Add a level and an area
4. Look for the **Reactaway** checkbox in the Area Options section
5. It should appear as the 8th option after: UV-C, RecoAir, Marvel, Vent CLG, Pollustop, Aerolys, XEU

**Expected**: You should see a checkbox labeled "Reactaway"

### Step 2: Check the Option Before Generating

Before clicking "Generate Excel", expand the "Detailed Structure" section and verify:
- Your area is listed
- The options show "Reactaway: Yes" if checked

Example:
```
Ground Floor
  • Kitchen (1 canopies)
    Options: Reactaway
```

### Step 3: Generate and Verify

1. Check the "Reactaway" checkbox
2. Click "Generate Excel"
3. Download the Excel file
4. Open in Excel
5. Scroll through the sheet tabs (there are 200+ sheets)
6. Look for: `REACTAWAY - [Your Level Name] (1)`

**Example**: `REACTAWAY - Ground Floor (1)`

## 🐛 Common Issues

### Issue 1: Checkbox Not Visible
**Solution**: The checkbox was recently added. Try:
1. Hard refresh the browser (Ctrl+Shift+R or Cmd+Shift+R)
2. Restart the Streamlit app
3. Clear browser cache

### Issue 2: Option Not Saving
**Problem**: The checkbox value isn't being saved to session state

**Debug Steps**:
1. Check the checkbox
2. Add a canopy to the area (this triggers a state update)
3. Look at the "Detailed Structure" - does it show "Reactaway: Yes"?
4. If not, there's a state management issue

**Solution**:
- Make sure you're checking the box BEFORE adding canopies
- Or, after checking the box, click into another field to trigger the update callback

### Issue 3: Sheet Not in Excel File
**Problem**: Generated Excel doesn't have REACTAWAY sheet

**Verification Steps**:
1. Run the test script: `python test_reactaway.py`
2. If test passes but app doesn't work, the issue is in how the app saves state

**Expected Test Output**:
```
✅ SUCCESS: Found 1 REACTAWAY sheet(s)
  - REACTAWAY - Ground Floor (1)
```

## 📝 Step-by-Step Test Procedure

### Complete Test From Scratch:

1. **Start Fresh**:
   ```bash
   cd /Users/yazan/Desktop/Efficiency/UKCS
   source venv/bin/activate
   streamlit run app.py
   ```

2. **Create Project**:
   - Go to "Single Page Setup"
   - Fill in project details (name, number, etc.)

3. **Add Level**:
   - Level Number: 1
   - Level Name: "Ground Floor"

4. **Add Area with Reactaway**:
   - Area Name: "Kitchen"
   - **CHECK the "Reactaway" checkbox** ← IMPORTANT
   - Leave other options unchecked

5. **Add a Canopy** (at least one):
   - Reference: C001
   - Model: KVH-6
   - Configuration: Island
   - Length/Width/Height: Default values

6. **Verify Before Generating**:
   - Expand "Project Structure" in sidebar
   - Should show: "Kitchen" with options including "Reactaway"

7. **Generate**:
   - Click "Generate Excel"
   - Download the file

8. **Verify in Excel**:
   - Open the downloaded file
   - Right-click on sheet tab navigation arrows → "More Sheets..."
   - Search for "REACTAWAY"
   - Should find: "REACTAWAY - Ground Floor (1)"

## 🔬 Advanced Debugging

### Check Session State

Add this temporary code to see what's in session state:

In `src/app.py`, add before the generate Excel button:
```python
# Temporary debug code
if st.checkbox("Show Debug Info"):
    st.write("Session State Levels:")
    st.json(st.session_state.levels)
```

Then check if `reactaway` appears in the area options.

### Check Generated Data

The generate_excel_section function (line 2041) creates:
```python
final_project_data = st.session_state.project_info.copy()
final_project_data['levels'] = st.session_state.levels
```

The `st.session_state.levels` should contain your area with `options.reactaway = True`.

## ✅ Verification Checklist

- [ ] Reactaway checkbox is visible in the UI
- [ ] Checking the box updates the UI
- [ ] "Detailed Structure" shows "Reactaway: Yes"
- [ ] Test script passes: `python test_reactaway.py`
- [ ] Generated Excel contains REACTAWAY sheet
- [ ] REACTAWAY sheet is visible (not hidden)
- [ ] Sheet has correct title in B1

## 📞 Still Not Working?

If you've followed all steps and it's still not working:

1. **Verify your code is updated**:
   - Check `src/utils/excel.py` line 242: Should include `'REACTAWAY' in sheet_name`
   - Check `src/utils/excel.py` line 3901: Should include `reactaway_sheets`
   - Check `src/utils/excel.py` line 3916: Should include `'REACTAWAY -'`

2. **Run automated test**:
   ```bash
   python test_reactaway.py
   ```
   If this passes, the backend is working.

3. **Check the actual checkbox value**:
   After checking the box, look at the session state (use debug code above)
   The area should have: `{"options": {"reactaway": True, ...}}`

4. **Try the debug script**:
   ```bash
   python debug_reactaway_state.py
   ```
   This simulates what the app should do.

## 📊 Expected Behavior Summary

| Action | Expected Result |
|--------|----------------|
| Check Reactaway box | Box shows as checked |
| Add canopy | Reactaway option saves to state |
| View "Detailed Structure" | Shows "Reactaway" in options list |
| Generate Excel | Creates REACTAWAY sheet |
| Open Excel file | REACTAWAY sheet is visible |
| Click on REACTAWAY sheet | Shows project data with "REACTAWAY SYSTEM" title |

## 🎯 Success Criteria

The feature is working correctly when:
1. ✅ Checkbox appears in UI
2. ✅ Checking box updates state
3. ✅ Option appears in project summary
4. ✅ REACTAWAY sheet is created in Excel
5. ✅ Sheet is visible (not hidden)
6. ✅ Sheet contains correct project metadata
