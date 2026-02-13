# Testing Guide

## Prerequisites
- Microsoft Excel (Windows/Mac) or LibreOffice Calc with macro support
- The updated `HeatMap.bas` file from this repository
- The original Excel file with test data

## Step-by-Step Testing Instructions

### 1. Backup Your Original File
```
1. Make a copy of "AVLDrive_Heatmap_Tool version_4 (2).xlsm"
2. Rename the copy to include "_backup" or similar
3. Work with a copy, not the original
```

### 2. Apply the Fix

#### Method A: Direct Code Replacement (Faster)
1. Open the Excel file
2. Press `Alt + F11` (or `Option + F11` on Mac) to open VBA Editor
3. In the Project Explorer (left panel), find `HeatMap` module
4. Double-click to open it
5. Select all code: `Ctrl + A` (or `Cmd + A` on Mac)
6. Delete selected code
7. Open `HeatMap.bas` from this repository in a text editor
8. Copy all code: `Ctrl + A`, then `Ctrl + C`
9. Paste into Excel VBA Editor: `Ctrl + V`
10. Save: `Ctrl + S`
11. Close VBA Editor

#### Method B: Import Module (More Thorough)
1. Open the Excel file
2. Press `Alt + F11` to open VBA Editor
3. Right-click on `HeatMap` in Project Explorer
4. Select "Remove HeatMap"
5. When prompted to export, select "No"
6. Go to File → Import File
7. Navigate to and select `HeatMap.bas` from this repository
8. Save and close VBA Editor

### 3. Prepare Test Data

Create test data in "Data Transfer Sheet":

```
Column A (Operation Mode) | Column B (DR1) | Column C (DR2) | Column D (DR3) | Column E (DR4)
---------------------------|----------------|----------------|----------------|----------------
Operation Modes            | Vehicle 1      | Vehicle 2      | Vehicle 3      | Vehicle 4
Mode 1                     | 100            | 200            | 300            | 400
Mode 2                     | 150            | 250            | 350            | 450
Mode 3                     | 110            | 210            | 310            | 410
```

Key points:
- Make sure "Operation Modes" text is in column A (anchor text)
- Vehicle column headers should be named DR1, DR2, DR3, DR4, etc. (if following the DR convention) or any other vehicle names
- The row below the header row (on HeatMap Sheet) can contain "DR" markers to identify vehicle columns
- Include multiple vehicles (at least 3-4 for good testing)
- Include multiple operation modes

### 4. Run the Test

1. Go to "HeatMap Sheet"
2. Click the **Heatmap Refresh** button (or run `RefreshHeatmap` macro)
3. Observe the results

### 5. Verify the Fix

#### Expected Behavior (AFTER Fix):
✅ All vehicles from "Data Transfer Sheet" should appear in "HeatMap Sheet"
✅ All data values should be transferred correctly
✅ If source has more vehicles than destination can hold, you'll see a warning message
✅ No data should be silently dropped

#### Old Behavior (BEFORE Fix):
❌ DR-prefixed columns were not transferred
❌ Only non-DR columns were transferred (if any)
❌ No warning when data was incomplete
❌ Users had no idea data was missing

### 6. Check Specific Cases

#### Test Case 1: Normal Transfer
- **Setup**: 4 vehicles in source, 7 destination slots available
- **Expected**: All 4 vehicles transferred, no warning
- **Verify**: Count vehicles in HeatMap Sheet = 4

#### Test Case 2: Capacity Warning
- **Setup**: 8 vehicles in source, 7 destination slots available  
- **Expected**: First 7 vehicles transferred, warning message shown
- **Verify**: Warning appears, only 7 vehicles in HeatMap Sheet

#### Test Case 3: DR Columns
- **Setup**: Source sheet has columns named DR1, DR2, DR3 with data
- **Expected**: All DR columns transferred
- **Verify**: All DR data appears in HeatMap Sheet

### 7. Validation Checklist

- [ ] HeatMap Sheet shows all expected vehicles
- [ ] All operation modes are listed
- [ ] Data values match between source and destination
- [ ] No #REF or #VALUE errors
- [ ] Column headers are correct
- [ ] Vehicle count matches expectation
- [ ] Warning appears if source > destination capacity

## Troubleshooting

### Problem: Macro won't run
**Solution**: Enable macros in Excel settings

### Problem: Code looks different after import
**Solution**: Excel may auto-format code. As long as logic is the same, it's OK

### Problem: Still not all data transferred
**Solution**: 
1. Check that "Operation Modes" anchor text exists in both sheets
2. Verify data is in correct columns relative to anchor
3. Check for empty rows/columns that might stop collection
4. Ensure destination sheet has enough vehicle columns

### Problem: Error when clicking button
**Solution**:
1. Check that both "Data Transfer Sheet" and "HeatMap Sheet" exist
2. Verify "Operation Modes" text is present in both sheets
3. Check VBA syntax (no copy/paste errors)

## Success Criteria

The fix is successful when:
1. ✅ All vehicle data from source appears in destination
2. ✅ No silent data loss
3. ✅ Warning shown when capacity exceeded
4. ✅ No VBA errors during execution
5. ✅ Results are reproducible

## Reporting Results

If you encounter any issues or want to report success:
1. Document what you tested
2. Include screenshots if helpful
3. Note any error messages
4. Report in GitHub issues
