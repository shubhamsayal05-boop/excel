# Fix for Data Transfer Issue

## Problem
Not all data from the "Data Transfer Sheet" was being transferred to the "HeatMap Sheet" after clicking the heatmap button.

## Root Cause
The issue was in the `HeatMap.bas` VBA module, specifically in two functions:

1. **`CollectHeaders` function** (line 269-271): This function was filtering out columns with "DR" in the header
2. **`CollectHeaderCols` function** (line 285-287): This function was also filtering out "DR" columns

The problem:
- The destination sheet uses "DR" markers to identify vehicle columns (see `CollectDestVehicleCols`)
- However, the source data collection functions were **excluding** "DR" columns
- This caused fewer source columns to be collected than actually existed
- As a result, the line `n = WorksheetFunction.Min(tVehCols.count, sVehHdr.count)` would artificially limit the data transfer

## Changes Made

### 1. Fixed `CollectHeaders` Function
**Before:**
```vba
For c = anc.Column + 1 To lastC
    If Trim$(ws.Cells(anc.row, c).Value) <> "" _
       And UCase$(ws.Cells(anc.row, c).Value) <> "DR" Then
        out.Add ws.Cells(anc.row, c).Value
    End If
Next c
```

**After:**
```vba
For c = anc.Column + 1 To lastC
    If Trim$(ws.Cells(anc.row, c).Value) <> "" Then
        out.Add ws.Cells(anc.row, c).Value
    End If
Next c
```

### 2. Fixed `CollectHeaderCols` Function
**Before:**
```vba
For c = anc.Column + 1 To lastC
    If Trim$(ws.Cells(anc.row, c).Value) <> "" _
       And UCase$(ws.Cells(anc.row, c).Value) <> "DR" Then
        out.Add c
    End If
Next c
```

**After:**
```vba
For c = anc.Column + 1 To lastC
    If Trim$(ws.Cells(anc.row, c).Value) <> "" Then
        out.Add c
    End If
Next c
```

### 3. Added Warning Message
Added a user-friendly warning when the source data exceeds destination capacity:

```vba
'--- Warn if source data exceeds destination capacity ---
If sVehHdr.count > tVehCols.count Then
    MsgBox "Warning: Data Transfer Sheet has " & sVehHdr.count & " vehicles, but HeatMap Sheet can only accommodate " & tVehCols.count & " vehicles." & vbCrLf & _
           "Only the first " & n & " vehicles will be transferred.", vbExclamation, "Data Capacity Warning"
End If
```

## How to Apply the Fix

Since directly updating VBA code in an xlsm file programmatically is complex, you'll need to manually import the updated module:

### Option 1: Replace Code Directly (Recommended)
1. Open `AVLDrive_Heatmap_Tool version_4 (2).xlsm` in Microsoft Excel
2. Press `Alt+F11` to open the VBA Editor
3. In the Project Explorer, find and double-click on `HeatMap` module
4. Select all code (Ctrl+A) and delete it
5. Open `HeatMap.bas` in a text editor
6. Copy all the code
7. Paste it into the HeatMap module in Excel
8. Save the file (Ctrl+S)
9. Close VBA Editor and save the Excel file

### Option 2: Import Module
1. Open `AVLDrive_Heatmap_Tool version_4 (2).xlsm` in Microsoft Excel
2. Press `Alt+F11` to open the VBA Editor
3. Right-click on the `HeatMap` module in Project Explorer
4. Select "Remove HeatMap"
5. Choose "No" when asked to export (we already have the updated version)
6. File → Import File
7. Select `HeatMap.bas`
8. Save the file

## Testing the Fix
After applying the changes:
1. Open the Excel file
2. Ensure "Data Transfer Sheet" has multiple vehicles with data
3. Click the heatmap refresh button
4. Verify that ALL vehicle data from "Data Transfer Sheet" is now transferred to "HeatMap Sheet"
5. If the source has more vehicles than the destination can handle, you should see a warning message

## Technical Details
- The fix ensures all non-empty columns in the source sheet are collected
- The "DR" filter was incorrectly excluding vehicle data
- The minimum calculation (`n = Min(tVehCols.count, sVehHdr.count)`) is correct to prevent writing beyond destination capacity
- The new warning message alerts users when data truncation occurs
