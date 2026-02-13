# Bug Fix Documentation: Evaluation Button Issues

## Problems Statement
Two issues were identified after clicking the evaluation button and selecting target and tested vehicles:
1. **Tested responsiveness values were not being read correctly** - wrong column offset
2. **Tested AVL scores were not being read correctly** - hardcoded column instead of tested vehicle's column

## Root Cause Analysis

### Bug #1: Wrong Responsiveness Column Offset
The bug was in the **Evaluation.bas** VBA module, line 98.

**The Bug:**
```vba
' Resp columns are 7 positions after Driv
respTarget = ToDbl(wsSheet1.Cells(i, targetCol + 7).Value)
respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).Value)  ' ← BUG: Should be +7, not +6
```

**Why This Is Wrong:**
- The comment clearly states "Resp columns are 7 positions after Driv"
- `respTarget` correctly uses `targetCol + 7`
- `respTested` incorrectly uses `testedCol + 6` instead of `testedCol + 7`
- This causes the tested responsiveness value to be read from the wrong column

### Bug #2: Hardcoded AVL Column
The `GetTestedAVL` function was hardcoded to read from column 8 of the HeatMap sheet.

**The Bug:**
```vba
Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant) As Double
    avlCol = 8  ' ← BUG: Hardcoded column, should use tested vehicle's column
```

**Why This Is Wrong:**
- The HeatMap sheet has vehicle-specific columns with AVL scores
- Each vehicle has its own column in the HeatMap sheet
- The function should read from the tested vehicle's specific column, not a fixed column 8
- This causes the evaluation to always read AVL scores from the same column regardless of which vehicle is being tested

## The Fixes

### Fix #1: Responsiveness Column Offset
Change line 98 from:
```vba
respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).Value)
```

To:
```vba
respTested = ToDbl(wsSheet1.Cells(i, testedCol + 7).Value)
```

### Fix #2: Dynamic AVL Column Reading
Modified the `GetTestedAVL` function to:
1. Accept the tested vehicle name as a parameter
2. Find the tested vehicle's column in the HeatMap sheet (by searching row 2)
3. Read the AVL score from that specific column

**Before:**
```vba
Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant) As Double
    avlCol = 8  ' Hardcoded
    ...
End Function
```

**After:**
```vba
Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant, testedCarName As String) As Double
    ' Find the column for the tested vehicle in HeatMap sheet
    avlCol = 0
    For col = 1 To lastCol
        If Trim(CStr(wsHeatmap.Cells(2, col).Value)) = Trim(testedCarName) Then
            avlCol = col
            Exit For
        End If
    Next col
    
    ' If not found, default to column 8 for backward compatibility
    If avlCol = 0 Then avlCol = 8
    ...
End Function
```

And updated the function call in line 88:
```vba
testedAVL = GetTestedAVL(wsHeatmap, opCode, testedCarName)
```

## How to Apply the Fixes

### Option 1: Manual Fix in Excel (Recommended)
1. Open the file **AVLDrive_Heatmap_Tool version_4 (2).xlsm** in Microsoft Excel
2. Press **Alt+F11** to open the VBA Editor
3. In the Project Explorer, find and double-click the **Evaluation** module

**Fix #1: Responsiveness Column (Line 98)**
4. Search for "testedCol + 6" (around line 98)
5. Change `+ 6` to `+ 7`

**Fix #2: AVL Score Reading (GetTestedAVL function, around line 300)**
6. Find the `GetTestedAVL` function declaration:
   ```vba
   Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant) As Double
   ```
7. Change it to:
   ```vba
   Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant, testedCarName As String) As Double
   ```
8. Find the line `avlCol = 8` (around line 307)
9. Replace the entire section with the new dynamic column finding code (see fixed module)
10. Find the function call around line 88: `testedAVL = GetTestedAVL(wsHeatmap, opCode)`
11. Change it to: `testedAVL = GetTestedAVL(wsHeatmap, opCode, testedCarName)`

12. Press **Ctrl+S** to save
13. Close the VBA Editor
14. Save the Excel file

### Option 2: Import Fixed Module (Easier)
1. Open the file **AVLDrive_Heatmap_Tool version_4 (2).xlsm** in Microsoft Excel
2. Press **Alt+F11** to open the VBA Editor
3. In the Project Explorer, right-click on the **Evaluation** module and select **Remove Evaluation**
4. Click **No** when asked if you want to export before removing (or export for backup)
5. Right-click on any item in the Project Explorer and select **Import File...**
6. Select the **Evaluation_Module_FIXED.bas** file from this repository
7. Press **Ctrl+S** to save
8. Close the VBA Editor
9. Save the Excel file

## Files Included
- **Evaluation_Module_ORIGINAL.bas** - The original VBA module with both bugs (for reference)
- **Evaluation_Module_FIXED.bas** - The corrected VBA module ready to import (both fixes applied)
- **BUG_FIX_DOCUMENTATION.md** - This documentation file

## Impact
These bugs caused incorrect evaluation results because:

**Bug #1 Impact:**
1. The tested responsiveness values were read from the wrong column
2. This led to incorrect comparisons between target and tested responsiveness
3. The final evaluation status could be incorrect (showing GREEN/YELLOW/RED when it should show different colors)

**Bug #2 Impact:**
1. The AVL scores were always read from column 8 of the HeatMap sheet
2. This meant all evaluations used the same AVL scores regardless of which vehicle was being tested
3. Different vehicles have different AVL scores, so this caused incorrect evaluation results
4. The "Tested AVL" column in results would show wrong values for any vehicle other than the one in column 8

## Verification
After applying both fixes:
1. Run the evaluation button
2. Select target and tested vehicles (try different combinations)
3. Verify that:
   - The "Tested AVL" column shows AVL scores matching the tested vehicle in the HeatMap sheet
   - The "Resp Tested" column shows the correct responsiveness values
   - Compare with the original data in Sheet1 and HeatMap sheet to ensure columns match properly

## Technical Details

### Bug #1: Responsiveness Column Offset
- **File:** AVLDrive_Heatmap_Tool version_4 (2).xlsm
- **Module:** Evaluation.bas
- **Function:** EvaluateAVLStatus()
- **Line:** 98
- **Change:** `testedCol + 6` → `testedCol + 7`

### Bug #2: AVL Score Reading
- **File:** AVLDrive_Heatmap_Tool version_4 (2).xlsm
- **Module:** Evaluation.bas
- **Function:** GetTestedAVL()
- **Lines:** 88, 300-340
- **Changes:**
  - Added `testedCarName` parameter to function
  - Added dynamic column detection to find tested vehicle's column in HeatMap sheet
  - Updated function call to pass tested vehicle name

## Notes
- The drivability values (drivTarget, drivTested) are read correctly using `targetCol` and `testedCol`
- The responsiveness values should use the same offset (+7) since they are in parallel columns 7 positions to the right
- AVL scores are vehicle-specific and stored in the HeatMap sheet with vehicle names in row 2
- The fixed code maintains backward compatibility by defaulting to column 8 if the vehicle name is not found
