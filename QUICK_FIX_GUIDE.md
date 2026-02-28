# Quick Fix Guide

## Problems
1. After clicking evaluation button and selecting vehicles, tested responsiveness values are not being read correctly
2. AVL scores are not reading from the tested vehicle's column in HeatMap sheet

## Solutions
Two bugs need to be fixed in the Evaluation VBA module:
1. Wrong column offset for responsiveness (line 98)
2. Hardcoded AVL column instead of dynamic lookup (GetTestedAVL function)

## Quick Fix Steps

### Option 1: Import Fixed Module (Easiest - 2 minutes)

1. **Open Excel File**
   - Open: `AVLDrive_Heatmap_Tool version_4 (2).xlsm`

2. **Open VBA Editor**
   - Press: `Alt + F11`

3. **Remove Old Module**
   - Right-click on **Evaluation** module
   - Select: **Remove Evaluation**
   - Click **No** when asked to export

4. **Import Fixed Module**
   - Right-click in Project Explorer
   - Select: **Import File...**
   - Choose: `Evaluation_Module_FIXED.bas`

5. **Save**
   - Press: `Ctrl + S`
   - Close VBA Editor
   - Save Excel file

### Option 2: Manual Edits (5-10 minutes)

1. **Open Excel File**
   - Open: `AVLDrive_Heatmap_Tool version_4 (2).xlsm`

2. **Open VBA Editor**
   - Press: `Alt + F11`

3. **Find the Evaluation Module**
   - Double-click: **Evaluation** in Project Explorer

4. **Fix #1: Responsiveness Column (Line 98)**
   - Press: `Ctrl + F` (Find)
   - Search for: `testedCol + 6`
   - You'll find this line:
     ```vba
     respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).Value)
     ```
   - Change `+ 6` to `+ 7`:
     ```vba
     respTested = ToDbl(wsSheet1.Cells(i, testedCol + 7).Value)
     ```

5. **Fix #2: AVL Function Call (Line 88)**
   - Search for: `GetTestedAVL(wsHeatmap, opCode)`
   - Change to:
     ```vba
     testedAVL = GetTestedAVL(wsHeatmap, opCode, testedCarName)
     ```

6. **Fix #3: AVL Function Declaration (Around Line 300)**
   - Find: `Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant) As Double`
   - Change to:
     ```vba
     Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant, testedCarName As String) As Double
     ```

7. **Fix #4: AVL Column Detection (Around Line 307)**
   - Find the line: `avlCol = 8`
   - Replace the entire hardcoded section with dynamic column finding:
     ```vba
     ' Find the column for the tested vehicle in HeatMap sheet
     avlCol = 0
     lastCol = wsHeatmap.Cells(2, wsHeatmap.Columns.count).End(xlToLeft).Column
     
     For col = 1 To lastCol
         If Trim(CStr(wsHeatmap.Cells(2, col).Value)) = Trim(testedCarName) Then
             avlCol = col
             Exit For
         End If
     Next col
     
     ' If not found, default to column 8 for backward compatibility
     If avlCol = 0 Then avlCol = 8
     ```
   - Don't forget to add `Dim col As Long` and `Dim lastCol As Long` to the variable declarations at the top of the function

8. **Save**
   - Press: `Ctrl + S`
   - Close VBA Editor
   - Save Excel file

## What This Fixes
- ✓ Tested responsiveness values will now be read from the correct column
- ✓ AVL scores will be read from the tested vehicle's specific column in HeatMap sheet
- ✓ Evaluation results will be accurate for any vehicle combination
- ✓ All status colors (GREEN/YELLOW/RED) will be correctly calculated

## Need Help?
See **BUG_FIX_DOCUMENTATION.md** for detailed information including code examples and explanations.
