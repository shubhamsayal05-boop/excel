# Quick Fix Guide

## Problem
After clicking evaluation button and selecting vehicles, tested AVL score and target responsiveness values are not being read correctly.

## Solution
There's a bug in line 98 of the Evaluation VBA module where it uses the wrong column offset.

## Quick Fix Steps

1. **Open Excel File**
   - Open: `AVLDrive_Heatmap_Tool version_4 (2).xlsm`

2. **Open VBA Editor**
   - Press: `Alt + F11`

3. **Find the Bug**
   - In Project Explorer, double-click: **Evaluation**
   - Press: `Ctrl + F` (Find)
   - Search for: `testedCol + 6`
   - You'll find this line:
     ```vba
     respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).Value)
     ```

4. **Fix It**
   - Change `+ 6` to `+ 7`
   - The line should now read:
     ```vba
     respTested = ToDbl(wsSheet1.Cells(i, testedCol + 7).Value)
     ```

5. **Save**
   - Press: `Ctrl + S`
   - Close VBA Editor
   - Save Excel file

## What This Fixes
- ✓ Tested responsiveness values will now be read from the correct column
- ✓ Evaluation results will be accurate
- ✓ All status colors (GREEN/YELLOW/RED) will be correctly calculated

## Need Help?
See **BUG_FIX_DOCUMENTATION.md** for detailed information and alternative fix methods.
