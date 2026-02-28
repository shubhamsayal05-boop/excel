# Diagnostic Guide for "transition to constant speed" Issue

## Current Behavior

The VBA code is designed so that the **HeatMap Sheet acts as a template**. The data transfer process works as follows:

1. Collects operation modes from the **HeatMap Sheet** (destination)
2. Creates an index of operation modes from the **Data Transfer Sheet** (source)
3. For each mode in the HeatMap Sheet, looks it up in the Data Transfer Sheet and transfers the data

## Why "transition to constant speed" Might Not Appear

### Scenario 1: Mode Not in HeatMap Sheet Template
If "transition to constant speed" exists in the Data Transfer Sheet but is NOT listed in the HeatMap Sheet's "Operation Modes" column, it will NOT be transferred.

**Current Code (lines 109-121)**:
```vba
For i = 1 To tModes.count        ' Iterates HeatMap Sheet modes
    If sModeIx.Exists(tModes(i)) Then
        ' Transfer data
    End If
Next i
```

**Why**: The loop iterates through HeatMap Sheet modes only. If a mode exists only in the Data Transfer Sheet, we never check for it.

### Scenario 2: Name Mismatch
Even with case-insensitive matching, there could be issues:
- Extra spaces: "transition to constant speed " vs "transition to constant speed"
- Hidden characters (non-breaking spaces, tabs, etc.)
- Typos: "transtion to constant speed" vs "transition to constant speed"

## How to Diagnose

### Check 1: Is the mode in HeatMap Sheet?
1. Open the Excel file
2. Go to "HeatMap Sheet"
3. Look in the "Operation Modes" column (usually column B or C)
4. Scroll through and check if "transition to constant speed" is listed

**If YES**: Continue to Check 2  
**If NO**: See Solution 1 below

### Check 2: Are the names exactly the same?
1. In HeatMap Sheet, click on the cell with "transition to constant speed"
2. Copy it (Ctrl+C)
3. Paste into Notepad
4. In Data Transfer Sheet, click on the cell with "transition to constant speed"
5. Copy it (Ctrl+C)
6. Paste into Notepad
7. Compare - are they EXACTLY the same? (including spaces)

**If YES**: Continue to Check 3  
**If NO**: See Solution 2 below

### Check 3: Run with debugging
Add debug output to see what's happening:
1. In VBA Editor, add this line after line 71:
   ```vba
   Debug.Print "Source modes: " & sModeIx.Count & ", Target modes: " & tModes.Count
   ```
2. Add this line inside the loop at line 110:
   ```vba
   Debug.Print "Checking: '" & tModes(i) & "' - Exists: " & sModeIx.Exists(tModes(i))
   ```
3. Run the heatmap button
4. Press Ctrl+G to open Immediate window
5. Check the output

## Solutions

### Solution 1: Mode Not in HeatMap Sheet Template

**Option A: Add the mode to HeatMap Sheet manually**
1. Go to HeatMap Sheet
2. Find the last operation mode in the "Operation Modes" column
3. Add "transition to constant speed" in the next row
4. Run the heatmap button again

**Option B: Modify code to auto-add missing modes**
This would require changing the core logic to iterate through source modes instead of destination modes. This is a more complex change that changes the template-based design.

### Solution 2: Name Mismatch

**Fix spacing issues:**
1. In both sheets, click on the cell with the mode name
2. Press F2 to edit
3. Remove any extra spaces at the beginning or end
4. Press Enter
5. Try the heatmap button again

**Alternative: Use TRIM function**
The code already uses `Trim$()` but hidden characters might remain. You could:
1. In Data Transfer Sheet, add a helper column
2. Use formula: `=TRIM(CLEAN(A2))` where A2 is the mode name
3. Use this cleaned column for the mode matching

### Solution 3: Debug Output

If the above don't work, we need to see what's actually being compared. Add this to the VBA code:

```vba
' After line 71, add:
Dim debugMsg As String
debugMsg = "Modes in Data Transfer Sheet:" & vbCrLf
For Each Key In sModeIx.Keys()
    debugMsg = debugMsg & "  - '" & Key & "'" & vbCrLf
Next Key
debugMsg = debugMsg & vbCrLf & "Modes in HeatMap Sheet:" & vbCrLf
For i = 1 To tModes.Count
    debugMsg = debugMsg & "  - '" & tModes(i) & "'" & vbCrLf
Next i
MsgBox debugMsg, vbInformation, "Debug: Mode Names"
```

This will show you exactly what mode names are being collected from each sheet.

## Most Likely Issue

Based on the description, the most likely issue is **Scenario 1**: "transition to constant speed" exists in the Data Transfer Sheet but is not in the HeatMap Sheet template.

**Quick Fix**: Manually add "transition to constant speed" to the HeatMap Sheet's "Operation Modes" column, then run the heatmap button.

## Need Different Behavior?

If you want the code to automatically include ALL operation modes from the Data Transfer Sheet (not just those in the template), that would require modifying the core transfer logic. Let me know if that's the desired behavior.
