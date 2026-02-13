# Code Changes - Detailed Comparison

## Summary
Four changes were made to fix the data transfer issue:
1. Fixed `CollectHeaders` function - removed DR filter
2. Fixed `CollectHeaderCols` function - removed DR filter  
3. Added capacity warning in `RefreshHeatmap` - user notification
4. Fixed `BuildModeIndex` function - made case-insensitive for operation mode matching

---

## Change 1: CollectHeaders Function

**Location**: Line 267-279 in HeatMap.bas

### BEFORE (Broken)
```vba
Public Function CollectHeaders(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" _
           And UCase$(ws.Cells(anc.row, c).Value) <> "DR" Then  ← PROBLEM: Excludes DR
            out.Add ws.Cells(anc.row, c).Value
        End If
    Next c
    Set CollectHeaders = out
End Function
```

### AFTER (Fixed)
```vba
Public Function CollectHeaders(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" Then  ← FIXED: Includes all non-empty
            out.Add ws.Cells(anc.row, c).Value
        End If
    Next c
    Set CollectHeaders = out
End Function
```

**What Changed**: 
- ❌ Removed: `And UCase$(ws.Cells(anc.row, c).Value) <> "DR"`
- ✅ Now includes all non-empty column headers, including DR columns

**Impact**: 
- Source vehicle headers are now correctly collected
- Fixes the root cause of missing data

---

## Change 2: CollectHeaderCols Function

**Location**: Line 282-294 in HeatMap.bas

### BEFORE (Broken)
```vba
Public Function CollectHeaderCols(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" _
           And UCase$(ws.Cells(anc.row, c).Value) <> "DR" Then  ← PROBLEM: Excludes DR
            out.Add c
        End If
    Next c
    Set CollectHeaderCols = out
End Function
```

### AFTER (Fixed)
```vba
Public Function CollectHeaderCols(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" Then  ← FIXED: Includes all non-empty
            out.Add c
        End If
    Next c
    Set CollectHeaderCols = out
End Function
```

**What Changed**:
- ❌ Removed: `And UCase$(ws.Cells(anc.row, c).Value) <> "DR"`
- ✅ Now includes all non-empty column indices, including DR columns

**Impact**:
- Source vehicle column indices are now correctly collected
- Ensures the data copy loop has correct column references

---

## Change 3: Added Capacity Warning

**Location**: Line 78-84 in HeatMap.bas (in RefreshHeatmap function)

### BEFORE (No Warning)
```vba
    If tVehCols.count = 0 Or sVehCol.count = 0 Or tModes.count = 0 Then
        MsgBox "No data found to process.", vbInformation
        GoTo CleanExit
    End If

    n = WorksheetFunction.Min(tVehCols.count, sVehHdr.count)

    '--- Vehicle header rows ---
```

### AFTER (With Warning)
```vba
    If tVehCols.count = 0 Or sVehCol.count = 0 Or tModes.count = 0 Then
        MsgBox "No data found to process.", vbInformation
        GoTo CleanExit
    End If

    n = WorksheetFunction.Min(tVehCols.count, sVehHdr.count)
    
    '--- Warn if source data exceeds destination capacity ---
    If sVehHdr.count > tVehCols.count Then
        MsgBox "Warning: Data Transfer Sheet has " & sVehHdr.count & " vehicles, but HeatMap Sheet can only accommodate " & tVehCols.count & " vehicles." & vbCrLf & _
               "Only the first " & n & " vehicles will be transferred.", vbExclamation, "Data Capacity Warning"
    End If

    '--- Vehicle header rows ---
```

**What Changed**:
- ✅ Added: User-friendly warning when source exceeds destination capacity
- ✅ Shows exact counts: source vehicle count vs destination capacity
- ✅ Informs user that only first n vehicles will be transferred

**Impact**:
- Users are now aware when data is being truncated
- Prevents silent data loss
- Helps users understand they may need to expand destination capacity

---

## No Other Changes

All other code remains **exactly the same**:
- ✅ Main data transfer loop unchanged (lines 103-115)
- ✅ Row/column collection logic unchanged
- ✅ Formatting functions unchanged
- ✅ Error handling unchanged
- ✅ Protection/unprotection logic unchanged

---

## Change 4: BuildModeIndex Function (Case Sensitivity Fix)

**Location**: Line 316-332 in HeatMap.bas

### BEFORE (Case-Sensitive)
```vba
Public Function BuildModeIndex(ws As Worksheet, anc As Range) As Object
    Dim d As Object:  Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, v, lastR As Long
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0
    
    For r = anc.row + 2 To lastR
        v = Trim$(ws.Cells(r, anc.Column).Value)
        If v <> "" And Not d.Exists(v) Then d.Add v, r
    Next r
    Set BuildModeIndex = d
End Function
```

### AFTER (Case-Insensitive)
```vba
Public Function BuildModeIndex(ws As Worksheet, anc As Range) As Object
    Dim d As Object:  Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, v, lastR As Long
    
    '*** Make dictionary case-insensitive for mode matching ***
    d.CompareMode = vbTextCompare
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0
    
    For r = anc.row + 2 To lastR
        v = Trim$(ws.Cells(r, anc.Column).Value)
        If v <> "" And Not d.Exists(v) Then d.Add v, r
    Next r
    Set BuildModeIndex = d
End Function
```

**What Changed**: 
- ✅ Added: `d.CompareMode = vbTextCompare` to make dictionary case-insensitive

**Impact**: 
- Operation mode names now match regardless of capitalization
- "transition to constant speed" matches "Transition to Constant Speed"
- Fixes issue where modes with different casing wouldn't transfer data

---

## Why These Changes Fix the Issue

### The Original Bug Flow:
1. `CollectHeaders` skips DR columns → returns empty or incomplete list
2. `CollectHeaderCols` skips DR columns → returns empty or incomplete list
3. `n = Min(tVehCols.count, sVehHdr.count)` → n becomes 0 or very small
4. Data copy loop: `For j = 1 To n` → copies 0 or few vehicles
5. Mode matching is case-sensitive → "transition to constant speed" ≠ "Transition to Constant Speed"
6. Result: **Missing data**

### The Fixed Flow:
1. `CollectHeaders` includes DR columns → returns complete list ✅
2. `CollectHeaderCols` includes DR columns → returns complete list ✅
3. `n = Min(tVehCols.count, sVehHdr.count)` → n is correct count ✅
4. Data copy loop: `For j = 1 To n` → copies all available vehicles ✅
5. Warning shown if truncation occurs ✅
6. Mode matching is case-insensitive → matches regardless of capitalization ✅
7. Result: **All data transferred** ✅

---

## Lines Changed
- Line 269-271: Removed DR filter condition (CollectHeaders)
- Line 285-287: Removed DR filter condition (CollectHeaderCols)
- Line 80-84: Added capacity warning (RefreshHeatmap)
- Line 320-321: Added case-insensitive mode matching (BuildModeIndex)

**Total lines changed**: 11 lines
**Functions modified**: 3 functions
**New features added**: 1 warning message

This is a **minimal, surgical fix** that addresses the root cause without changing unrelated code.
