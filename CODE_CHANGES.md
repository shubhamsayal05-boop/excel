# Code Changes - Detailed Comparison

## Summary
Three changes were made to fix the data transfer issue:
1. Fixed `CollectHeaders` function - removed DR filter
2. Fixed `CollectHeaderCols` function - removed DR filter  
3. Added capacity warning in `RefreshHeatmap` - user notification

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

## Why These Changes Fix the Issue

### The Original Bug Flow:
1. `CollectHeaders` skips DR columns → returns empty or incomplete list
2. `CollectHeaderCols` skips DR columns → returns empty or incomplete list
3. `n = Min(tVehCols.count, sVehHdr.count)` → n becomes 0 or very small
4. Data copy loop: `For j = 1 To n` → copies 0 or few vehicles
5. Result: **Missing data**

### The Fixed Flow:
1. `CollectHeaders` includes DR columns → returns complete list ✅
2. `CollectHeaderCols` includes DR columns → returns complete list ✅
3. `n = Min(tVehCols.count, sVehHdr.count)` → n is correct count ✅
4. Data copy loop: `For j = 1 To n` → copies all available vehicles ✅
5. Warning shown if truncation occurs ✅
6. Result: **All data transferred** ✅

---

## Lines Changed
- Line 269-271: Removed DR filter condition
- Line 285-287: Removed DR filter condition
- Line 80-84: Added capacity warning

**Total lines changed**: 8 lines
**Functions modified**: 2 functions
**New features added**: 1 warning message

This is a **minimal, surgical fix** that addresses the root cause without changing unrelated code.
