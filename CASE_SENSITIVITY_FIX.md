# Case Sensitivity Fix for Operation Mode Matching

## Issue Reported
User reported that "transition to constant speed" has values in the Data Transfer Sheet but didn't get transferred to the HeatMap Sheet.

## Root Cause
The `BuildModeIndex` function creates a VBA Dictionary to match operation modes between sheets. By default, VBA's Scripting.Dictionary is **case-sensitive**, which means:

- "transition to constant speed" ≠ "Transition to Constant Speed"
- "TRANSITION TO CONSTANT SPEED" ≠ "transition to constant speed"

If the operation mode names have different capitalization between the Data Transfer Sheet and HeatMap Sheet, the dictionary lookup fails and the data won't transfer.

## The Fix
Added `d.CompareMode = vbTextCompare` to make the dictionary **case-insensitive**.

### Before:
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

### After:
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

## Impact
Now operation modes will match regardless of capitalization:
- ✅ "transition to constant speed" matches "Transition to Constant Speed"
- ✅ "TRANSITION TO CONSTANT SPEED" matches "transition to constant speed"
- ✅ "Transition To Constant Speed" matches "TRANSITION TO CONSTANT SPEED"

## Testing
After importing the updated HeatMap.bas:
1. Ensure "transition to constant speed" (or any case variant) exists in both sheets
2. Run the heatmap refresh
3. Verify the data transfers correctly regardless of case differences

## Note
This fix assumes the operation mode names are identical except for case. If there are other differences (extra spaces, typos, etc.), the modes still won't match. The names must be the same when case is ignored.
