# Alternative Fix: Transfer All Source Modes

## Current Design vs Alternative Design

### Current Design (Template-Based)
- HeatMap Sheet acts as a template with predefined operation modes
- Only modes that exist in HeatMap Sheet get data filled in
- Modes in Data Transfer Sheet that aren't in HeatMap Sheet are ignored

### Alternative Design (Source-Driven)
- All operation modes from Data Transfer Sheet are transferred
- HeatMap Sheet gets populated with whatever is in the source
- No predefined template structure

## When to Use Alternative Design

Use the alternative design if:
- You want ALL operation modes from Data Transfer Sheet to appear in HeatMap Sheet
- The set of operation modes changes frequently
- You don't want to manually maintain a template

Keep the current design if:
- HeatMap Sheet should only show specific, predefined operation modes
- You want consistent structure across different data sets
- The template defines the expected layout

## Code Changes for Alternative Design

### Option 1: Modify Existing Code (Major Change)

This would require significant changes to the transfer logic:

1. Change the loop to iterate over source modes instead of destination modes
2. Dynamically add rows to HeatMap Sheet for new modes
3. Handle formatting and structure maintenance

**Location**: Lines 108-121 in RefreshHeatmap function

**Current**:
```vba
'--- Fill ALL vehicle data correctly ---
For i = 1 To tModes.count
    If sModeIx.Exists(tModes(i)) Then
        rS = sModeIx(tModes(i))
        rt = tA.row + 1 + i

        For j = 1 To n
            v = wsS.Cells(rS, sVehCol(j)).Value
            If HasValue(v) Then
                wsT.Cells(rt, tVehCols(j)).Value = CDbl(v)
            End If
        Next j
    End If
Next i
```

**Alternative (iterate source modes)**:
```vba
'--- Fill ALL vehicle data from source ---
Dim sourceMode As Variant
Dim targetRow As Long
targetRow = tA.row + 2  ' Start after header

For Each sourceMode In sModeIx.Keys()
    ' Find or create row for this mode in target
    Dim foundRow As Long
    foundRow = FindModeRow(wsT, tA, sourceMode)
    
    If foundRow = 0 Then
        ' Mode doesn't exist in target, add it
        wsT.Cells(targetRow, tA.Column).Value = sourceMode
        foundRow = targetRow
        targetRow = targetRow + 1
    End If
    
    ' Transfer data
    rS = sModeIx(sourceMode)
    For j = 1 To n
        v = wsS.Cells(rS, sVehCol(j)).Value
        If HasValue(v) Then
            wsT.Cells(foundRow, tVehCols(j)).Value = CDbl(v)
        End If
    Next j
Next sourceMode
```

**Additional function needed**:
```vba
Private Function FindModeRow(ws As Worksheet, anc As Range, modeName As Variant) As Long
    Dim r As Long, lastR As Long
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.Count, anc.Column).End(xlUp).Row
    On Error GoTo 0
    
    For r = anc.row + 2 To lastR
        If StrComp(Trim$(ws.Cells(r, anc.Column).Value), Trim$(modeName), vbTextCompare) = 0 Then
            FindModeRow = r
            Exit Function
        End If
    Next r
    
    FindModeRow = 0  ' Not found
End Function
```

### Option 2: Hybrid Approach (Recommended)

Keep the template-based design but add a warning when source modes are missing:

**Add after line 71**:
```vba
'--- Check for missing modes in target ---
Dim missingModes As String
Dim sourceMode As Variant
missingModes = ""

For Each sourceMode In sModeIx.Keys()
    Dim found As Boolean
    found = False
    
    For i = 1 To tModes.Count
        If StrComp(Trim$(tModes(i)), Trim$(sourceMode), vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        If missingModes <> "" Then missingModes = missingModes & ", "
        missingModes = missingModes & sourceMode
    End If
Next sourceMode

If missingModes <> "" Then
    MsgBox "Warning: The following operation modes exist in Data Transfer Sheet but not in HeatMap Sheet:" & vbCrLf & vbCrLf & _
           missingModes & vbCrLf & vbCrLf & _
           "These modes will not be transferred. Add them to HeatMap Sheet if needed.", _
           vbExclamation, "Missing Operation Modes"
End If
```

## Recommendation

**For most use cases**: Use Option 2 (Hybrid Approach)
- Maintains the template-based design
- Alerts user when modes are missing
- Minimal code changes
- User can decide whether to add the modes

**For dynamic data**: Use Option 1 (Source-Driven)
- Best if operation modes change frequently
- Requires more extensive testing
- Changes the fundamental design

## Implementation

If you want either option implemented, please confirm:
1. Which approach you prefer
2. Whether the HeatMap Sheet should maintain a fixed template structure
3. Whether new modes should be automatically added or require manual addition

I can implement the chosen approach and test it thoroughly.
