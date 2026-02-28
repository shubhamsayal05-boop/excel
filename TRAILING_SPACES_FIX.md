# Fix: Trailing Spaces in Operation Mode Names

## Issue Discovered

User reported that "Transition to constant speed" and other operation modes exist in both the Data Transfer Sheet and HeatMap Sheet, but data was still not transferring after the case-sensitivity fix.

## Root Cause

**Inconsistent trimming of operation mode names** between two functions:

### The Bug

1. **`CollectRowLabels`** (HeatMap Sheet - destination):
   - Line 305: Checked if trimmed value is not empty: `If Trim$(ws.Cells(r, anc.Column).Value) <> ""`
   - Line 306: Added **untrimmed** value to collection: `out.Add ws.Cells(r, anc.Column).Value`
   - Result: Mode names with trailing/leading spaces were added as-is

2. **`BuildModeIndex`** (Data Transfer Sheet - source):
   - Line 328: Added **trimmed** value to dictionary: `v = Trim$(ws.Cells(r, anc.Column).Value)`
   - Result: Mode names were trimmed before adding

### Example of the Problem

```
HeatMap Sheet (CollectRowLabels):
  "Transition to constant speed   " (with trailing spaces)

Data Transfer Sheet (BuildModeIndex):
  "Transition to constant speed" (trimmed)

Dictionary lookup:
  sModeIx.Exists("Transition to constant speed   ")  → FALSE
  
Result: Data not transferred!
```

Even with case-insensitive matching (`d.CompareMode = vbTextCompare`), the trailing spaces caused a mismatch.

## The Fix

Applied `Trim$()` consistently in `CollectRowLabels` function:

### Before (Broken)
```vba
For r = anc.row + 2 To lastR
    If Trim$(ws.Cells(r, anc.Column).Value) <> "" Then
        out.Add ws.Cells(r, anc.Column).Value  ' No trim - keeps spaces!
        emptyRun = 0
```

### After (Fixed)
```vba
For r = anc.row + 2 To lastR
    If Trim$(ws.Cells(r, anc.Column).Value) <> "" Then
        out.Add Trim$(ws.Cells(r, anc.Column).Value)  ' Apply Trim$ for consistency
        emptyRun = 0
```

## Impact

Now both functions trim the mode names before processing:
- ✅ `CollectRowLabels`: Adds trimmed mode names
- ✅ `BuildModeIndex`: Adds trimmed mode names
- ✅ Mode matching works even with trailing/leading spaces in Excel cells
- ✅ "Transition to constant speed" and other modes now transfer correctly

## Why This Happened

Excel cells can contain leading/trailing spaces that aren't visible. When users:
- Copy/paste data
- Type with extra spaces
- Import from other sources

These hidden spaces can break exact string matching. The fix ensures all mode names are normalized (trimmed) before comparison.

## Files Changed

- **HeatMap.bas** (line 306): Added `Trim$()` to normalize operation mode names

## Commit

This fix is included in the commit that addresses the trailing spaces issue.
