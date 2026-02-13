# Fix: Use Operation Codes Instead of Names for Matching

## Issue
User reported that operation modes like "Transition to constant speed" still weren't transferring even after all previous fixes (DR filter, case sensitivity, trailing spaces).

## Root Cause Analysis

The fundamental issue was using **operation mode names** for matching between sheets. This approach is inherently fragile because:

1. **Case sensitivity** - "Transition" vs "transition"
2. **Trailing/leading spaces** - "Transition to constant speed   " vs "Transition to constant speed"
3. **Typos** - "transtion" vs "transition" 
4. **Inconsistent formatting** - Different people entering data differently

## The Better Solution: Use Operation Codes

The user identified that the workbook has a **Mapping Sheet** that contains standardized **operation codes** (e.g., "10070100") mapped to operation names. These codes are:

- **Unique** - Each operation has a distinct code
- **Standardized** - No case sensitivity (all numbers)
- **No spaces** - No trailing/leading space issues
- **No typos** - Codes are referenced from mapping, not manually typed

### Sheet Structure

From the user's screenshots:

**Data Transfer Sheet:**
| Column A (Code) | Column B (Operation Modes) | Column D-E (Vehicle Data) |
|-----------------|----------------------------|---------------------------|
| 10070100        | Transition to constant speed | 8.4, 8.7                  |
| 10120000        | Acceleration                | 8.1, 7.9                  |

**HeatMap Template:**
| Column A (Code) | Column B (Operation Modes) | Column D-F (Vehicle Columns) |
|-----------------|----------------------------|------------------------------|
| [empty initially]| Transition to constant speed | [to be filled]               |
| [empty initially]| Acceleration                | [to be filled]               |

**Note**: The HeatMap Template appears to only have operation names initially, but we need to ensure codes are in column A.

## The Fix

Modified two functions to use **operation codes** (column A, which is `anc.Column - 1`) instead of operation names for matching:

### Change 1: CollectRowLabels

**Before** (used operation names from anchor column):
```vba
For r = anc.row + 2 To lastR
    If Trim$(ws.Cells(r, anc.Column).Value) <> "" Then
        out.Add Trim$(ws.Cells(r, anc.Column).Value)  ' Operation name
        emptyRun = 0
```

**After** (uses operation codes from column before anchor):
```vba
codeCol = anc.Column - 1  ' Column A (codes)

For r = anc.row + 2 To lastR
    If Trim$(ws.Cells(r, anc.Column).Value) <> "" Then
        ' Use operation code instead of name for matching
        out.Add Trim$(ws.Cells(r, codeCol).Value)  ' Operation code
        emptyRun = 0
```

### Change 2: BuildModeIndex

**Before** (used operation names as dictionary keys):
```vba
For r = anc.row + 2 To lastR
    v = Trim$(ws.Cells(r, anc.Column).Value)
    If v <> "" And Not d.Exists(v) Then d.Add v, r  ' Name as key
Next r
```

**After** (uses operation codes as dictionary keys):
```vba
codeCol = anc.Column - 1  ' Column A (codes)

For r = anc.row + 2 To lastR
    v = Trim$(ws.Cells(r, anc.Column).Value)
    If v <> "" Then
        ' Use operation code as dictionary key
        codeVal = Trim$(ws.Cells(r, codeCol).Value)
        If codeVal <> "" And Not d.Exists(codeVal) Then d.Add codeVal, r
    End If
Next r
```

## Impact

✅ **Eliminates all previous issues**:
- No case sensitivity problems (codes are numeric)
- No trailing spaces problems (codes are standardized)
- No typo problems (codes are fixed)
- No localization issues (codes work in any language)

✅ **More reliable**:
- Codes are the source of truth from Mapping Sheet
- Standardized across the organization
- Less prone to user error

✅ **Better maintainability**:
- No need for case-insensitive comparison
- No need for extensive trimming logic
- Simpler code

## Important Note for Users

**Both sheets must have operation codes in column A** (the column before "Operation Modes"):

1. **Data Transfer Sheet** - Already has codes in column A ✅
2. **HeatMap Template** - Must also have codes in column A

If the HeatMap Template doesn't have codes in column A, you need to add them. You can:
- Copy the codes from the Mapping Sheet
- Use a VLOOKUP formula to populate codes based on operation names
- Manually enter the codes

## Files Changed
- **HeatMap.bas** (lines 297-332): Modified `CollectRowLabels` and `BuildModeIndex` functions

## Commit
This fix addresses the fundamental matching issue by using operation codes instead of names.
