# VBA Source Files

This folder contains the complete, readable VBA source code for every module
embedded in `AVLDrive_Heatmap_Tool version_5 (1).xlsm`.

> **These files are for reference and review only.**  
> The authoritative code lives inside the `.xlsm` binary; both should always
> be kept in sync.

---

## Module overview

| File | Module name | Purpose |
|------|-------------|---------|
| `Evaluation.bas` | `Evaluation` | **Main evaluation engine** – builds the *Evaluation Results* sheet (**see fixes below**) |
| `carselection.bas` | `carselection` | Car-selection popup dialog (InputBox-based) |
| `HeatMap.bas` | `HeatMap` | Refreshes the HeatMap Sheet from the Data Transfer Sheet |
| `Updatesuboperationstatus.bas` | `Updatesuboperationstatus` | Writes colored bullet dots back to the HeatMap Sheet |
| `OperationModeStatus.bas` | `OperationModeStatus` | Rolls up sub-operation dots into a group NOK/Acceptable/OK header |
| `Reset.bas` | `Reset` | Resets the HeatMap Sheet from the HeatMap Template |
| `Clearall.bas` | `Clearall` | Clears Sheet1 while preserving the Clear button |
| `Export.bas` | `Export` | Exports the visible selection as a clipboard picture |
| `Greendot.bas` | `Greendot` | Inserts a green dot (●) into the selected cell |
| `YellowDot.bas` | `YellowDot` | Inserts a yellow dot (●) into the selected cell |
| `RedDot.bas` | `RedDot` | Inserts a red dot (●) into the selected cell |
| `BlueDot.bas` | `BlueDot` | Inserts a blue dot (●) into the selected cell |

---

## Fixes applied to `Evaluation.bas`

Three bugs in `EvaluateAVLStatus` were corrected:

### Fix 1 – Operation name column (B → C)

The operation name text lives in **column C (3)**, not column B (2).

```vba
' Before (wrong – column B was always empty):
wsResults.Cells(outRow, 2).Value = wsSheet1.Cells(i, 2).Value

' After (fixed):
wsResults.Cells(outRow, 2).Value = wsSheet1.Cells(i, 3).Value
```

### Fix 2 – Drivability P1 column (E → F)

The Drivability P1 status indicator dots (●) are in **column F (6)**, not column E (5).

```vba
' Before (wrong – column E was always empty):
drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 5))

' After (fixed):
drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 6))
```

### Fix 3 – Section-header rows included in evaluation

Rows such as `"Accelerations"`, `"Drive away"`, `"Decelerations"` share column A
with real op-codes but are plain text, not 8-digit numbers.  The old guard
`If Trim(CStr(opCode)) <> "" Then` passed those rows through, producing bogus
results.  The fix checks `IsNumeric`:

```vba
' Before (wrong – included section headers):
If Trim(CStr(opCode)) <> "" Then

' After (fixed – only real numeric op-codes like 10101300):
If IsNumeric(opCode) = True Then
```

---

## How to re-import into Excel (if needed)

1. Open `AVLDrive_Heatmap_Tool version_5 (1).xlsm` in Excel.
2. Press **Alt + F11** to open the VBA editor.
3. In the Project Explorer, right-click the target module and choose
   **Remove module** (export a copy when prompted if you want a backup).
4. Right-click the project root and choose **Import File…**, then select
   the `.bas` file from this folder.
5. Save the workbook as `.xlsm`.
