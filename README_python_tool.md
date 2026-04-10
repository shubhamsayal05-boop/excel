# AVL-DRIVE Heatmap Tool — Python Edition

A standalone Python replica of the Excel VBA-based AVL-DRIVE Heatmap Tool (version 5.1). This tool reads the original `.xlsm` workbook, performs all operations in-memory, and writes results back — completely independent of Excel.

## Features

All six functions from the Excel VBA tool are replicated 1:1:

| Button | Function | VBA Module |
|--------|----------|------------|
| **HeatMap** | Transfer data from Data Transfer Sheet → HeatMap Sheet | `HeatMap.bas` |
| **Reset** | Restore HeatMap Sheet from HeatMap Template | `Reset.bas` |
| **Evaluation** | Evaluate AVL statuses with interactive car selection | `Evaluation.bas` + `carselection.bas` |
| **Suboperation Status** | Write colored status dots to HeatMap Sheet column R | `Updatesuboperationstatus.bas` |
| **Operation Mode Status** | Aggregate group statuses (NOK/Acceptable/OK) | `OperationModeStatus.bas` |
| **Export** | Download any sheet's visible data as a standalone XLSX | `Export.bas` |

Additionally:
- **Clear Sheet1** — Clear all data in Sheet1 (`Clearall.bas`)
- **Sheet Preview** — View any sheet's data in the browser
- **Download** — Download the modified workbook with all changes applied

## Evaluation Logic (1:1 match with VBA)

The evaluation follows these exact rules from the VBA code:

### P1 Status Detection
- Font colour `#008000` (indexed 17) → **GREEN**
- Font colour `#FFFF00` (indexed 13) → **YELLOW**  
- Font colour `#FF0000` (indexed 10) → **RED**
- Font colour `#FFFFFF` (white/default) → **N/A**

### Status Evaluation Rules
1. P1 = N/A → **N/A**
2. AVL < 7 OR P1 = RED → **RED**
3. AVL ≥ 7 AND P1 = YELLOW → **YELLOW**
4. AVL ≥ 7 AND P1 = GREEN AND no benchmark → **GREEN**
5. AVL ≥ 7 AND P1 = GREEN AND tested ≥ target → **GREEN**
6. AVL ≥ 7 AND P1 = GREEN AND (target − tested) ≤ 2 → **GREEN**
7. AVL ≥ 7 AND P1 = GREEN AND (target − tested) > 2 → **YELLOW**

### Final Status Combination
- Either RED → **RED**
- Either YELLOW → **YELLOW**
- Both GREEN → **GREEN**
- One GREEN + one N/A → **GREEN**
- Both N/A → **N/A**

### Operation Mode Status Aggregation
- Any RED sub-operation → **NOK** (red)
- Yellow sub-operations > 35% → **Acceptable** (yellow)
- Otherwise → **OK** (green)

## Installation

```bash
pip install -r requirements.txt
```

Requirements:
- Python 3.9+
- openpyxl ≥ 3.1.0
- streamlit ≥ 1.30.0
- pandas ≥ 2.0.0

## Usage

### Streamlit GUI (recommended)

```bash
streamlit run avl_heatmap_tool.py
```

Then open http://localhost:8501 in your browser and upload the `.xlsm` file.

### Python API (programmatic use)

```python
import openpyxl
from avl_heatmap_tool import *

wb = openpyxl.load_workbook("AVLDrive_Heatmap_Tool version_5.1.xlsm", keep_vba=True)

# Refresh heatmap
msg = refresh_heatmap(wb)
print(msg)

# Run evaluation
msg = evaluate_avl_status(wb, "MY27_K0_X0_QG3_V131_0725", "MY27_K0_QG3_X0_V528_0226")
print(msg)

# Update sub-operation status
msg = update_sub_operation_heatmap(wb)
print(msg)

# Update operation mode status
msg = update_operation_mode_status(wb)
print(msg)

# Save
wb.save("output.xlsm")
wb.close()
```

## Sheet Structure

The tool works with these sheets in the workbook:

| Sheet | Purpose |
|-------|---------|
| **Sheet1** | Source data with op codes (col B), operation names (col C), P1 status dots (col F/L), benchmark values |
| **HeatMap Sheet** | Visual heatmap with vehicle AVL scores, status column (R) |
| **HeatMap Template** | Clean template for resetting HeatMap Sheet |
| **Data Transfer Sheet** | Source data for HeatMap refresh |
| **Mapping Sheet** | Op code → operation name mapping (58 entries) |
| **AVL-Odriv Mapping** | Extended op code mapping with sub-operation variants (99 entries) |
| **Evaluation Results** | Output sheet created by the Evaluation function |

## Architecture

The Python tool is a single file (`avl_heatmap_tool.py`) with:
- **Utility functions** — cell value reading, type conversion, colour resolution
- **Colour/status logic** — P1 detection, evaluation rules, status combination
- **Six operation modules** — each mirroring a VBA module exactly
- **Streamlit GUI** — interactive web interface with file upload/download
- **Data preview** — pandas DataFrame display of any sheet
