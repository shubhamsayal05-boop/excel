# AVL-DRIVE Heatmap Tool — Python Edition

A **fully independent** Python replica of the Excel VBA-based AVL-DRIVE Heatmap Tool (version 5.1). This tool does **not** require the original `.xlsm` file — all reference data (Mapping Sheet, AVL-Odriv Mapping, HeatMap Template) is embedded directly in the code.

## Quick Start

```bash
pip install -r requirements.txt
streamlit run avl_heatmap_tool.py
```

## Input Modes

### Standalone Mode (Recommended)
Upload your two data files directly — no `.xlsm` needed:

1. **Data Transfer Sheet** (`.xlsx` or `.csv`) — AVL-DRIVE scores per vehicle
2. **Sheet1** (`.xlsx` only) — Benchmark data with coloured P1 status dots

### Legacy Mode
Upload the original `AVLDrive_Heatmap_Tool version_5.1.xlsm` file.

## Data Formats

### Data Transfer Sheet
Users paste AVL-DRIVE scores into this sheet. The format is fixed but car names and values change:

| Col A | Col B | Col D | Col F | Col H | Col J |
|-------|-------|-------|-------|-------|-------|
| | Operation Modes | *Car 1* | *Car 2* | *Car 3* | *Car 4* |
| | Operation Modes | DR | DR | DR | DR |
| 10000000 | AVL-DRIVE Rating | 7.6 | 7.4 | 7.7 | 7.6 |
| 10100000 | Drive away | 7.7 | 7.8 | 7.6 | 7.8 |
| 10101300 | Creep | 7.9 | 7.5 | 7.3 | 7.9 |
| ... | ... | ... | ... | ... | ... |

- **Col A**: Operation code (numeric, e.g. 10101300)
- **Col B**: Operation name
- **Cols D, F, H, J…**: Vehicle scores (even columns, odd columns are empty separators)
- **Row 1**: Car/vehicle names
- **Row 2**: "DR" markers

### Sheet1 (Benchmark Data)
Must be `.xlsx` to preserve P1 status colours. Users paste benchmark data here:

| Col A | Col B | Col C | Col F | Col I | Col J | Col L | Col O | Col P |
|-------|-------|-------|-------|-------|-------|-------|-------|-------|
| | | | Drivability | | | Responsiveness | | |
| | | | Current Status | *Car 1* | *Car 2* | Current Status | *Car 1* | *Car 2* |
| USE CASE | | | P1 | | | P1 | | |
| Drive away | | | | 98.8 | | | 70.9 | |
| | 10101300 | Creep Eng On | ● | 100 | 73.7 | ● | 100 | 100 |

- **Col B**: Op code (numeric)
- **Col C**: Sub-operation name
- **Col F**: Drivability P1 status (coloured ● dot — green=OK, yellow=warning, red=fail)
- **Col L**: Responsiveness P1 status (coloured ● dot)
- **Car columns (I/J, O/P)**: Benchmark percentage values
- **Row 2**: Car names appear here

## Operations

| # | Button | Description |
|---|--------|-------------|
| 1 | **HeatMap** | Transfer data from Data Transfer Sheet → HeatMap Sheet |
| 2 | **Reset** | Restore HeatMap Sheet from built-in template |
| 3 | **Evaluation** | Evaluate AVL statuses (select Target + Tested cars) |
| 4 | **Suboperation Status** | Write coloured status dots to HeatMap Sheet |
| 5 | **Operation Mode Status** | Aggregate NOK/Acceptable/OK per operation group |
| 6 | **Clear Sheet1** | Clear all data from Sheet1 |

## Architecture

The tool embeds all reference data as Python constants:
- **MAPPING_SHEET_DATA** — 58 operation code → name mappings
- **AVL_ODRIV_MAPPING_DATA** — 99 op code → sub-operation name mappings
- **HEATMAP_TEMPLATE_ROWS** — 58 template rows with bold/group-header flags

When users upload their data files, the tool:
1. Copies Sheet1 (preserving font colours for P1 status detection)
2. Copies Data Transfer Sheet
3. Builds HeatMap Template, HeatMap Sheet, Mapping Sheet, and AVL-Odriv Mapping from embedded data
4. Assembles a complete in-memory workbook ready for all operations

## Dependencies

- `openpyxl` — Excel file reading/writing with formatting
- `streamlit` — Web GUI
- `pandas` — Data preview and CSV handling
