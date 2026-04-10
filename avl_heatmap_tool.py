#!/usr/bin/env python3
"""
AVL-DRIVE Heatmap Tool — Python Edition
=========================================
A 1:1 standalone Python replica of the Excel VBA-based AVL-DRIVE Heatmap Tool
(version 5.1). Reads the original `.xlsm` workbook, performs all operations
in-memory using openpyxl, and writes the results back.

Provides a Streamlit GUI with the same six buttons as the Excel ribbon:
  1. HeatMap   — Refresh heatmap from Data Transfer Sheet
  2. Reset     — Restore HeatMap Sheet from HeatMap Template
  3. Evaluation — Evaluate AVL statuses with car selection
  4. Suboperation Status — Write colored dots to HeatMap Sheet
  5. Operation Mode Status — Aggregate group statuses
  6. Export    — Export visible selection as image (PNG)

Usage:
    streamlit run avl_heatmap_tool.py
"""

from __future__ import annotations

import copy
import io
import os
import sys
import textwrap
from collections import OrderedDict
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ─── Constants (mirrors VBA Public Const) ────────────────────────────────────

SHEET_T = "HeatMap Sheet"
SHEET_S = "Data Transfer Sheet"
TEMPLATE_SHEET = "HeatMap Template"
MAPPING_SHEET = "Mapping Sheet"
AVL_ODRIV_MAPPING = "AVL-Odriv Mapping"
SHEET1 = "Sheet1"
EVAL_RESULTS = "Evaluation Results"

ANCHOR_TEXT = "Operation Modes"
TARGET_VEHICLE_HEADER = "Target Vehicle"
TESTED_VEHICLE_HEADER = "Tested Vehicle"

HIDE_IDS_COLA = True
DELETE_EMPTY = False

CAR_DATA_START_COL = 8  # Column H in Sheet1

# Vehicle columns in HeatMap Sheet (1-indexed)
VEHICLE_COLS = [4, 6, 8, 10, 12, 14, 16]      # D F H J L N P
SEPARATOR_COLS = [5, 7, 9, 11, 13, 15, 17]     # E G I K M O Q

# Excel indexed colour palette (standard)
_INDEXED_COLORS = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
    5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
    10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF",
    15: "00FFFF", 16: "800000", 17: "008000", 18: "000080", 19: "808000",
    20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080",
}

# Color → P1 status mapping
_GREEN_HEX = {"008000", "00B050", "009E47"}
_YELLOW_HEX = {"FFFF00", "FFC000", "FFD966", "E3E100"}
_RED_HEX = {"FF0000", "C00000"}
_WHITE_HEX = {"FFFFFF"}

# Fill for status cells
FILL_GREEN = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
FILL_RED = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
FILL_HEADER = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
FILL_SUMMARY = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

FONT_WHITE = Font(color="FFFFFF")
FONT_BLACK = Font(color="000000")
FONT_BOLD_WHITE = Font(color="FFFFFF", bold=True)

BULLET = "\u25CF"  # ●

# ═══════════════════════════════════════════════════════════════════════════════
#  UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _cell_val(ws, row: int, col: int) -> Any:
    """Return cell value (None-safe)."""
    return ws.cell(row, col).value


def _to_float(v: Any) -> float:
    """Convert to float; return 0.0 for non-numeric."""
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _is_numeric(v: Any) -> bool:
    """Return True if *v* can be interpreted as a number."""
    if v is None or v == "":
        return False
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False


def _trim(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _resolve_font_color_hex(cell) -> str:
    """Return 6-char uppercase hex of the font colour, or 'FFFFFF' (white)."""
    fc = cell.font.color
    if fc is None:
        return "FFFFFF"
    if fc.type == "rgb" and fc.rgb:
        h = str(fc.rgb)
        if len(h) == 8:
            h = h[2:]  # strip alpha
        return h.upper()
    if fc.type == "indexed" and fc.indexed is not None:
        idx = fc.indexed
        if idx in _INDEXED_COLORS:
            return _INDEXED_COLORS[idx].upper()
    if fc.type == "theme":
        # Theme colours need the workbook theme; approximate common ones
        return "FFFFFF"
    return "FFFFFF"


def _resolve_fill_color_hex(cell) -> str:
    """Return 6-char uppercase hex of the cell fill colour, or '000000'."""
    fill = cell.fill
    if fill is None or fill.fgColor is None:
        return "000000"
    fg = fill.fgColor
    if fg.type == "rgb" and fg.rgb:
        h = str(fg.rgb)
        if len(h) == 8:
            h = h[2:]
        return h.upper()
    if fg.type == "indexed" and fg.indexed is not None:
        idx = fg.indexed
        if idx in _INDEXED_COLORS:
            return _INDEXED_COLORS[idx].upper()
    return "000000"


def _is_near(r: int, g: int, b: int, rt: int, gt: int, bt: int, tol: int = 45) -> bool:
    return abs(r - rt) <= tol and abs(g - gt) <= tol and abs(b - bt) <= tol


def _hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    if len(h) == 8:
        h = h[2:]
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _find_anchor(ws, text: str) -> Optional[Tuple[int, int]]:
    """Find the row, col of the anchor text (case-insensitive, whole-cell)."""
    tl = text.strip().lower()
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            v = _cell_val(ws, row, col)
            if v is not None and _trim(v).lower() == tl:
                return (row, col)
    return None


def _has_value(v: Any) -> bool:
    """Mirror VBA HasValue: True if v is a positive number or non-empty non-zero string."""
    if v is None:
        return False
    try:
        if isinstance(v, (int, float)):
            return float(v) > 0
        fv = float(v)
        return fv > 0
    except (ValueError, TypeError):
        s = str(v).strip()
        return s != "" and s != "0"


# ═══════════════════════════════════════════════════════════════════════════════
#  COLOUR / STATUS LOGIC  (mirrors Evaluation.bas)
# ═══════════════════════════════════════════════════════════════════════════════

def get_p1_status_from_color(cell) -> str:
    """Determine P1 status from the cell's font colour (indexed or RGB).

    Mapping (from actual data analysis):
        indexed 17 / #008000  → GREEN
        indexed 13 / #FFFF00  → YELLOW
        indexed 10 / #FF0000  → RED
        #FFFFFF / white       → N/A
    """
    hex_color = _resolve_font_color_hex(cell)
    r, g, b = _hex_to_rgb(hex_color)

    # Check fill colour first (mirrors VBA MapColorToStatus)
    fill_hex = _resolve_fill_color_hex(cell)
    fr, fg_, fb = _hex_to_rgb(fill_hex)

    # Fill colour checks
    if _is_near(fr, fg_, fb, 0, 176, 80) or _is_near(fr, fg_, fb, 0, 158, 71):
        return "GREEN"
    if _is_near(fr, fg_, fb, 255, 192, 0) or _is_near(fr, fg_, fb, 255, 217, 102, 60):
        return "YELLOW"
    if _is_near(fr, fg_, fb, 255, 0, 0) or _is_near(fr, fg_, fb, 192, 0, 0):
        return "RED"

    # Font colour checks (broader tolerance for the most common)
    if _is_near(r, g, b, 0, 128, 0, 5):
        return "GREEN"
    if _is_near(r, g, b, 0, 176, 80) or _is_near(r, g, b, 0, 158, 71):
        return "GREEN"
    if _is_near(r, g, b, 255, 255, 0, 5):
        return "YELLOW"
    if _is_near(r, g, b, 255, 192, 0) or _is_near(r, g, b, 255, 217, 102, 60):
        return "YELLOW"
    if _is_near(r, g, b, 227, 225, 0, 30):
        return "YELLOW"
    if _is_near(r, g, b, 255, 0, 0) or _is_near(r, g, b, 192, 0, 0):
        return "RED"

    return "N/A"


def bench_diff(target_val: float, tested_val: float) -> float:
    """Return abs difference; 999 sentinel when target is zero."""
    if target_val == 0 and tested_val == 0:
        return 999.0
    if target_val == 0:
        return 999.0
    return abs(tested_val - target_val)


def evaluate_status(avl: float, p1: str, bdiff: float,
                    target_val: float, tested_val: float) -> str:
    """Mirror EvaluateStatus VBA function."""
    p1u = p1.upper().strip()

    if p1u == "N/A":
        return "N/A"
    if avl < 7 or p1u == "RED":
        return "RED"
    if avl >= 7 and p1u == "YELLOW":
        return "YELLOW"

    # AVL >= 7 and P1 == GREEN
    if bdiff == 999:
        return "GREEN"
    if tested_val >= target_val:
        return "GREEN"
    if (target_val - tested_val) <= 2:
        return "GREEN"
    return "YELLOW"


def combine_status(driv: str, resp: str) -> str:
    """Mirror CombineStatus VBA function."""
    d = driv.upper().strip() if driv else "N/A"
    r = resp.upper().strip() if resp else "N/A"
    if d == "":
        d = "N/A"
    if r == "":
        r = "N/A"

    if d == "RED" or r == "RED":
        return "RED"
    if d == "YELLOW" or r == "YELLOW":
        return "YELLOW"
    if d == "GREEN" and r == "GREEN":
        return "GREEN"
    if (d == "GREEN" and r == "N/A") or (d == "N/A" and r == "GREEN"):
        return "GREEN"
    return "N/A"


def color_cell(ws, row: int, col: int, status: str):
    """Apply fill + font colour to a results cell."""
    cell = ws.cell(row, col)
    s = status.upper().strip()
    if s == "GREEN":
        cell.fill = FILL_GREEN
        cell.font = FONT_WHITE
    elif s == "YELLOW":
        cell.fill = FILL_YELLOW
        cell.font = FONT_BLACK
    elif s == "RED":
        cell.fill = FILL_RED
        cell.font = FONT_WHITE
    else:
        cell.fill = PatternFill(fill_type=None)
        cell.font = FONT_BLACK


# ═══════════════════════════════════════════════════════════════════════════════
#  1.  HEATMAP REFRESH  (mirrors HeatMap.bas → RefreshHeatmap)
# ═══════════════════════════════════════════════════════════════════════════════

def _collect_dest_vehicle_cols(ws, anc_row: int, anc_col: int) -> List[int]:
    """Collect destination vehicle columns by DR marker in row anc_row+1."""
    last_c = ws.max_column
    out: List[int] = []
    for c in range(anc_col + 1, last_c + 1):
        v = _cell_val(ws, anc_row + 1, c)
        if isinstance(v, str) and v.strip().upper().startswith("DR"):
            out.append(c)
    if out:
        return out
    # Fallback: contiguous headers until COMMENTS
    for c in range(anc_col + 1, last_c + 1):
        v = _trim(_cell_val(ws, anc_row, c))
        if v.upper() == "COMMENTS":
            break
        if v:
            out.append(c)
        elif out:
            break
    return out


def _collect_headers(ws, anc_row: int, anc_col: int) -> List[str]:
    out: List[str] = []
    for c in range(anc_col + 1, ws.max_column + 1):
        v = _trim(_cell_val(ws, anc_row, c))
        if v:
            out.append(v)
    return out


def _collect_header_cols(ws, anc_row: int, anc_col: int) -> List[int]:
    out: List[int] = []
    for c in range(anc_col + 1, ws.max_column + 1):
        v = _trim(_cell_val(ws, anc_row, c))
        if v:
            out.append(c)
    return out


def _collect_row_labels(ws, anc_row: int, anc_col: int) -> List[str]:
    out: List[str] = []
    empty_run = 0
    last_r = ws.max_row
    for r in range(anc_row + 2, last_r + 1):
        v = _trim(_cell_val(ws, r, anc_col))
        if v:
            out.append(v)
            empty_run = 0
        else:
            empty_run += 1
            if empty_run >= 10:
                break
    return out


def _build_mode_index(ws, anc_row: int, anc_col: int) -> Dict[str, int]:
    d: Dict[str, int] = {}
    last_r = ws.max_row
    for r in range(anc_row + 2, last_r + 1):
        v = _trim(_cell_val(ws, r, anc_col))
        if v and v.lower() not in {k.lower() for k in d}:
            d[v] = r
    return d


def refresh_heatmap(wb) -> str:
    """Transfer data from Data Transfer Sheet to HeatMap Sheet."""
    ws_t = wb[SHEET_T]
    ws_s = wb[SHEET_S]

    t_anc = _find_anchor(ws_t, ANCHOR_TEXT)
    s_anc = _find_anchor(ws_s, ANCHOR_TEXT)
    if t_anc is None or s_anc is None:
        return f"Anchor text '{ANCHOR_TEXT}' not found in one or both sheets."

    t_row, t_col = t_anc
    s_row, s_col = s_anc

    t_veh_cols = _collect_dest_vehicle_cols(ws_t, t_row, t_col)
    t_modes = _collect_row_labels(ws_t, t_row, t_col)
    s_veh_hdr = _collect_headers(ws_s, s_row, s_col)
    s_veh_col = _collect_header_cols(ws_s, s_row, s_col)
    s_mode_ix = _build_mode_index(ws_s, s_row, s_col)

    if not t_veh_cols or not s_veh_col or not t_modes:
        return "No data found to process."

    n = min(len(t_veh_cols), len(s_veh_hdr))
    warnings: List[str] = []
    if len(s_veh_hdr) > len(t_veh_cols):
        warnings.append(
            f"Data Transfer Sheet has {len(s_veh_hdr)} vehicles but HeatMap "
            f"can accommodate {len(t_veh_cols)}. Only first {n} transferred."
        )

    # Vehicle header labels
    if n > 0:
        ws_t.cell(t_row - 1, t_veh_cols[0]).value = TARGET_VEHICLE_HEADER
        ws_t.cell(t_row - 1, t_veh_cols[0]).font = Font(name="Arial", size=16)
        ws_t.cell(t_row - 1, t_veh_cols[0]).alignment = Alignment(horizontal="center")

        ws_t.cell(t_row - 1, t_veh_cols[n - 1]).value = TESTED_VEHICLE_HEADER
        ws_t.cell(t_row - 1, t_veh_cols[n - 1]).font = Font(name="Arial", size=16)
        ws_t.cell(t_row - 1, t_veh_cols[n - 1]).alignment = Alignment(horizontal="center")

    # Vehicle names
    for i in range(n):
        ws_t.cell(t_row, t_veh_cols[i]).value = s_veh_hdr[i]

    # Clear old data
    last_r = t_row + 1 + len(t_modes)
    for j in range(n):
        for r in range(t_row + 2, last_r + 1):
            ws_t.cell(r, t_veh_cols[j]).value = None

    # Fill data
    filled = 0
    for i, mode in enumerate(t_modes):
        matched_key = None
        for k in s_mode_ix:
            if k.lower() == mode.lower():
                matched_key = k
                break
        if matched_key is not None:
            r_s = s_mode_ix[matched_key]
            r_t = t_row + 1 + (i + 1)
            for j in range(n):
                v = _cell_val(ws_s, r_s, s_veh_col[j])
                if _has_value(v):
                    try:
                        ws_t.cell(r_t, t_veh_cols[j]).value = float(v)
                        filled += 1
                    except (ValueError, TypeError):
                        pass

    # Hide rows missing last vehicle data
    if not DELETE_EMPTY:
        _hide_rows_missing_last_vehicle(ws_t, t_row, t_col, t_veh_cols[n - 1])
    else:
        _delete_rows_missing_last_vehicle(ws_t, t_row, t_col, t_veh_cols[n - 1])

    msg = f"HeatMap refreshed: {filled} values transferred for {n} vehicle(s)."
    if warnings:
        msg += "\n" + "\n".join(warnings)
    return msg


def _hide_rows_missing_last_vehicle(ws, anc_row: int, anc_col: int, last_veh_col: int):
    """Mark rows hidden (openpyxl row_dimensions)."""
    last_r = ws.max_row
    for r in range(anc_row + 2, last_r + 1):
        v = _trim(_cell_val(ws, r, anc_col))
        veh = _cell_val(ws, r, last_veh_col)
        if v == "" or not _has_value(veh):
            ws.row_dimensions[r].hidden = True
        else:
            ws.row_dimensions[r].hidden = False


def _delete_rows_missing_last_vehicle(ws, anc_row: int, anc_col: int, last_veh_col: int):
    last_r = ws.max_row
    for r in range(last_r, anc_row + 1, -1):
        v = _trim(_cell_val(ws, r, anc_col))
        veh = _cell_val(ws, r, last_veh_col)
        if v == "" or not _has_value(veh):
            ws.delete_rows(r)


# ═══════════════════════════════════════════════════════════════════════════════
#  2.  RESET  (mirrors Reset.bas → ResetTemplate_From_Sheet4)
# ═══════════════════════════════════════════════════════════════════════════════

def reset_heatmap(wb) -> str:
    """Copy HeatMap Template over HeatMap Sheet."""
    if TEMPLATE_SHEET not in wb.sheetnames:
        return f"Template sheet '{TEMPLATE_SHEET}' not found."
    if SHEET_T not in wb.sheetnames:
        return f"Destination sheet '{SHEET_T}' not found."

    ws_src = wb[TEMPLATE_SHEET]
    ws_dst = wb[SHEET_T]

    # Unmerge all merged cells in destination first
    for merge_range in list(ws_dst.merged_cells.ranges):
        ws_dst.unmerge_cells(str(merge_range))

    # Clear destination
    for row in ws_dst.iter_rows(min_row=1, max_row=ws_dst.max_row,
                                 min_col=1, max_col=ws_dst.max_column):
        for cell in row:
            cell.value = None

    # Copy values and basic formatting from template
    for row in ws_src.iter_rows(min_row=1, max_row=ws_src.max_row,
                                 min_col=1, max_col=ws_src.max_column):
        for cell in row:
            dst_cell = ws_dst.cell(cell.row, cell.column)
            dst_cell.value = cell.value
            if cell.has_style:
                dst_cell.font = copy.copy(cell.font)
                dst_cell.fill = copy.copy(cell.fill)
                dst_cell.alignment = copy.copy(cell.alignment)
                dst_cell.border = copy.copy(cell.border)
                dst_cell.number_format = cell.number_format

    # Copy column widths
    for col_letter, dim in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_num, dim in ws_src.row_dimensions.items():
        ws_dst.row_dimensions[row_num].height = dim.height

    # Copy merged cell ranges from template
    for merge_range in ws_src.merged_cells.ranges:
        ws_dst.merge_cells(str(merge_range))

    # Unhide all rows
    for r in range(1, ws_dst.max_row + 1):
        ws_dst.row_dimensions[r].hidden = False

    return "HeatMap Sheet has been reset from template."


# ═══════════════════════════════════════════════════════════════════════════════
#  3.  EVALUATION  (mirrors Evaluation.bas → EvaluateAVLStatus)
# ═══════════════════════════════════════════════════════════════════════════════

def get_available_car_names(ws) -> List[str]:
    """Scan row 2 of Sheet1 from col H onwards for unique car names."""
    names: List[str] = []
    last_col = ws.max_column
    skip_words = {"status", "p1", "p2", "p3", "lowest events"}
    for col in range(CAR_DATA_START_COL, last_col + 1):
        name = _trim(_cell_val(ws, 2, col))
        if not name:
            continue
        if any(w in name.lower() for w in skip_words):
            continue
        if name not in names:
            names.append(name)
    return names


def _find_car_column(ws, car_name: str, start_col: int = CAR_DATA_START_COL) -> int:
    """Find column for a car name in row 2 from start_col onwards."""
    for col in range(start_col, ws.max_column + 1):
        if _trim(_cell_val(ws, 2, col)) == car_name.strip():
            return col
    return 0


def _get_tested_avl(ws_heatmap, op_code, tested_car_name: str) -> float:
    """Look up Tested AVL score from HeatMap Sheet."""
    op_key = _trim(str(op_code))

    # Find the column under "Tested Vehicle" header in row 1
    avl_col = 0
    for col in range(1, ws_heatmap.max_column + 1):
        v = _trim(_cell_val(ws_heatmap, 1, col))
        if v and "tested vehicle" in v.lower():
            avl_col = col
            break

    # Fallback: search row 2 for tested car name
    if avl_col == 0:
        for col in range(1, ws_heatmap.max_column + 1):
            if _trim(_cell_val(ws_heatmap, 2, col)) == tested_car_name.strip():
                avl_col = col
                break

    # Default to column 8
    if avl_col == 0:
        avl_col = 8

    # Search column A for op code
    for r in range(1, ws_heatmap.max_row + 1):
        v = _cell_val(ws_heatmap, r, 1)
        if v is not None and _trim(str(v)) == op_key:
            return _to_float(_cell_val(ws_heatmap, r, avl_col))

    # Numeric match
    if _is_numeric(op_key):
        op_num = int(float(op_key))
        for r in range(1, ws_heatmap.max_row + 1):
            v = _cell_val(ws_heatmap, r, 1)
            if v is not None:
                try:
                    if int(float(v)) == op_num:
                        return _to_float(_cell_val(ws_heatmap, r, avl_col))
                except (ValueError, TypeError):
                    pass

    return 0.0


def evaluate_avl_status(wb, target_car: str, tested_car: str) -> str:
    """Run the full evaluation and write to Evaluation Results sheet."""
    ws1 = wb[SHEET1]
    ws_hm = wb[SHEET_T]

    # Find columns in Drivability section (starts around col 8)
    target_col = _find_car_column(ws1, target_car)
    tested_col = _find_car_column(ws1, tested_car)

    if target_col == 0 or tested_col == 0:
        return f"Could not find data columns for selected cars.\nTarget: {target_car} (col {target_col})\nTested: {tested_car} (col {tested_col})"

    # Find columns in Responsiveness section (starts at col 12)
    target_resp_col = _find_car_column(ws1, target_car, 12)
    tested_resp_col = _find_car_column(ws1, tested_car, 12)

    if target_resp_col == 0 or tested_resp_col == 0:
        return f"Could not find responsiveness columns for selected cars."

    # Delete existing results sheet if present
    if EVAL_RESULTS in wb.sheetnames:
        del wb[EVAL_RESULTS]

    # Create new results sheet
    ws_r = wb.create_sheet(EVAL_RESULTS)

    # Header row
    headers = [
        "Op Code", "Operation", "Tested AVL",
        f"Driv P1", f"Driv Target ({target_car})", f"Driv Tested ({tested_car})", "Driv Status",
        f"Resp P1", f"Resp Target ({target_car})", f"Resp Tested ({tested_car})", "Resp Status",
        "Final Status"
    ]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_r.cell(1, col_idx, h)
        cell.font = FONT_BOLD_WHITE
        cell.fill = FILL_HEADER

    # Compute lastRow from max of columns A, B, C
    last_row = max(
        _last_data_row(ws1, 1),
        _last_data_row(ws1, 2),
        _last_data_row(ws1, 3),
    )

    out_row = 2
    for i in range(5, last_row + 1):
        op_code = _cell_val(ws1, i, 2)  # Column B

        # Skip empty / non-numeric (section headers)
        if op_code is None or not _is_numeric(op_code):
            continue

        tested_avl = _get_tested_avl(ws_hm, op_code, tested_car)

        driv_p1 = get_p1_status_from_color(ws1.cell(i, 6))   # Column F
        resp_p1 = get_p1_status_from_color(ws1.cell(i, 12))  # Column L

        driv_target = _to_float(_cell_val(ws1, i, target_col))
        driv_tested = _to_float(_cell_val(ws1, i, tested_col))
        resp_target = _to_float(_cell_val(ws1, i, target_resp_col))
        resp_tested = _to_float(_cell_val(ws1, i, tested_resp_col))

        driv_bdiff = bench_diff(driv_target, driv_tested)
        resp_bdiff = bench_diff(resp_target, resp_tested)

        driv_status = evaluate_status(tested_avl, driv_p1, driv_bdiff, driv_target, driv_tested)
        resp_status = evaluate_status(tested_avl, resp_p1, resp_bdiff, resp_target, resp_tested)
        final_status = combine_status(driv_status, resp_status)

        operation_name = _cell_val(ws1, i, 3)  # Column C

        ws_r.cell(out_row, 1).value = op_code
        ws_r.cell(out_row, 2).value = operation_name
        ws_r.cell(out_row, 3).value = tested_avl
        ws_r.cell(out_row, 4).value = driv_p1
        ws_r.cell(out_row, 5).value = driv_target
        ws_r.cell(out_row, 6).value = driv_tested
        ws_r.cell(out_row, 7).value = driv_status
        ws_r.cell(out_row, 8).value = resp_p1
        ws_r.cell(out_row, 9).value = resp_target
        ws_r.cell(out_row, 10).value = resp_tested
        ws_r.cell(out_row, 11).value = resp_status
        ws_r.cell(out_row, 12).value = final_status

        color_cell(ws_r, out_row, 7, driv_status)
        color_cell(ws_r, out_row, 11, resp_status)
        color_cell(ws_r, out_row, 12, final_status)

        out_row += 1

    # Auto-fit columns (approximate)
    for col_idx in range(1, 13):
        ws_r.column_dimensions[get_column_letter(col_idx)].width = 20

    # Build summary
    _build_overall_status(ws_r)

    return (
        f"Evaluation complete!\n"
        f"Target: {target_car}\n"
        f"Tested: {tested_car}\n"
        f"Results: {out_row - 2} rows written to '{EVAL_RESULTS}'."
    )


def _last_data_row(ws, col: int) -> int:
    """Find the last non-empty row in a column."""
    for r in range(ws.max_row, 0, -1):
        if _cell_val(ws, r, col) is not None:
            return r
    return 1


def _build_overall_status(ws_r):
    """Build 'Overall Status by Op Code' summary table at bottom of results."""
    last_row = _last_data_row(ws_r, 1)

    codes: Dict[str, dict] = OrderedDict()
    for i in range(2, last_row + 1):
        code = _trim(_cell_val(ws_r, i, 1))
        if not code:
            continue
        status = _trim(_cell_val(ws_r, i, 12))
        name = _trim(_cell_val(ws_r, i, 2))

        if code not in codes:
            codes[code] = {"name": name, "statuses": [status]}
        else:
            codes[code]["statuses"].append(status)

    start_row = last_row + 2

    # Summary header
    ws_r.cell(start_row, 1).value = "Overall Status by Op Code"
    ws_r.cell(start_row, 1).font = Font(bold=True)
    ws_r.cell(start_row, 1).fill = FILL_SUMMARY
    ws_r.merge_cells(
        start_row=start_row, start_column=1,
        end_row=start_row, end_column=4
    )

    ws_r.cell(start_row + 1, 1).value = "Op Code"
    ws_r.cell(start_row + 1, 2).value = "Operation"
    ws_r.cell(start_row + 1, 3).value = "Overall Status"
    for c in range(1, 4):
        ws_r.cell(start_row + 1, c).font = Font(bold=True)

    r = start_row + 2
    for code, info in codes.items():
        any_red = False
        all_green = True
        has_valid = False

        for s in info["statuses"]:
            su = s.upper().strip()
            if su and su != "N/A":
                has_valid = True
                if su == "RED":
                    any_red = True
                if su != "GREEN":
                    all_green = False

        if not has_valid:
            overall = "N/A"
        elif any_red:
            overall = "RED"
        elif all_green:
            overall = "GREEN"
        else:
            overall = "YELLOW"

        ws_r.cell(r, 1).value = code
        ws_r.cell(r, 2).value = info["name"]
        ws_r.cell(r, 3).value = overall
        color_cell(ws_r, r, 3, overall)
        r += 1


# ═══════════════════════════════════════════════════════════════════════════════
#  4.  SUB-OPERATION STATUS  (mirrors Updatesuboperationstatus.bas)
# ═══════════════════════════════════════════════════════════════════════════════

def update_sub_operation_heatmap(wb) -> str:
    """Read Evaluation Results (A:C), write colored dots to HeatMap Sheet col R."""
    if EVAL_RESULTS not in wb.sheetnames:
        return "Evaluation Results sheet not found. Run Evaluation first."

    ws_eval = wb[EVAL_RESULTS]
    ws_heat = wb[SHEET_T]

    # Build dict: op_code → status from Evaluation Results
    d: Dict[str, str] = {}
    last_eval = _last_data_row(ws_eval, 1)
    for i in range(2, last_eval + 1):
        code = _trim(_cell_val(ws_eval, i, 1))
        if code:
            status = _trim(_cell_val(ws_eval, i, 3)).upper()
            if not status:
                status = _trim(_cell_val(ws_eval, i, 12)).upper()
            d[code] = status

    # Write to HeatMap Sheet column R (18) — non-bold rows only
    last_heat = _last_data_row(ws_heat, 1)
    updated = 0
    for i in range(2, last_heat + 1):
        cell_b = ws_heat.cell(i, 2)
        if cell_b.font and cell_b.font.bold:
            continue

        op_code = _trim(_cell_val(ws_heat, i, 1))
        if not op_code:
            continue

        # Clear existing
        target_cell = ws_heat.cell(i, 18)  # Column R
        target_cell.value = None

        if op_code in d:
            status = d[op_code]
            if status in ("RED", "YELLOW", "GREEN"):
                target_cell.value = BULLET
                target_cell.font = Font(size=14)
                target_cell.alignment = Alignment(horizontal="center")

                if status == "RED":
                    target_cell.font = Font(size=14, color="FF0000")
                elif status == "YELLOW":
                    target_cell.font = Font(size=14, color="E3E100")
                elif status == "GREEN":
                    target_cell.font = Font(size=14, color="00B050")
                updated += 1

    return f"Sub-operation HeatMap updated: {updated} status dots written."


# ═══════════════════════════════════════════════════════════════════════════════
#  5.  OPERATION MODE STATUS  (mirrors OperationModeStatus.bas)
# ═══════════════════════════════════════════════════════════════════════════════

def update_operation_mode_status(wb, status_col: int = 18) -> str:
    """Aggregate sub-operation statuses into group header rows.

    Groups are defined by bold rows in column B of HeatMap Sheet.
    The status column defaults to 18 (R) — same as sub-operation dots.
    """
    ws = wb[SHEET_T]
    last_row = _last_data_row(ws, 2)

    grp_start = 0
    groups_updated = 0

    for i in range(2, last_row + 2):
        cell_b = ws.cell(i, 2) if i <= last_row else None
        is_bold = cell_b is not None and cell_b.font and cell_b.font.bold

        if i > last_row or is_bold:
            if grp_start != 0 and grp_start - 1 > 2:
                _evaluate_group_status(ws, grp_start, i - 1, status_col)
                groups_updated += 1
            if i <= last_row:
                grp_start = i + 1
        # else: accumulate into current group

    return f"Operation mode statuses updated: {groups_updated} groups evaluated."


def _evaluate_group_status(ws, start_row: int, end_row: int, status_col: int):
    """Evaluate a single group's status and write to header cell."""
    if start_row - 1 <= 2:
        return

    red_cnt = 0
    yellow_cnt = 0
    total_cnt = 0

    for r in range(start_row, end_row + 1):
        cell = ws.cell(r, status_col)
        if cell.value and str(cell.value).strip():
            total_cnt += 1
            fc = _resolve_font_color_hex(cell)
            r_, g_, b_ = _hex_to_rgb(fc)
            if _is_near(r_, g_, b_, 255, 0, 0, 45):
                red_cnt += 1
            elif _is_near(r_, g_, b_, 227, 225, 0, 45) or _is_near(r_, g_, b_, 255, 255, 0, 45):
                yellow_cnt += 1
            # green: no counting needed

    if total_cnt == 0:
        return

    header_cell = ws.cell(start_row - 1, status_col)
    header_cell.font = Font(bold=True, color="000000")
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    if red_cnt > 0:
        header_cell.value = "NOK"
        header_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    elif total_cnt > 0 and yellow_cnt / total_cnt > 0.35:
        header_cell.value = "Acceptable"
        header_cell.fill = PatternFill(start_color="E3E100", end_color="E3E100", fill_type="solid")
    else:
        header_cell.value = "OK"
        header_cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")


# ═══════════════════════════════════════════════════════════════════════════════
#  6.  CLEAR ALL  (mirrors Clearall.bas)
# ═══════════════════════════════════════════════════════════════════════════════

def clear_sheet1(wb) -> str:
    """Clear all cells in Sheet1."""
    if SHEET1 not in wb.sheetnames:
        return f"Sheet '{SHEET1}' not found."
    ws = wb[SHEET1]
    # Unmerge all merged cells first
    for merge_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge_range))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.value = None
    return "Sheet1 cleared."


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT  (mirrors Export.bas — generates a formatted XLSX extract)
# ═══════════════════════════════════════════════════════════════════════════════

def export_sheet_data(wb, sheet_name: str) -> bytes:
    """Export a sheet's visible data to a new XLSX file (in memory).

    Since Python cannot copy-as-picture like VBA, we export to a clean
    XLSX that preserves all values, formatting, and column widths but
    omits hidden rows/columns. The caller can save or download this.
    """
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")

    ws = wb[sheet_name]
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = sheet_name

    out_r = 1
    for r in range(1, ws.max_row + 1):
        if ws.row_dimensions[r].hidden:
            continue
        out_c = 1
        for c in range(1, ws.max_column + 1):
            col_letter = get_column_letter(c)
            if ws.column_dimensions[col_letter].hidden:
                continue
            src = ws.cell(r, c)
            dst = new_ws.cell(out_r, out_c)
            dst.value = src.value
            if src.has_style:
                dst.font = copy.copy(src.font)
                dst.fill = copy.copy(src.fill)
                dst.alignment = copy.copy(src.alignment)
                dst.border = copy.copy(src.border)
                dst.number_format = src.number_format
            out_c += 1
        out_r += 1

    buf = io.BytesIO()
    new_wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA PREVIEW HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def sheet_to_dataframe(ws, max_rows: int = 500) -> "pd.DataFrame":
    """Convert a worksheet to a pandas DataFrame for display."""
    import pandas as pd

    data: List[List[Any]] = []
    for r in range(1, min(ws.max_row + 1, max_rows + 1)):
        row_data: List[Any] = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            row_data.append(v)
        data.append(row_data)

    if not data:
        return pd.DataFrame()

    # Use first row as header
    headers = [str(v) if v else f"Col{i+1}" for i, v in enumerate(data[0])]
    df = pd.DataFrame(data[1:], columns=headers)
    return df


# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT GUI
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="AVL-DRIVE Heatmap Tool",
        page_icon="🔧",
        layout="wide",
    )

    st.title("🔧 AVL-DRIVE Heatmap Tool — Python Edition")
    st.caption("Standalone Python replica of the Excel VBA tool (version 5.1)")

    # ── Sidebar: File Upload ─────────────────────────────────────────────
    st.sidebar.header("📂 Workbook")

    uploaded = st.sidebar.file_uploader(
        "Upload .xlsm workbook",
        type=["xlsm", "xlsx"],
        help="Upload the AVLDrive_Heatmap_Tool .xlsm file",
    )

    if uploaded is not None:
        if "wb_bytes" not in st.session_state or st.session_state.get("wb_name") != uploaded.name:
            st.session_state["wb_bytes"] = uploaded.read()
            st.session_state["wb_name"] = uploaded.name
            st.session_state["messages"] = []

    if "wb_bytes" not in st.session_state:
        st.info("👆 Upload an `.xlsm` workbook to get started.")
        st.markdown("""
        ### Features
        This tool replicates all six functions of the Excel VBA macro tool:

        | Button | Function |
        |--------|----------|
        | **HeatMap** | Transfer data from Data Transfer Sheet → HeatMap Sheet |
        | **Reset** | Restore HeatMap Sheet from HeatMap Template |
        | **Evaluation** | Evaluate AVL statuses with car selection |
        | **Suboperation Status** | Write colored status dots to HeatMap |
        | **Operation Mode Status** | Aggregate group statuses |
        | **Export** | Download a sheet's visible data as XLSX |
        """)
        return

    st.sidebar.success(f"✅ Loaded: `{st.session_state['wb_name']}`")

    # Load workbook
    wb = openpyxl.load_workbook(
        io.BytesIO(st.session_state["wb_bytes"]),
        keep_vba=True,
        data_only=False,
    )
    st.sidebar.write(f"Sheets: {', '.join(wb.sheetnames)}")

    # ── Messages Area ────────────────────────────────────────────────────
    if "messages" not in st.session_state:
        st.session_state["messages"] = []

    # ── Action Buttons (six columns) ─────────────────────────────────────
    st.subheader("⚡ Actions")

    col1, col2, col3, col4, col5, col6 = st.columns(6)

    with col1:
        if st.button("🗺️ HeatMap", use_container_width=True, help="Refresh heatmap from Data Transfer Sheet"):
            msg = refresh_heatmap(wb)
            _save_wb(wb)
            st.session_state["messages"].append(("info", msg))
            st.rerun()

    with col2:
        if st.button("🔄 Reset", use_container_width=True, help="Reset HeatMap from template"):
            msg = reset_heatmap(wb)
            _save_wb(wb)
            st.session_state["messages"].append(("info", msg))
            st.rerun()

    with col3:
        if st.button("📊 Evaluation", use_container_width=True, help="Evaluate AVL statuses"):
            st.session_state["show_eval_dialog"] = True

    with col4:
        if st.button("🔵 Subop Status", use_container_width=True, help="Update sub-operation status dots"):
            msg = update_sub_operation_heatmap(wb)
            _save_wb(wb)
            st.session_state["messages"].append(("info", msg))
            st.rerun()

    with col5:
        if st.button("📈 Op Mode Status", use_container_width=True, help="Update operation mode statuses"):
            msg = update_operation_mode_status(wb)
            _save_wb(wb)
            st.session_state["messages"].append(("info", msg))
            st.rerun()

    with col6:
        if st.button("🗑️ Clear Sheet1", use_container_width=True, help="Clear all data in Sheet1"):
            msg = clear_sheet1(wb)
            _save_wb(wb)
            st.session_state["messages"].append(("info", msg))
            st.rerun()

    # ── Evaluation Dialog ────────────────────────────────────────────────
    if st.session_state.get("show_eval_dialog"):
        st.divider()
        st.subheader("🚗 Car Selection for Evaluation")

        ws1 = wb[SHEET1]
        car_names = get_available_car_names(ws1)

        if not car_names:
            st.error("No car names found in Sheet1 row 2 (starting from column H).")
        else:
            ecol1, ecol2 = st.columns(2)
            with ecol1:
                target_car = st.selectbox("🎯 Target Car", car_names, key="eval_target")
            with ecol2:
                tested_idx = min(1, len(car_names) - 1) if len(car_names) > 1 else 0
                tested_car = st.selectbox("🔬 Tested Car", car_names, index=tested_idx, key="eval_tested")

            if target_car == tested_car:
                st.warning("⚠️ Same car selected for both Target and Tested.")

            bcol1, bcol2 = st.columns(2)
            with bcol1:
                if st.button("✅ Run Evaluation", type="primary"):
                    msg = evaluate_avl_status(wb, target_car, tested_car)
                    _save_wb(wb)
                    st.session_state["messages"].append(("success", msg))
                    st.session_state["show_eval_dialog"] = False
                    st.rerun()
            with bcol2:
                if st.button("❌ Cancel"):
                    st.session_state["show_eval_dialog"] = False
                    st.rerun()

    # ── Show messages ────────────────────────────────────────────────────
    for msg_type, msg_text in st.session_state.get("messages", []):
        if msg_type == "success":
            st.success(msg_text)
        elif msg_type == "error":
            st.error(msg_text)
        else:
            st.info(msg_text)

    # Clear messages after display
    if st.session_state.get("messages"):
        st.session_state["messages"] = []

    # ── Download modified workbook ───────────────────────────────────────
    st.divider()
    st.subheader("💾 Download")
    dcol1, dcol2 = st.columns(2)

    with dcol1:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            "📥 Download Modified Workbook",
            data=buf.getvalue(),
            file_name=st.session_state.get("wb_name", "output.xlsm"),
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            use_container_width=True,
        )

    with dcol2:
        export_sheet = st.selectbox("Export sheet", wb.sheetnames, key="export_sheet")
        if st.button("📤 Export Sheet as XLSX", use_container_width=True):
            try:
                data = export_sheet_data(wb, export_sheet)
                st.download_button(
                    f"⬇️ Download {export_sheet}.xlsx",
                    data=data,
                    file_name=f"{export_sheet}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Export error: {e}")

    # ── Sheet Preview ────────────────────────────────────────────────────
    st.divider()
    st.subheader("📋 Sheet Preview")

    import pandas as pd
    preview_sheet = st.selectbox("Select sheet to preview", wb.sheetnames, key="preview_sheet")
    ws_preview = wb[preview_sheet]

    df = sheet_to_dataframe(ws_preview, max_rows=200)
    if not df.empty:
        st.dataframe(df, use_container_width=True, height=400)
    else:
        st.write("Sheet is empty.")

    wb.close()


def _save_wb(wb):
    """Persist modified workbook back to session state."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.session_state["wb_bytes"] = buf.read()


if __name__ == "__main__":
    main()
