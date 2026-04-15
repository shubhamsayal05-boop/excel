# How to Apply the Corrected VBA Code

## The Problem
When clicking the **HeatMap** button, all other buttons (Export, Reset, Evaluation, etc.) disappear.

## The Fix
Buttons are form controls anchored to cells. When rows/columns are hidden, buttons anchored to those cells get hidden too. The fix **detaches buttons from the cell grid BEFORE hiding**, then **re-shows them AFTER hiding**.

## Steps to Apply

1. Open your `.xlsm` file in Excel
2. Press **Alt + F11** to open the VBA Editor
3. In the Project Explorer (left panel), find and expand **Modules**
4. **Double-click** the `HeatMap` module
5. Press **Ctrl + A** to select all code, then **Delete**
6. Copy the **entire** contents of `HeatMap.bas` from this repository and paste it in
7. Repeat steps 4-6 for the `Evaluation` module using `Evaluation.bas`
8. Press **Ctrl + S** to save
9. Close the VBA Editor (Alt + Q)

## Files to Copy

| File | Module | What it fixes |
|------|--------|---------------|
| `HeatMap.bas` | HeatMap | Buttons disappearing when HeatMap runs |
| `Evaluation.bas` | Evaluation | Op code column, lastRow range, dynamic Tested Vehicle lookup |
