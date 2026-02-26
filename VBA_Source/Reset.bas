Attribute VB_Name = "Reset"
Option Explicit

Sub ResetTemplate_From_Sheet4()
    Const SRC_NAME As String = "HeatMap Template"   'backup template
    Const DST_NAME As String = "HeatMap Sheet"      'live template

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet, wsDst As Worksheet, wsNew As Worksheet
    Dim srcVis As XlSheetVisibility
    Dim prevScr As Boolean, prevEv As Boolean, prevCalc As XlCalculation, prevAlerts As Boolean

    On Error GoTo Fail
    Set wsSrc = wb.Worksheets(SRC_NAME)
    Set wsDst = wb.Worksheets(DST_NAME)

    'preserve app state
    prevScr = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    prevCalc = Application.Calculation
    prevAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    'remember source visibility and unhide temporarily
    srcVis = wsSrc.Visible
    If wsSrc.Visible <> xlSheetVisible Then wsSrc.Visible = xlSheetVisible

    'copy source before destination
    wsSrc.Copy Before:=wsDst
    Set wsNew = wsDst.Previous
    wsNew.Visible = xlSheetVisible
    wsNew.Activate

    'delete old Sheet1 and rename the copy
    wsDst.Delete
    wsNew.Name = DST_NAME

    'restore source visibility
    If wsSrc.Visible <> srcVis Then wsSrc.Visible = srcVis

    'ensure Sheet1 is first and visible
    wsNew.Move Before:=wb.Worksheets(1)
    wsNew.Visible = xlSheetVisible
    wsNew.Select

Done:
    Application.DisplayAlerts = prevAlerts
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevScr
    Exit Sub

Fail:
    'best-effort cleanup
    Resume Done
End Sub
