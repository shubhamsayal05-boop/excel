'================= MODULE:  Heatmap_Fill_ZeroIsBlank =================
Option Explicit

'================= CONFIG =================
Public Const SHEET_T As String = "HeatMap Sheet"
Public Const SHEET_S As String = "Data Transfer Sheet"
Public Const TEMPLATE_SHEET As String = "HeatMap Template"
Public Const ANCHOR_TEXT As String = "Operation Modes"

Public Const HIDE_IDS_COLA As Boolean = True
Public Const DELETE_EMPTY As Boolean = False   ' False = Hide rows, True = Delete rows

Public Const TARGET_VEHICLE_HEADER As String = "Target Vehicle"
Public Const TESTED_VEHICLE_HEADER As String = "Tested Vehicle"

'*** ADD PASSWORD CONSTANT - Change if your sheet has a password ***
Public Const SHEET_PASSWORD As String = ""     ' Leave blank if no password

'================= ENTRY =================
Public Sub RefreshHeatmap()

    Dim wsT As Worksheet, wsS As Worksheet, wsTemplate As Worksheet
    Dim tA As Range, sA As Range
    Dim tVehCols As Collection, tModes As Collection
    Dim sVehHdr As Collection, sVehCol As Collection
    Dim sModeIx As Object
    Dim n As Long, i As Long, j As Long
    Dim rt As Long, rS As Long, v
    Dim lastR As Long
    Dim wasProtected As Boolean

    On Error GoTo ErrorHandler

    Set wsT = ThisWorkbook.Worksheets(SHEET_T)
    Set wsS = ThisWorkbook.Worksheets(SHEET_S)
    Set wsTemplate = ThisWorkbook.Worksheets(TEMPLATE_SHEET)

    '*** UNPROTECT SHEET IF PROTECTED ***
    wasProtected = wsT.ProtectContents
    If wasProtected Then
        On Error Resume Next
        If SHEET_PASSWORD <> "" Then
            wsT.Unprotect Password:=SHEET_PASSWORD
        Else
            wsT.Unprotect
        End If
        If Err.Number <> 0 Then
            MsgBox "Cannot unprotect HeatMap Sheet. Password may be incorrect." & vbCrLf & _
                   "Error: " & Err.Description, vbCritical, "Protection Error"
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    End If

    Set tA = FindAnchor(wsT, ANCHOR_TEXT)
    Set sA = FindAnchor(wsS, ANCHOR_TEXT)
    If tA Is Nothing Or sA Is Nothing Then
        MsgBox "Anchor text '" & ANCHOR_TEXT & "' not found in one or both sheets.", vbExclamation
        GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- Discover structure ---
    Set tVehCols = CollectDestVehicleCols(wsT, tA)
    Set tModes = CollectRowLabels(wsT, tA)
    Set sVehHdr = CollectHeaders(wsS, sA)
    Set sVehCol = CollectHeaderCols(wsS, sA)
    Set sModeIx = BuildModeIndex(wsS, sA)

    If tVehCols.count = 0 Or sVehCol.count = 0 Or tModes.count = 0 Then
        MsgBox "No data found to process.", vbInformation
        GoTo CleanExit
    End If

    n = WorksheetFunction.Min(tVehCols.count, sVehHdr.count)
    
    '--- Warn if source data exceeds destination capacity ---
    If sVehHdr.count > tVehCols.count Then
        MsgBox "Warning: Data Transfer Sheet has " & sVehHdr.count & " vehicles, but HeatMap Sheet can only accommodate " & tVehCols.count & " vehicles." & vbCrLf & _
               "Only the first " & n & " vehicles will be transferred.", vbExclamation, "Data Capacity Warning"
    End If

    '--- Vehicle header rows ---
    If n > 0 Then
        wsT.Cells(tA.row - 1, tVehCols(1)).Value = TARGET_VEHICLE_HEADER
        ApplyVehicleHeaderFormatting wsT.Cells(tA.row - 1, tVehCols(1))

        wsT.Cells(tA.row - 1, tVehCols(n)).Value = TESTED_VEHICLE_HEADER
        ApplyVehicleHeaderFormatting wsT.Cells(tA.row - 1, tVehCols(n))
    End If

    For i = 1 To n
        wsT.Cells(tA.row, tVehCols(i)).Value = sVehHdr(i)
    Next i

    AutoAdjustVehicleColumns wsT, tA, tVehCols, n

    '--- Clear old data ---
    lastR = tA.row + 1 + tModes.count
    For j = 1 To n
        wsT.Range(wsT.Cells(tA.row + 2, tVehCols(j)), _
                  wsT.Cells(lastR, tVehCols(j))).ClearContents
    Next j

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

    '================= CORE RULE =================
    ' Remove / Hide entire operation if LAST vehicle has no data
    If DELETE_EMPTY Then
        DeleteRowsMissingLastVehicle wsT, tA, tVehCols(n)
    Else
        HideRowsMissingLastVehicle wsT, tA, tVehCols(n)
    End If

    '--- Hide ID column ---
    If HIDE_IDS_COLA Then
        wsT.Range(wsT.Cells(tA.row + 2, tA.Column - 1), _
                  wsT.Cells(wsT.Rows.count, tA.Column - 1)).NumberFormat = ";;;"
    End If

    '--- Restore separator formatting ---
    RestoreSeparatorColumnsFromTemplate wsT, wsTemplate, tA

    '--- Hide unused vehicle blocks ---
    ManageVehicleColumnVisibility wsT, n

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    '*** RE-PROTECT SHEET IF IT WAS PROTECTED ***
    If wasProtected Then
        On Error Resume Next
        If SHEET_PASSWORD <> "" Then
            wsT.Protect Password:=SHEET_PASSWORD, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True
        Else
            wsT.Protect DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True
        End If
        On Error GoTo 0
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & " in RefreshHeatmap: " & vbCrLf & Err.Description, _
           vbCritical, "Heatmap Error"
    Resume CleanExit

End Sub

'================= LAST VEHICLE MANDATORY =================
Public Sub HideRowsMissingLastVehicle(ws As Worksheet, anc As Range, lastVehCol As Long)
    Dim r As Long, lastR As Long
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0

    For r = anc.row + 2 To lastR
        If Trim$(ws.Cells(r, anc.Column).Value) = "" _
           Or Not HasValue(ws.Cells(r, lastVehCol).Value) Then
            ws.Rows(r).Hidden = True
        Else
            ws.Rows(r).Hidden = False
        End If
    Next r
End Sub

Public Sub DeleteRowsMissingLastVehicle(ws As Worksheet, anc As Range, lastVehCol As Long)
    Dim r As Long, lastR As Long
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0

    For r = lastR To anc.row + 2 Step -1
        If Trim$(ws.Cells(r, anc.Column).Value) = "" _
           Or Not HasValue(ws.Cells(r, lastVehCol).Value) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

'================= VALUE TEST =================
Public Function HasValue(x As Variant) As Boolean
    On Error Resume Next
    If IsError(x) Or IsEmpty(x) Then
        HasValue = False
        Exit Function
    End If
    
    If IsNumeric(x) Then
        HasValue = (CDbl(x) > 0)
    Else
        HasValue = (Trim$(CStr(x)) <> "" And Trim$(CStr(x)) <> "0")
    End If
    On Error GoTo 0
End Function

'================= DISCOVERY HELPERS =================
Public Function FindAnchor(ws As Worksheet, txt As String) As Range
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:=txt, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then Set FindAnchor = f
    On Error GoTo 0
End Function

'*** Correct vehicle column detection (DR-based + fallback) ***
Public Function CollectDestVehicleCols(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection
    Dim c As Long, lastC As Long, v

    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0

    'Preferred:  DR markers
    For c = anc.Column + 1 To lastC
        v = ws.Cells(anc.row + 1, c).Value
        If VarType(v) = vbString Then
            If Left$(UCase$(Trim$(v)), 2) = "DR" Then out.Add c
        End If
    Next c

    If out.count > 0 Then
        Set CollectDestVehicleCols = out
        Exit Function
    End If

    'Fallback: contiguous headers until COMMENTS
    For c = anc.Column + 1 To lastC
        v = ws.Cells(anc.row, c).Value
        If UCase$(Trim$(CStr(v))) = "COMMENTS" Then Exit For
        If Trim$(CStr(v)) <> "" Then
            out.Add c
        ElseIf out.count > 0 Then
            Exit For
        End If
    Next c

    Set CollectDestVehicleCols = out
End Function

Public Function CollectHeaders(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" Then
            out.Add ws.Cells(anc.row, c).Value
        End If
    Next c
    Set CollectHeaders = out
End Function

Public Function CollectHeaderCols(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, c As Long, lastC As Long
    
    On Error Resume Next
    lastC = ws.Cells(anc.row, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    For c = anc.Column + 1 To lastC
        If Trim$(ws.Cells(anc.row, c).Value) <> "" Then
            out.Add c
        End If
    Next c
    Set CollectHeaderCols = out
End Function

Public Function CollectRowLabels(ws As Worksheet, anc As Range) As Collection
    Dim out As New Collection, r As Long, emptyRun As Long, lastR As Long
    Dim codeCol As Long
    
    '*** Use operation codes (column before anchor) for reliable matching ***
    codeCol = anc.Column - 1
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0
    
    For r = anc.row + 2 To lastR
        If Trim$(ws.Cells(r, anc.Column).Value) <> "" Then
            ' Use operation code instead of name for matching
            out.Add Trim$(ws.Cells(r, codeCol).Value)
            emptyRun = 0
        Else
            emptyRun = emptyRun + 1
            If emptyRun >= 10 Then Exit For
        End If
    Next r
    Set CollectRowLabels = out
End Function

Public Function BuildModeIndex(ws As Worksheet, anc As Range) As Object
    Dim d As Object:  Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, v, codeVal, lastR As Long
    Dim codeCol As Long
    
    '*** Use operation codes (column before anchor) for reliable matching ***
    codeCol = anc.Column - 1
    
    On Error Resume Next
    lastR = ws.Cells(ws.Rows.count, anc.Column).End(xlUp).row
    On Error GoTo 0
    
    For r = anc.row + 2 To lastR
        v = Trim$(ws.Cells(r, anc.Column).Value)
        If v <> "" Then
            ' Use operation code as dictionary key
            codeVal = Trim$(ws.Cells(r, codeCol).Value)
            If codeVal <> "" And Not d.Exists(codeVal) Then d.Add codeVal, r
        End If
    Next r
    Set BuildModeIndex = d
End Function

'================= FORMATTING =================
Public Sub ApplyVehicleHeaderFormatting(cell As Range)
    On Error Resume Next
    With cell
        .Font.Name = "Arial"
        .Font.Size = 16
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlThick
    End With
    On Error GoTo 0
End Sub

Public Sub AutoAdjustVehicleColumns(ws As Worksheet, anc As Range, cols As Collection, n As Long)
    Dim i As Long
    On Error Resume Next
    For i = 1 To n
        ws.Columns(cols(i)).AutoFit
    Next i
    On Error GoTo 0
End Sub

'================= TEMPLATE / VISIBILITY =================
Public Sub RestoreSeparatorColumnsFromTemplate(wsTarget As Worksheet, wsTemplate As Worksheet, anc As Range)
    'Template formatting preserved – no overwrite needed
    'Add specific logic here if needed
End Sub

'*** CRITICAL: Hide unused vehicle blocks ***
Public Sub ManageVehicleColumnVisibility(ws As Worksheet, vehicleCount As Long)

    Dim vehicleCols As Variant
    Dim separatorCols As Variant
    Dim i As Long

    vehicleCols = Array(4, 6, 8, 10, 12, 14, 16)      ' D F H J L N P
    separatorCols = Array(5, 7, 9, 11, 13, 15, 17)   ' E G I K M O Q

    On Error Resume Next
    
    'Hide all first
    For i = 0 To UBound(vehicleCols)
        ws.Columns(vehicleCols(i)).Hidden = True
        ws.Columns(separatorCols(i)).Hidden = True
    Next i

    'Show only used vehicles
    For i = 0 To vehicleCount - 1
        If i <= UBound(vehicleCols) Then
            ws.Columns(vehicleCols(i)).Hidden = False
            ws.Columns(separatorCols(i)).Hidden = False
        End If
    Next i

    'Always show main separator
    ws.Columns(3).Hidden = False
    
    On Error GoTo 0

End Sub

'================= END =================

