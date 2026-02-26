Attribute VB_Name = "Evaluation"
Option Explicit

' ============================================================================
' Main entry:  builds "Evaluation Results" sheet and summaries
' Uses a popup dialog for car selection.
'
' FIXES APPLIED (2025):
'   1. Operation name now read from column C (3) instead of column B (2).
'   2. Drivability P1 status now read from column F (6) instead of column E (5).
'   3. Section-header rows (e.g. "Accelerations", "Decelerations") are skipped;
'      only rows whose op-code is numeric are evaluated.
' ============================================================================
Public Sub EvaluateAVLStatus()
    Dim wsSheet1 As Worksheet
    Dim wsHeatmap As Worksheet
    Dim wsResults As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim opCode As Variant
    Dim testedAVL As Double
    Dim drivP1 As String, respP1 As String
    Dim drivTarget As Double, drivTested As Double
    Dim respTarget As Double, respTested As Double
    Dim drivBenchDiff As Double, respBenchDiff As Double
    Dim drivStatus As String, respStatus As String, finalStatus As String
    Dim outRow As Long

    ' Car selection variables
    Dim targetCarName As String, testedCarName As String
    Dim targetCol As Integer, testedCol As Integer
    Dim cols As Variant

    ' Activate Sheet1 so the user can see data while selecting cars
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet1").Activate
    On Error GoTo 0

    ' Show car selection dialog
    If Not ShowCarSelectionDialog() Then
        MsgBox "Evaluation cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Get selected car names
    targetCarName = GetSelectedTargetCar()
    testedCarName = GetSelectedTestedCar()

    ' Get column indices for selected cars
    cols = GetSelectedCarColumns()
    targetCol = cols(0)
    testedCol = cols(1)

    If targetCol = 0 Or testedCol = 0 Then
        MsgBox "Error: Could not find data columns for selected cars.", vbCritical, "Error"
        Exit Sub
    End If

    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    Set wsHeatmap = ThisWorkbook.Sheets("HeatMap Sheet")

    ' Get column indices for the Responsiveness section (separate from Drivability)
    Dim targetRespCol As Integer, testedRespCol As Integer
    targetRespCol = FindCarColumnInSection(wsSheet1, targetCarName, 12) ' Responsiveness starts at column 12
    testedRespCol = FindCarColumnInSection(wsSheet1, testedCarName, 12)

    If targetRespCol = 0 Or testedRespCol = 0 Then
        MsgBox "Error: Could not find responsiveness columns for selected cars.", vbCritical, "Error"
        Exit Sub
    End If

    ' Delete existing results sheet if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Evaluation Results").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create fresh results sheet
    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = "Evaluation Results"

    ' Header row (columns A-L) with dynamic car names
    wsResults.Range("A1:L1").Value = Array( _
        "Op Code", "Operation", "Tested AVL", _
        "Driv P1", "Driv Target (" & targetCarName & ")", "Driv Tested (" & testedCarName & ")", "Driv Status", _
        "Resp P1", "Resp Target (" & targetCarName & ")", "Resp Tested (" & testedCarName & ")", "Resp Status", "Final Status")

    With wsResults.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With

    lastRow = wsSheet1.Cells(wsSheet1.Rows.Count, 1).End(xlUp).Row
    outRow = 2

    For i = 5 To lastRow
        opCode = wsSheet1.Cells(i, 1).Value

        ' FIX 3: Skip section-header rows (e.g. "Accelerations", "Drive away").
        '         Only process rows whose op-code is a number.
        If IsNumeric(opCode) = True Then
            testedAVL = GetTestedAVL(wsHeatmap, opCode, testedCarName)

            ' FIX 2: Drivability P1 is in column F (6), NOT column E (5).
            drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 6))
            ' Responsiveness P1 is in column L (12) - unchanged.
            respP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 12))

            ' Use dynamic columns for drivability benchmark values
            drivTarget = ToDbl(wsSheet1.Cells(i, targetCol).Value)
            drivTested = ToDbl(wsSheet1.Cells(i, testedCol).Value)

            ' Responsiveness benchmark values use their own dedicated columns
            respTarget = ToDbl(wsSheet1.Cells(i, targetRespCol).Value)
            respTested = ToDbl(wsSheet1.Cells(i, testedRespCol).Value)

            drivBenchDiff = benchDiff(drivTarget, drivTested)
            respBenchDiff = benchDiff(respTarget, respTested)

            drivStatus  = EvaluateStatus(testedAVL, drivP1, drivBenchDiff, drivTarget, drivTested)
            respStatus  = EvaluateStatus(testedAVL, respP1, respBenchDiff, respTarget, respTested)
            finalStatus = CombineStatus(drivStatus, respStatus)

            wsResults.Cells(outRow, 1).Value = opCode
            ' FIX 1: Operation name is in column C (3), NOT column B (2).
            wsResults.Cells(outRow, 2).Value = wsSheet1.Cells(i, 3).Value
            wsResults.Cells(outRow, 3).Value = testedAVL
            wsResults.Cells(outRow, 4).Value = drivP1
            wsResults.Cells(outRow, 5).Value = drivTarget
            wsResults.Cells(outRow, 6).Value = drivTested
            wsResults.Cells(outRow, 7).Value = drivStatus
            wsResults.Cells(outRow, 8).Value = respP1
            wsResults.Cells(outRow, 9).Value = respTarget
            wsResults.Cells(outRow, 10).Value = respTested
            wsResults.Cells(outRow, 11).Value = respStatus
            wsResults.Cells(outRow, 12).Value = finalStatus

            ColorCell wsResults.Cells(outRow, 7), drivStatus
            ColorCell wsResults.Cells(outRow, 11), respStatus
            ColorCell wsResults.Cells(outRow, 12), finalStatus

            outRow = outRow + 1
        End If
    Next i

    wsResults.Columns("A:L").AutoFit

    ' Build "Overall Status by Op Code" summary table
    BuildUniqueOverallStatus wsResults

    MsgBox "Evaluation complete!" & vbCrLf & vbCrLf & _
           "Target:   " & targetCarName & vbCrLf & _
           "Tested:  " & testedCarName & vbCrLf & vbCrLf & _
           "Results written to sheet:   " & wsResults.Name, _
           vbInformation, "Success"
End Sub

' ============================================================================
' Builds "Overall Status by Op Code" summary at the bottom of the results sheet.
' N/A statuses are excluded from the RED/YELLOW/GREEN roll-up.
' ============================================================================
Private Sub BuildUniqueOverallStatus(wsResults As Worksheet)
    Dim lastRowRes As Long, i As Long
    Dim code As String, status As String

    Dim codes() As String
    Dim names() As String
    Dim statuses() As String
    Dim codeCount As Long
    Dim foundIndex As Long

    lastRowRes = wsResults.Cells(wsResults.Rows.Count, 1).End(xlUp).Row

    ReDim codes(1 To 1)
    ReDim names(1 To 1)
    ReDim statuses(1 To 1)
    codeCount = 0

    ' Collect all codes and their final statuses
    For i = 2 To lastRowRes
        code = Trim(CStr(wsResults.Cells(i, 1).Value))

        If code <> "" Then
            status = Trim(CStr(wsResults.Cells(i, 12).Value))

            foundIndex = FindInArray(codes, code, codeCount)

            If foundIndex = 0 Then
                codeCount = codeCount + 1
                ReDim Preserve codes(1 To codeCount)
                ReDim Preserve names(1 To codeCount)
                ReDim Preserve statuses(1 To codeCount)

                codes(codeCount) = code
                names(codeCount) = Trim(CStr(wsResults.Cells(i, 2).Value))
                statuses(codeCount) = status
            Else
                statuses(foundIndex) = statuses(foundIndex) & "|" & status
            End If
        End If
    Next i

    ' Summary header
    Dim startRow As Long
    startRow = lastRowRes + 2

    wsResults.Cells(startRow, 1).Value = "Overall Status by Op Code"
    With wsResults.Range(wsResults.Cells(startRow, 1), wsResults.Cells(startRow, 4))
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    wsResults.Cells(startRow + 1, 1).Value = "Op Code"
    wsResults.Cells(startRow + 1, 2).Value = "Operation"
    wsResults.Cells(startRow + 1, 3).Value = "Overall Status"
    wsResults.Range(wsResults.Cells(startRow + 1, 1), wsResults.Cells(startRow + 1, 3)).Font.Bold = True

    Dim r As Long
    r = startRow + 2

    For i = 1 To codeCount
        Dim statusList() As String
        Dim anyRed As Boolean
        Dim allGreen As Boolean
        Dim hasValidStatus As Boolean
        Dim overall As String
        Dim j As Long

        anyRed = False
        allGreen = True
        hasValidStatus = False

        statusList = Split(statuses(i), "|")

        ' Exclude N/A from RED/YELLOW/GREEN evaluation
        For j = LBound(statusList) To UBound(statusList)
            status = Trim(statusList(j))

            If status <> "" And status <> "N/A" Then
                hasValidStatus = True
                If status = "RED" Then anyRed = True
                If status <> "GREEN" Then allGreen = False
            End If
        Next j

        If Not hasValidStatus Then
            overall = "N/A"
        ElseIf anyRed Then
            overall = "RED"
        ElseIf allGreen Then
            overall = "GREEN"
        Else
            overall = "YELLOW"
        End If

        wsResults.Cells(r, 1).Value = codes(i)
        wsResults.Cells(r, 2).Value = names(i)
        wsResults.Cells(r, 3).Value = overall
        ColorCell wsResults.Cells(r, 3), overall
        r = r + 1
    Next i

    wsResults.Columns("A:C").AutoFit
End Sub

' ============================================================================
' Find string in array - returns 1-based index, or 0 if not found
' ============================================================================
Private Function FindInArray(arr() As String, searchValue As String, arraySize As Long) As Long
    Dim i As Long
    FindInArray = 0

    For i = 1 To arraySize
        If arr(i) = searchValue Then
            FindInArray = i
            Exit Function
        End If
    Next i
End Function

' ============================================================================
' Convert Variant to Double safely (returns 0 for non-numeric values)
' ============================================================================
Private Function ToDbl(v As Variant) As Double
    If IsNumeric(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = 0
    End If
End Function

' ============================================================================
' Benchmark difference.
' Returns 999 (sentinel) when target is zero or both are zero.
' ============================================================================
Private Function benchDiff(targetVal As Double, testedVal As Double) As Double
    If targetVal = 0 And testedVal = 0 Then
        benchDiff = 999
    ElseIf targetVal = 0 Then
        benchDiff = 999
    Else
        benchDiff = Abs(testedVal - targetVal)
    End If
End Function

' ============================================================================
' Look up the Tested AVL score from the HeatMap sheet for a given op-code
' and the selected tested vehicle column.
' ============================================================================
Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant, testedCarName As String) As Double
    Dim opKey As String
    Dim f As Range
    Dim avlCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Range
    Dim col As Long

    ' Find the column for the tested vehicle in HeatMap (vehicle names in row 2)
    avlCol = 0
    lastCol = wsHeatmap.Cells(2, wsHeatmap.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        If Trim(CStr(wsHeatmap.Cells(2, col).Value)) = Trim(testedCarName) Then
            avlCol = col
            Exit For
        End If
    Next col

    ' Default to column 8 when vehicle column not found (backward compatibility)
    If avlCol = 0 Then avlCol = 8

    opKey = Trim(CStr(opCode))

    ' 1. Exact string match
    Set f = wsHeatmap.Columns(1).Find(What:=opKey, LookIn:=xlValues, LookAt:=xlWhole, _
                                      MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If Not f Is Nothing Then
        GetTestedAVL = ToDbl(wsHeatmap.Cells(f.Row, avlCol).Value)
        Exit Function
    End If

    ' 2. Numeric match
    If IsNumeric(opKey) Then
        Set f = wsHeatmap.Columns(1).Find(What:=CLng(Val(opKey)), LookIn:=xlValues, LookAt:=xlWhole, _
                                          MatchCase:=False)
        If Not f Is Nothing Then
            GetTestedAVL = ToDbl(wsHeatmap.Cells(f.Row, avlCol).Value)
            Exit Function
        End If
    End If

    ' 3. Manual loop with trimmed string compare (fallback)
    lastRow = wsHeatmap.Cells(wsHeatmap.Rows.Count, 1).End(xlUp).Row

    For Each c In wsHeatmap.Range(wsHeatmap.Cells(1, 1), wsHeatmap.Cells(lastRow, 1))
        If Trim(CStr(c.Value)) = opKey Then
            GetTestedAVL = ToDbl(wsHeatmap.Cells(c.Row, avlCol).Value)
            Exit Function
        End If
    Next c

    GetTestedAVL = 0
End Function

' ============================================================================
' Determine P1 status from cell fill / font color.
' Prefers DisplayFormat; falls back to Interior / Font.
' ============================================================================
Private Function GetP1StatusFromColor(rng As Range) As String
    On Error GoTo Fallback

    Dim clr As Long, fclr As Long

    clr  = rng.DisplayFormat.Interior.Color
    fclr = rng.DisplayFormat.Font.Color

    GetP1StatusFromColor = MapColorToStatus(clr, fclr)
    If GetP1StatusFromColor <> "N/A" Then Exit Function

Fallback:
    On Error Resume Next
    clr  = rng.Interior.Color
    fclr = rng.Font.Color

    GetP1StatusFromColor = MapColorToStatus(clr, fclr)

    If GetP1StatusFromColor = "" Then GetP1StatusFromColor = "N/A"
End Function

' ============================================================================
' Map fill color / font color to GREEN / YELLOW / RED / N/A
' ============================================================================
Private Function MapColorToStatus(fillClr As Long, fontClr As Long) As String
    Dim r As Long, g As Long, b As Long
    Dim rf As Long, gf As Long, bf As Long

    ' Check fill color first
    If fillClr > 0 Then
        r = fillClr Mod 256
        g = (fillClr \ 256) Mod 256
        b = (fillClr \ 65536) Mod 256

        If IsNearRGB(r, g, b, 0, 176, 80, 45) Or IsNearRGB(r, g, b, 0, 158, 71, 45) Then
            MapColorToStatus = "GREEN":  Exit Function
        End If

        If IsNearRGB(r, g, b, 255, 192, 0, 45) Or IsNearRGB(r, g, b, 255, 217, 102, 60) Then
            MapColorToStatus = "YELLOW": Exit Function
        End If

        If IsNearRGB(r, g, b, 255, 0, 0, 45) Or IsNearRGB(r, g, b, 192, 0, 0, 45) Then
            MapColorToStatus = "RED":    Exit Function
        End If
    End If

    ' Check font color
    If fontClr > 0 Then
        rf = fontClr Mod 256
        gf = (fontClr \ 256) Mod 256
        bf = (fontClr \ 65536) Mod 256

        If IsNearRGB(rf, gf, bf, 0, 128, 0, 5) Then
            MapColorToStatus = "GREEN":  Exit Function
        End If

        If IsNearRGB(rf, gf, bf, 255, 255, 0, 5) Then
            MapColorToStatus = "YELLOW": Exit Function
        End If

        If IsNearRGB(rf, gf, bf, 0, 176, 80, 45) Or IsNearRGB(rf, gf, bf, 0, 158, 71, 45) Then
            MapColorToStatus = "GREEN":  Exit Function
        End If

        If IsNearRGB(rf, gf, bf, 255, 192, 0, 45) Or IsNearRGB(rf, gf, bf, 255, 217, 102, 60) Then
            MapColorToStatus = "YELLOW": Exit Function
        End If

        If IsNearRGB(rf, gf, bf, 255, 0, 0, 45) Or IsNearRGB(rf, gf, bf, 192, 0, 0, 45) Then
            MapColorToStatus = "RED":    Exit Function
        End If
    End If

    MapColorToStatus = "N/A"
End Function

' ============================================================================
' Test whether (r,g,b) is within tolerance of (rt,gt,bt)
' ============================================================================
Private Function IsNearRGB(r As Long, g As Long, b As Long, _
                           rt As Long, gt As Long, bt As Long, tol As Long) As Boolean
    IsNearRGB = (Abs(r - rt) <= tol) And _
                (Abs(g - gt) <= tol) And _
                (Abs(b - bt) <= tol)
End Function

' ============================================================================
' Evaluate individual status using AVL score, P1 color and benchmark difference.
'
' Rules:
'   1. P1 = N/A                               -> N/A
'   2. AVL < 7  OR  P1 = RED                  -> RED
'   3. AVL >= 7 AND P1 = YELLOW               -> YELLOW
'   4. AVL >= 7 AND P1 = GREEN AND no bench   -> GREEN
'   5. AVL >= 7 AND P1 = GREEN AND bench OK   -> GREEN
'   6. AVL >= 7 AND P1 = GREEN AND bench fail -> YELLOW
' ============================================================================
Private Function EvaluateStatus(avl As Double, p1 As String, benchDiff As Double, _
                                targetVal As Double, testedVal As Double) As String

    If UCase(Trim(p1)) = "N/A" Then
        EvaluateStatus = "N/A"
        Exit Function
    End If

    ' Rule 2: AVL < 7 or P1 RED always gives RED
    If avl < 7 Or UCase(Trim(p1)) = "RED" Then
        EvaluateStatus = "RED"
        Exit Function
    End If

    ' Rule 3: P1 YELLOW always gives YELLOW
    If avl >= 7 And UCase(Trim(p1)) = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If

    ' Below here: AVL >= 7 AND P1 = GREEN

    ' Rule 4: No benchmark data -> GREEN
    If benchDiff = 999 Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Rule 4 (also): non-numeric benchmark -> GREEN
    If Not IsNumeric(targetVal) Or Not IsNumeric(testedVal) Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Rule 5 & 6: Evaluate benchmark (tolerance = 2 points)
    If testedVal >= targetVal Then
        EvaluateStatus = "GREEN"
    ElseIf (targetVal - testedVal) <= 2 Then
        EvaluateStatus = "GREEN"
    Else
        EvaluateStatus = "YELLOW"
    End If
End Function

' ============================================================================
' Combine Drivability and Responsiveness statuses into a single Final Status.
' ============================================================================
Private Function CombineStatus(drivStatus As String, respStatus As String) As String
    Dim driv As String, resp As String

    driv = UCase(Trim(drivStatus))
    resp = UCase(Trim(respStatus))

    If driv = "" Then driv = "N/A"
    If resp = "" Then resp = "N/A"

    ' Rule 1: Either RED -> RED
    If driv = "RED" Or resp = "RED" Then
        CombineStatus = "RED":  Exit Function
    End If

    ' Rule 2: Either YELLOW -> YELLOW
    If driv = "YELLOW" Or resp = "YELLOW" Then
        CombineStatus = "YELLOW": Exit Function
    End If

    ' Rule 3: Both GREEN -> GREEN
    If driv = "GREEN" And resp = "GREEN" Then
        CombineStatus = "GREEN": Exit Function
    End If

    ' Rule 4: One GREEN, one N/A -> GREEN
    If (driv = "GREEN" And resp = "N/A") Or (driv = "N/A" And resp = "GREEN") Then
        CombineStatus = "GREEN": Exit Function
    End If

    ' Rule 5: Both N/A (or any other combination) -> N/A
    CombineStatus = "N/A"
End Function

' ============================================================================
' Apply fill and font color to a result cell based on status string.
' ============================================================================
Private Sub ColorCell(c As Range, s As String)
    Select Case UCase$(s)
        Case "GREEN"
            c.Interior.Color = RGB(0, 176, 80)
            c.Font.Color = vbWhite
        Case "YELLOW"
            c.Interior.Color = RGB(255, 192, 0)
            c.Font.Color = vbBlack
        Case "RED"
            c.Interior.Color = RGB(192, 0, 0)
            c.Font.Color = vbWhite
        Case Else
            c.Interior.ColorIndex = xlNone
            c.Font.Color = vbBlack
    End Select
End Sub

' ============================================================================
' Find the column in a specific section (Drivability or Responsiveness)
' by searching from startCol onwards in row 2 for the given car name.
' ============================================================================
Private Function FindCarColumnInSection(ws As Worksheet, carName As String, startCol As Integer) As Integer
    Dim col As Integer
    Dim cellValue As String
    Dim lastCol As Integer

    FindCarColumnInSection = 0

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    For col = startCol To lastCol
        cellValue = Trim(CStr(ws.Cells(2, col).Value))
        If cellValue = Trim(carName) Then
            FindCarColumnInSection = col
            Exit Function
        End If
    Next col
End Function
