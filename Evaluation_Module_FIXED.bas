VBA MACRO Evaluation.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Evaluation'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

' ============================================================================
' Main entry:    builds "Evaluation Results" sheet and summaries
' Now uses popup dialog for car selection
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
    
    ' NEW: Car selection variables
    Dim targetCarName As String, testedCarName As String
    Dim targetCol As Integer, testedCol As Integer
    Dim cols As Variant
    
    ' Activate Sheet1 so user can see data when selecting cars
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet1").Activate
    On Error GoTo 0
    
    ' NEW: Show car selection dialog
    If Not ShowCarSelectionDialog() Then
        MsgBox "Evaluation cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' NEW:    Get selected car names
    targetCarName = GetSelectedTargetCar()
    testedCarName = GetSelectedTestedCar()
    
    ' NEW:  Get column indices for selected cars
    cols = GetSelectedCarColumns()
    targetCol = cols(0)
    testedCol = cols(1)
    
    If targetCol = 0 Or testedCol = 0 Then
        MsgBox "Error: Could not find data columns for selected cars.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    Set wsHeatmap = ThisWorkbook.Sheets("HeatMap Sheet")
    
    ' NEW: Get column indices for responsiveness section (separate from drivability)
    Dim targetRespCol As Integer, testedRespCol As Integer
    targetRespCol = FindCarColumnInSection(wsSheet1, targetCarName, 12)  ' Responsiveness starts around column 12
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
    
    ' Create results sheet
    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsResults.Name = "Evaluation Results"
    
    ' Header row with car names
    wsResults.Range("A1:L1").Value = Array( _
        "Op Code", "Operation", "Tested AVL", _
        "Driv P1", "Driv Target (" & targetCarName & ")", "Driv Tested (" & testedCarName & ")", "Driv Status", _
        "Resp P1", "Resp Target (" & targetCarName & ")", "Resp Tested (" & testedCarName & ")", "Resp Status", "Final Status")
    
    With wsResults.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With
    
    lastRow = wsSheet1.Cells(wsSheet1.Rows.count, 1).End(xlUp).row
    outRow = 2
    
    For i = 5 To lastRow
        opCode = wsSheet1.Cells(i, 1).Value
        
        If Trim(CStr(opCode)) <> "" Then
            testedAVL = GetTestedAVL(wsHeatmap, opCode, testedCarName)
            drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 5))
            respP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 12))
            
            ' MODIFIED: Use dynamic columns instead of fixed columns
            drivTarget = ToDbl(wsSheet1.Cells(i, targetCol).Value)
            drivTested = ToDbl(wsSheet1.Cells(i, testedCol).Value)
            
            ' Resp columns are in separate section - use specific responsiveness columns
            respTarget = ToDbl(wsSheet1.Cells(i, targetRespCol).Value)
            respTested = ToDbl(wsSheet1.Cells(i, testedRespCol).Value)
            
            drivBenchDiff = benchDiff(drivTarget, drivTested)
            respBenchDiff = benchDiff(respTarget, respTested)
            
            drivStatus = EvaluateStatus(testedAVL, drivP1, drivBenchDiff, drivTarget, drivTested)
            respStatus = EvaluateStatus(testedAVL, respP1, respBenchDiff, respTarget, respTested)
            finalStatus = CombineStatus(drivStatus, respStatus)
            
            wsResults.Cells(outRow, 1).Value = opCode
            wsResults.Cells(outRow, 2).Value = wsSheet1.Cells(i, 2).Value
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
    
    ' ? Only build the "Overall Status by Op Code" table (Operation Mode Summary removed)
    BuildUniqueOverallStatus wsResults
    
    MsgBox "Evaluation complete!" & vbCrLf & vbCrLf & _
           "Target:   " & targetCarName & vbCrLf & _
           "Tested:  " & testedCarName & vbCrLf & vbCrLf & _
           "Results written to sheet:   " & wsResults.Name, _
           vbInformation, "Success"
End Sub

' ============================================================================
' Builds overall status by op code (unique codes) - FIXED N/A HANDLING
' ============================================================================
Private Sub BuildUniqueOverallStatus(wsResults As Worksheet)
    Dim lastRowRes As Long, i As Long
    Dim code As String, status As String
    
    ' Arrays to store unique codes and their data
    Dim codes() As String
    Dim names() As String
    Dim statuses() As String
    Dim codeCount As Long
    Dim foundIndex As Long
    
    lastRowRes = wsResults.Cells(wsResults.Rows.count, 1).End(xlUp).row
    
    ' Initialize arrays
    ReDim codes(1 To 1)
    ReDim names(1 To 1)
    ReDim statuses(1 To 1)
    codeCount = 0
    
    ' Collect all codes and statuses
    For i = 2 To lastRowRes
        code = Trim(CStr(wsResults.Cells(i, 1).Value))
        
        If code <> "" Then
            status = Trim(CStr(wsResults.Cells(i, 12).Value))
            
            ' Find if code already exists
            foundIndex = FindInArray(codes, code, codeCount)
            
            If foundIndex = 0 Then
                ' New code - add it
                codeCount = codeCount + 1
                ReDim Preserve codes(1 To codeCount)
                ReDim Preserve names(1 To codeCount)
                ReDim Preserve statuses(1 To codeCount)
                
                codes(codeCount) = code
                names(codeCount) = Trim(CStr(wsResults.Cells(i, 2).Value))
                statuses(codeCount) = status
            Else
                ' Existing code - append status with delimiter
                statuses(foundIndex) = statuses(foundIndex) & "|" & status
            End If
        End If
    Next i
    
    ' Build summary section
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
    
    ' Process each unique code
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
        
        ' Split statuses by delimiter
        statusList = Split(statuses(i), "|")
        
        ' ? FIXED: Exclude N/A from evaluation
        For j = LBound(statusList) To UBound(statusList)
            status = Trim(statusList(j))
            
            ' Only evaluate non-N/A statuses
            If status <> "" And status <> "N/A" Then
                hasValidStatus = True
                If status = "RED" Then anyRed = True
                If status <> "GREEN" Then allGreen = False
            End If
        Next j
        
        ' Determine overall status
        If Not hasValidStatus Then
            ' All statuses are N/A
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
' Find string in array - returns index or 0 if not found
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
' Convert variant to double safely
' ============================================================================
Private Function ToDbl(v As Variant) As Double
    If IsNumeric(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = 0
    End If
End Function

' ============================================================================
' Bench difference:    uses 999 as sentinel when target/tested are both zero or target is zero
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
' Look up Tested AVL from HeatMap sheet - reads from tested vehicle's column
' ============================================================================
Private Function GetTestedAVL(wsHeatmap As Worksheet, opCode As Variant, testedCarName As String) As Double
    Dim opKey As String
    Dim f As Range
    Dim avlCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Range
    Dim col As Long
    
    ' Find the column for the tested vehicle in HeatMap sheet
    ' Vehicle names are in row 2
    avlCol = 0
    lastCol = wsHeatmap.Cells(2, wsHeatmap.Columns.count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If Trim(CStr(wsHeatmap.Cells(2, col).Value)) = Trim(testedCarName) Then
            avlCol = col
            Exit For
        End If
    Next col
    
    ' If vehicle column not found, default to column 8 for backward compatibility
    If avlCol = 0 Then
        avlCol = 8
    End If
    
    opKey = Trim(CStr(opCode))
    
    ' Try exact find (string)
    Set f = wsHeatmap.Columns(1).Find(What:=opKey, LookIn:=xlValues, LookAt:=xlWhole, _
                                      MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    If Not f Is Nothing Then
        GetTestedAVL = ToDbl(wsHeatmap.Cells(f.row, avlCol).Value)
        Exit Function
    End If
    
    ' Try numeric match (if opKey numeric-looking)
    If IsNumeric(opKey) Then
        Set f = wsHeatmap.Columns(1).Find(What:=CLng(val(opKey)), LookIn:=xlValues, LookAt:=xlWhole, _
                                          MatchCase:=False)
        
        If Not f Is Nothing Then
            GetTestedAVL = ToDbl(wsHeatmap.Cells(f.row, avlCol).Value)
            Exit Function
        End If
    End If
    
    ' Fall back to manual loop (trimmed string compare)
    lastRow = wsHeatmap.Cells(wsHeatmap.Rows.count, 1).End(xlUp).row
    
    For Each c In wsHeatmap.Range(wsHeatmap.Cells(1, 1), wsHeatmap.Cells(lastRow, 1))
        If Trim(CStr(c.Value)) = opKey Then
            GetTestedAVL = ToDbl(wsHeatmap.Cells(c.row, avlCol).Value)
            Exit Function
        End If
    Next c
    
    GetTestedAVL = 0
End Function

' ============================================================================
' Determine P1 status from cell color (prefers DisplayFormat, falls back to Interior/Font)
' ============================================================================
Private Function GetP1StatusFromColor(rng As Range) As String
    On Error GoTo Fallback
    
    Dim clr As Long, fclr As Long
    
    clr = rng.DisplayFormat.Interior.Color
    fclr = rng.DisplayFormat.Font.Color
    
    GetP1StatusFromColor = MapColorToStatus(clr, fclr)
    If GetP1StatusFromColor <> "N/A" Then Exit Function
    
Fallback:
    On Error Resume Next
    clr = rng.Interior.Color
    fclr = rng.Font.Color
    
    GetP1StatusFromColor = MapColorToStatus(clr, fclr)
    
    If GetP1StatusFromColor = "" Then GetP1StatusFromColor = "N/A"
End Function

' ============================================================================
' Map fill/font RGB to GREEN / YELLOW / RED / N/A
' ============================================================================
Private Function MapColorToStatus(fillClr As Long, fontClr As Long) As String
    Dim r As Long, g As Long, b As Long
    Dim rf As Long, gf As Long, bf As Long
    
    ' Check fill color
    If fillClr > 0 Then
        r = fillClr Mod 256
        g = (fillClr \ 256) Mod 256
        b = (fillClr \ 65536) Mod 256
        
        If IsNearRGB(r, g, b, 0, 176, 80, 45) Or IsNearRGB(r, g, b, 0, 158, 71, 45) Then
            MapColorToStatus = "GREEN":    Exit Function
        End If
        
        If IsNearRGB(r, g, b, 255, 192, 0, 45) Or IsNearRGB(r, g, b, 255, 217, 102, 60) Then
            MapColorToStatus = "YELLOW":  Exit Function
        End If
        
        If IsNearRGB(r, g, b, 255, 0, 0, 45) Or IsNearRGB(r, g, b, 192, 0, 0, 45) Then
            MapColorToStatus = "RED":  Exit Function
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
            MapColorToStatus = "YELLOW":  Exit Function
        End If
        
        If IsNearRGB(rf, gf, bf, 0, 176, 80, 45) Or IsNearRGB(rf, gf, bf, 0, 158, 71, 45) Then
            MapColorToStatus = "GREEN": Exit Function
        End If
        
        If IsNearRGB(rf, gf, bf, 255, 192, 0, 45) Or IsNearRGB(rf, gf, bf, 255, 217, 102, 60) Then
            MapColorToStatus = "YELLOW": Exit Function
        End If
        
        If IsNearRGB(rf, gf, bf, 255, 0, 0, 45) Or IsNearRGB(rf, gf, bf, 192, 0, 0, 45) Then
            MapColorToStatus = "RED": Exit Function
        End If
    End If
    
    MapColorToStatus = "N/A"
End Function

' ============================================================================
' RGB proximity test
' ============================================================================
Private Function IsNearRGB(r As Long, g As Long, b As Long, _
                          rt As Long, gt As Long, bt As Long, tol As Long) As Boolean
    IsNearRGB = (Abs(r - rt) <= tol) And _
                (Abs(g - gt) <= tol) And _
                (Abs(b - bt) <= tol)
End Function

' ============================================================================
' Evaluate status using AVL, P1 color and bench difference
' ============================================================================
Private Function EvaluateStatus(avl As Double, p1 As String, benchDiff As Double, _
                                targetVal As Double, testedVal As Double) As String
    ' Updated evaluation logic per specification:
    ' 1. AVL >= 7 AND P1 = GREEN AND meeting benchmark ? GREEN (OK)
    ' 2. AVL >= 7 AND P1 = GREEN AND NOT meeting benchmark ? YELLOW
    ' 3. AVL >= 7 AND P1 = YELLOW AND meeting benchmark ? YELLOW
    ' 4. AVL >= 7 AND P1 = YELLOW AND NOT meeting benchmark ? YELLOW
    ' 5. AVL < 7 OR P1 = RED ? RED
    ' 6. AVL < 7 OR P1 = RED AND meeting benchmark ? RED
    ' 7. If no benchmark data ? ignore benchmark and evaluate on AVL and P1 only
    
    ' Return "N/A" string instead of empty string
    If UCase(Trim(p1)) = "N/A" Then
        EvaluateStatus = "N/A"
        Exit Function
    End If
    
    ' Rule 5 & 6: If AVL < 7 OR P1 = RED ? Always RED (regardless of benchmark)
    If avl < 7 Or UCase(Trim(p1)) = "RED" Then
        EvaluateStatus = "RED"
        Exit Function
    End If
    
    ' Rule 3 & 4: If P1 = YELLOW ? Always YELLOW (regardless of benchmark)
    If avl >= 7 And UCase(Trim(p1)) = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If
    
    ' At this point:  AVL >= 7 AND P1 = GREEN
    ' Need to check benchmark data
    
    ' If benchmark data is missing, ignore it and evaluate on AVL/P1 only
    If benchDiff = 999 Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If
    
    ' If benchmark values not numeric, ignore benchmark
    If Not IsNumeric(targetVal) Or Not IsNumeric(testedVal) Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If
    
    ' Benchmark data is available, evaluate it
    If testedVal >= targetVal Then
        EvaluateStatus = "GREEN"
    Else
        If (targetVal - testedVal) <= 2 Then
            EvaluateStatus = "GREEN"
        Else
            EvaluateStatus = "YELLOW"
        End If
    End If
End Function

' ============================================================================
' Combine drive & response statuses into final
' ============================================================================
Private Function CombineStatus(drivStatus As String, respStatus As String) As String
    Dim driv As String, resp As String
    
    ' Normalize to uppercase and treat empty strings as "N/A"
    driv = UCase(Trim(drivStatus))
    resp = UCase(Trim(respStatus))
    
    If driv = "" Then driv = "N/A"
    If resp = "" Then resp = "N/A"
    
    ' Rule 1: If either is RED ? Final is RED (critical failure)
    If driv = "RED" Or resp = "RED" Then
        CombineStatus = "RED"
        Exit Function
    End If
    
    ' Rule 2: If either is YELLOW ? Final is YELLOW (warning)
    If driv = "YELLOW" Or resp = "YELLOW" Then
        CombineStatus = "YELLOW"
        Exit Function
    End If
    
    ' Rule 3: If both are GREEN ? Final is GREEN (perfect)
    If driv = "GREEN" And resp = "GREEN" Then
        CombineStatus = "GREEN"
        Exit Function
    End If
    
    ' Rule 4: If one is GREEN and the other is N/A ? Final is GREEN
    If (driv = "GREEN" And resp = "N/A") Or (driv = "N/A" And resp = "GREEN") Then
        CombineStatus = "GREEN"
        Exit Function
    End If
    
    ' Rule 5: If both are N/A or any other combination ? Final is N/A
    CombineStatus = "N/A"
End Function

' ============================================================================
' Color cell based on status string
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
' Find car column in a specific section (Drivability or Responsiveness)
' Searches from startCol onwards in row 2 for the car name
' ============================================================================
Private Function FindCarColumnInSection(ws As Worksheet, carName As String, startCol As Integer) As Integer
    Dim col As Integer
    Dim cellValue As String
    Dim lastCol As Integer
    
    FindCarColumnInSection = 0
    
    ' Find last column with data in row 2
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    
    ' Search from startCol to lastCol
    For col = startCol To lastCol
        cellValue = Trim(CStr(ws.Cells(2, col).Value))
        
        ' Match the car name
        If cellValue = Trim(carName) Then
            FindCarColumnInSection = col
            Exit Function
        End If
    Next col
    
End Function

-------------------------------------------------------------------------------
