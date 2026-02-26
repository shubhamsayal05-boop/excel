Attribute VB_Name = "OperationModeStatus"
Option Explicit

Sub Update_All_Operation_Mode_Status()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim selRange As Range
    Dim STATUS_COL As String

    ' Ask user to click any cell in the STATUS column
    On Error Resume Next
    Set selRange = Application.InputBox( _
        "Click any cell in the STATUS column (dot column)", _
        "Select Status Column", Type:=8)
    On Error GoTo 0

    If selRange Is Nothing Then
        MsgBox "No column selected. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    STATUS_COL = Split(selRange.Address, "$")(1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim grpStartRow As Long
    Dim i As Long

    grpStartRow = 0

    For i = 2 To lastRow + 1

        ' Detect group header (bold B) or end of data
        If i > lastRow Or ws.Cells(i, "B").Font.Bold = True Then

            If grpStartRow <> 0 Then

                ' Safety: never write into row 1 or row 2
                If grpStartRow - 1 > 2 Then
                    Evaluate_Group_Status ws, grpStartRow, i - 1, STATUS_COL
                End If

            End If

            If i <= lastRow Then grpStartRow = i + 1
        End If
    Next i

    MsgBox "All operation mode statuses updated.", vbInformation

End Sub


' Evaluate and write the status for a single sub-group of rows.
Sub Evaluate_Group_Status(ws As Worksheet, startRow As Long, endRow As Long, statusCol As String)

    Dim c As Range
    Dim redCnt As Long, yellowCnt As Long, totalCnt As Long
    Dim pctYellow As Double

    ' Safety: skip if the header row would be row 1 or 2
    If startRow - 1 <= 2 Then Exit Sub

    For Each c In ws.Range(statusCol & startRow & ":" & statusCol & endRow)

        If Len(c.Value) > 0 Then

            totalCnt = totalCnt + 1

            Select Case c.Font.Color
                Case RGB(255, 0, 0):    redCnt    = redCnt    + 1
                Case RGB(227, 225, 0):  yellowCnt = yellowCnt + 1
                Case RGB(0, 176, 80)
                    ' green - no counter needed
            End Select

        End If

    Next c

    If totalCnt = 0 Then Exit Sub

    Dim headerCell As Range
    Set headerCell = ws.Cells(startRow - 1, statusCol)

    With headerCell
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        If redCnt > 0 Then
            .Value = "NOK"
            .Interior.Color = RGB(255, 0, 0)

        ElseIf yellowCnt / totalCnt > 0.35 Then
            .Value = "Acceptable"
            .Interior.Color = RGB(227, 225, 0)

        Else
            .Value = "OK"
            .Interior.Color = RGB(0, 176, 80)
        End If

    End With

End Sub
