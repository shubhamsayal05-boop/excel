Attribute VB_Name = "Updatesuboperationstatus"
Option Explicit

Sub UpdateSubOperationHeatMap()

    Dim wsEval As Worksheet
    Dim wsHeat As Worksheet
    Dim lastEvalRow As Long
    Dim lastHeatRow As Long
    Dim i As Long
    Dim opCode As String
    Dim status As String
    Dim dict As Object
    Dim tgt As Range
    Dim wasProtected As Boolean
    Dim sheetPassword As String

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Set wsEval = ThisWorkbook.Worksheets("Evaluation Results")
    Set wsHeat = ThisWorkbook.Worksheets("HeatMap Sheet")
    Set dict = CreateObject("Scripting.Dictionary")

    ' ---------------------------------
    ' Check if sheet is protected; unprotect if needed
    ' ---------------------------------
    wasProtected = wsHeat.ProtectContents
    If wasProtected Then
        sheetPassword = "" ' Change this if your sheet has a password
        On Error Resume Next
        wsHeat.Unprotect Password:=sheetPassword
        If Err.Number <> 0 Then
            MsgBox "Unable to unprotect the HeatMap Sheet. Please enter the correct password or unprotect manually.", vbCritical
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    End If

    ' ---------------------------------
    ' 1. Read the Overall Status summary from Evaluation Results
    '    The summary block starts below the detail rows; dictionary key = op code,
    '    value = overall status (RED / YELLOW / GREEN / N/A).
    ' ---------------------------------
    lastEvalRow = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastEvalRow
        If wsEval.Cells(i, "A").Value <> "" Then
            dict(Trim(CStr(wsEval.Cells(i, "A").Value))) = _
                UCase(Trim(wsEval.Cells(i, "C").Value))
        End If
    Next i

    ' ---------------------------------
    ' 2. Write colored bullet dots to the HeatMap Sheet
    '    Only non-bold rows (sub-operations) are updated.
    ' ---------------------------------
    lastHeatRow = wsHeat.Cells(wsHeat.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastHeatRow

        ' Sub-operation = NOT bold column B
        If wsHeat.Cells(i, "B").Font.Bold = False Then

            opCode = Trim(CStr(wsHeat.Cells(i, "A").Value))

            If opCode <> "" Then

                Set tgt = wsHeat.Cells(i, "R").MergeArea
                tgt.ClearContents

                If dict.Exists(opCode) Then

                    status = dict(opCode)

                    With tgt
                        .Value = ChrW(&H25CF)   ' Bullet character ●
                        .Font.Size = 14
                        .HorizontalAlignment = xlCenter

                        Select Case status
                            Case "RED"
                                .Font.Color = RGB(255, 0, 0)
                            Case "YELLOW"
                                .Font.Color = RGB(227, 225, 0)
                            Case "GREEN"
                                .Font.Color = RGB(0, 176, 80)
                            Case Else
                                .ClearContents
                        End Select
                    End With

                End If
            End If
        End If
    Next i

    ' ---------------------------------
    ' Re-protect sheet if it was protected
    ' ---------------------------------
    If wasProtected Then
        wsHeat.Protect Password:=sheetPassword, _
                      DrawingObjects:=True, _
                      Contents:=True, _
                      Scenarios:=True
    End If

    Application.ScreenUpdating = True

    MsgBox "Sub-operation HeatMap updated successfully.", vbInformation

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True

    ' Try to re-protect sheet even if an error occurred
    If wasProtected Then
        On Error Resume Next
        wsHeat.Protect Password:=sheetPassword
        On Error GoTo 0
    End If

    MsgBox "An error occurred:  " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Error"

End Sub
