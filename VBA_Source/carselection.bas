Attribute VB_Name = "carselection"
Option Explicit

' Module-level variables to store selections
Private m_TargetCar As String
Private m_TestedCar As String
Private m_UserCancelled As Boolean

' Constants
Private Const DATA_SHEET_NAME As String = "Sheet1"
Private Const CAR_DATA_START_COL As Integer = 8  ' Column H

' ===================================================================
' PUBLIC FUNCTIONS
' ===================================================================

' ShowCarSelectionDialog
' Displays a popup dialog for selecting the Target and Tested cars.
' Returns: True if the user clicked OK, False if the user cancelled.
Public Function ShowCarSelectionDialog() As Boolean
    Dim ws As Worksheet
    Dim carNames() As String
    Dim i As Integer
    Dim targetCar As String
    Dim testedCar As String
    Dim response As VbMsgBoxResult

    ShowCarSelectionDialog = False
    m_UserCancelled = True
    m_TargetCar = ""
    m_TestedCar = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Error: Could not find '" & DATA_SHEET_NAME & "' worksheet.", vbCritical, "Error"
        Exit Function
    End If

    carNames = GetAvailableCarNames(ws)

    If UBound(carNames) < 0 Then
        MsgBox "No car names found in the data sheet." & vbCrLf & vbCrLf & _
               "Please ensure car names are in row 2, starting from column H.", _
               vbExclamation, "No Cars Found"
        Exit Function
    End If

    ' Build a readable list of available cars
    Dim carList As String
    carList = Join(carNames, vbCrLf)

    ' --- Prompt for Target car ---
    targetCar = Application.InputBox( _
        "Available cars:" & vbCrLf & carList & vbCrLf & vbCrLf & _
        "Enter the TARGET car name:", _
        "Select Target Car", _
        Type:=2)  ' Type 2 = text

    If targetCar = "False" Or Trim(targetCar) = "" Then Exit Function

    If Not IsCarNameValid(targetCar, carNames) Then
        MsgBox "Invalid Target car name: " & targetCar & vbCrLf & vbCrLf & _
               "Please enter one of the available car names exactly as shown.", _
               vbExclamation, "Invalid Selection"
        Exit Function
    End If

    ' --- Prompt for Tested car ---
    testedCar = Application.InputBox( _
        "Available cars:" & vbCrLf & carList & vbCrLf & vbCrLf & _
        "Enter the TESTED car name:", _
        "Select Tested Car", _
        Type:=2)  ' Type 2 = text

    If testedCar = "False" Or Trim(testedCar) = "" Then Exit Function

    If Not IsCarNameValid(testedCar, carNames) Then
        MsgBox "Invalid Tested car name: " & testedCar & vbCrLf & vbCrLf & _
               "Please enter one of the available car names exactly as shown.", _
               vbExclamation, "Invalid Selection"
        Exit Function
    End If

    ' Warn if same car selected for both
    If targetCar = testedCar Then
        response = MsgBox( _
            "You have selected the same car for both Target and Tested:" & vbCrLf & vbCrLf & _
            "    " & targetCar & vbCrLf & vbCrLf & _
            "This will compare the car against itself. Continue?", _
            vbQuestion + vbYesNo, "Same Car Selected")

        If response = vbNo Then Exit Function
    End If

    ' Store selections
    m_TargetCar = targetCar
    m_TestedCar = testedCar
    m_UserCancelled = False
    ShowCarSelectionDialog = True
End Function

' GetSelectedTargetCar
' Returns the Target car name stored by ShowCarSelectionDialog.
Public Function GetSelectedTargetCar() As String
    GetSelectedTargetCar = m_TargetCar
End Function

' GetSelectedTestedCar
' Returns the Tested car name stored by ShowCarSelectionDialog.
Public Function GetSelectedTestedCar() As String
    GetSelectedTestedCar = m_TestedCar
End Function

' GetSelectedCarColumns
' Returns a two-element array: Array(targetColIndex, testedColIndex).
' Returns Array(0, 0) if either column cannot be found.
Public Function GetSelectedCarColumns() As Variant
    Dim ws As Worksheet
    Dim targetCol As Integer
    Dim testedCol As Integer
    Dim result(1) As Integer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        GetSelectedCarColumns = Array(0, 0)
        Exit Function
    End If

    targetCol = FindCarColumn(ws, m_TargetCar)
    testedCol = FindCarColumn(ws, m_TestedCar)

    If targetCol = 0 Or testedCol = 0 Then
        MsgBox "Error: Could not find data columns for selected cars." & vbCrLf & vbCrLf & _
               "Target: " & m_TargetCar & " (Column: " & targetCol & ")" & vbCrLf & _
               "Tested: " & m_TestedCar & " (Column: " & testedCol & ")", _
               vbCritical, "Error"
        GetSelectedCarColumns = Array(0, 0)
        Exit Function
    End If

    result(0) = targetCol
    result(1) = testedCol
    GetSelectedCarColumns = result
End Function

' ===================================================================
' PRIVATE HELPER FUNCTIONS
' ===================================================================

' GetAvailableCarNames
' Scans row 2 of the worksheet from CAR_DATA_START_COL onwards.
' Returns a unique array of car names (skips empty cells and column
' headers that contain "Status", "P1", "P2", or "P3").
Private Function GetAvailableCarNames(ws As Worksheet) As String()
    Dim carNames() As String
    Dim col As Integer
    Dim carName As String
    Dim count As Integer
    Dim lastCol As Integer

    ReDim carNames(0)
    count = 0

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    col = CAR_DATA_START_COL
    Do While col <= lastCol
        carName = Trim(ws.Cells(2, col).Value)

        If carName <> "" And _
           InStr(1, carName, "Status", vbTextCompare) = 0 And _
           InStr(1, carName, "P1", vbTextCompare) = 0 And _
           InStr(1, carName, "P2", vbTextCompare) = 0 And _
           InStr(1, carName, "P3", vbTextCompare) = 0 Then

            ' Avoid duplicates
            Dim alreadyAdded As Boolean
            Dim i As Integer
            alreadyAdded = False
            For i = 0 To count - 1
                If carNames(i) = carName Then
                    alreadyAdded = True
                    Exit For
                End If
            Next i

            If Not alreadyAdded Then
                ReDim Preserve carNames(count)
                carNames(count) = carName
                count = count + 1
            End If
        End If

        col = col + 1
    Loop

    ' Return empty array if no cars found
    If count = 0 Then
        GetAvailableCarNames = Split("", ",")  ' Empty array
        Exit Function
    End If

    GetAvailableCarNames = carNames
End Function

' IsCarNameValid
' Returns True if carName exists (case-sensitive) in the carNames array.
Private Function IsCarNameValid(carName As String, carNames() As String) As Boolean
    Dim i As Integer

    IsCarNameValid = False

    For i = LBound(carNames) To UBound(carNames)
        If Trim(carNames(i)) = Trim(carName) Then
            IsCarNameValid = True
            Exit Function
        End If
    Next i
End Function

' FindCarColumn
' Returns the column number whose row-2 cell matches carName.
' Returns 0 if not found.
Private Function FindCarColumn(ws As Worksheet, carName As String) As Integer
    Dim col As Integer
    Dim cellValue As String
    Dim lastCol As Integer

    FindCarColumn = 0

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    col = CAR_DATA_START_COL
    Do While col <= lastCol
        cellValue = Trim(ws.Cells(2, col).Value)
        If cellValue = Trim(carName) Then
            FindCarColumn = col
            Exit Function
        End If
        col = col + 1
    Loop
End Function
