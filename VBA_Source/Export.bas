Attribute VB_Name = "Export"
Option Explicit

Sub ExportSelectionVisibleOnly_AsPicture()
    Const HEADER_ROWS As Long = 2   'rows 1..HEADER_ROWS are headers

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim sel0 As Range, tmp As Worksheet
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    Dim nRows As Long, nCols As Long
    Dim prevAlerts As Boolean, prevScreen As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range.", vbExclamation: Exit Sub
    End If
    Set sel0 = Selection

    'Bounding rectangle
    firstRow = sel0.Cells(1).Row: firstCol = sel0.Cells(1).Column
    lastRow = firstRow + sel0.Rows.Count - 1
    lastCol = firstCol + sel0.Columns.Count - 1
    If sel0.Areas.Count > 1 Then
        Dim a As Range
        For Each a In sel0.Areas
            If a.Row < firstRow Then firstRow = a.Row
            If a.Column < firstCol Then firstCol = a.Column
            If a.Row + a.Rows.Count - 1 > lastRow Then lastRow = a.Row + a.Rows.Count - 1
            If a.Column + a.Columns.Count - 1 > lastCol Then lastCol = a.Column + a.Columns.Count - 1
        Next a
    End If

    'Ignore headers
    If firstRow <= HEADER_ROWS Then firstRow = HEADER_ROWS + 1
    If firstRow > lastRow Then
        Dim maxUsed As Long, c As Long, u As Long
        maxUsed = HEADER_ROWS
        For c = firstCol To lastCol
            u = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            If u > maxUsed Then maxUsed = u
        Next c
        If maxUsed <= HEADER_ROWS Then
            MsgBox "No data under headers.", vbInformation: Exit Sub
        End If
        lastRow = maxUsed
    End If

    nRows = lastRow - firstRow + 1
    nCols = lastCol - firstCol + 1

    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    'Add temp sheet BUT keep track of original
    Set tmp = Worksheets.Add(After:=ws)

    'Copy headers + data
    ws.Range(ws.Cells(1, firstCol), ws.Cells(HEADER_ROWS, lastCol)).Copy tmp.Range("A1")
    ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).Copy tmp.Cells(HEADER_ROWS + 1, 1)

    'Match col widths
    Dim c2 As Long
    For c2 = 1 To nCols
        tmp.Columns(c2).ColumnWidth = ws.Columns(firstCol + c2 - 1).ColumnWidth
    Next c2

    'Delete hidden cols
    Dim cSrc As Long
    For cSrc = lastCol To firstCol Step -1
        If ws.Columns(cSrc).Hidden Then
            tmp.Columns(cSrc - firstCol + 1).Delete
        End If
    Next cSrc

    'Delete hidden rows
    Dim rSrc As Long
    For rSrc = lastRow To firstRow Step -1
        If ws.Rows(rSrc).Hidden Then
            tmp.Rows(HEADER_ROWS + (rSrc - firstRow + 1)).Delete
        End If
    Next rSrc

    'Copy visible-only block as picture
    Dim exportBlock As Range
    Set exportBlock = tmp.UsedRange
    exportBlock.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Application.CutCopyMode = False

    'Return to original sheet
    ws.Activate

    MsgBox "Copied only visible headers + data as a picture.", vbInformation

CleanExit:
    If Not tmp Is Nothing Then tmp.Delete
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Exit Sub

CleanFail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
