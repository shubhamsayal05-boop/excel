Attribute VB_Name = "Clearall"
Sub Clear_Sheet1_Keep_ClearButton()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' 1. Clear all cells
    ws.Cells.Clear

    ' 2. Delete shapes SAFELY (loop backwards, keep shapes tagged KEEP)
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).AlternativeText <> "KEEP" Then
            ws.Shapes(i).Delete
        End If
    Next i

    ' 3. Delete ActiveX controls (if any), keep the Clear_All button
    For i = ws.OLEObjects.Count To 1 Step -1
        If ws.OLEObjects(i).Object.Name <> "Clear_All" Then
            ws.OLEObjects(i).Delete
        End If
    Next i

End Sub
