Attribute VB_Name = "YellowDot"
Sub CopyYellowDotAsYellow()
Attribute CopyYellowDotAsYellow.VB_ProcData.VB_Invoke_Func = "y\n14"
' Change U4 to your yellow dot cell, select the target cell first
Dim targetCell As Range
Set targetCell = Selection
targetCell.Value = Range("U4").Value
targetCell.Font.Color = RGB(227, 225, 0) ' Yellow
targetCell.HorizontalAlignment = xlLeft ' Always left-align
End Sub
