Attribute VB_Name = "RedDot"
Sub CopyRedDotAsRed()
Attribute CopyRedDotAsRed.VB_ProcData.VB_Invoke_Func = "r\n14"
' Change U5 to your red dot cell, select the target cell first
Dim targetCell As Range
Set targetCell = Selection
targetCell.Value = Range("U5").Value
targetCell.Font.Color = RGB(255, 0, 0) ' Dark red
targetCell.HorizontalAlignment = xlLeft ' Always left-align
End Sub
