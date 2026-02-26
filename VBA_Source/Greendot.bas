Attribute VB_Name = "Greendot"
Sub CopyGreenDotAsGreen()
Attribute CopyGreenDotAsGreen.VB_ProcData.VB_Invoke_Func = "g\n14"
' Change U3 to your green dot cell, select the target cell first
Dim targetCell As Range
Set targetCell = Selection
targetCell.Value = Range("U3").Value
targetCell.Font.Color = RGB(0, 176, 80) ' Dark green
targetCell.HorizontalAlignment = xlLeft ' Always left-align
End Sub
