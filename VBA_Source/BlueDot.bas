Attribute VB_Name = "BlueDot"
Sub CopyBlueDotAsBlue()
Attribute CopyBlueDotAsBlue.VB_ProcData.VB_Invoke_Func = "b\n14"
' Change U2 to your Blue dot cell, select the target cell first
Dim targetCell As Range
Set targetCell = Selection
targetCell.Value = Range("U2").Value
targetCell.Font.Color = RGB(153, 251, 251) ' Blue
targetCell.Font.Size = 14                    ' Larger size for blue dot
targetCell.HorizontalAlignment = xlCenter    ' Blue dot is centred (intentional difference from other dots)
End Sub
