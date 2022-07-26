Sub SplitCells()
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each Rng In WorkRng
    lLFs = VBA.Len(Rng) - VBA.Len(VBA.Replace(Rng, vbLf, ""))
    If lLFs > 0 Then
        Rng.Offset(1, 0).Resize(lLFs).Insert shift:=xlShiftDown
        Rng.Resize(lLFs + 1).Value = Application.WorksheetFunction.Transpose(VBA.Split(Rng, vbLf))
    End If
Next
End Sub