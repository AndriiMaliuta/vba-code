Attribute VB_Name = "ModuleFR"
Sub FindReplacePrefix()

Set MyRange = ActiveDocument.Content

MyRange.Find.Execute FindText:="[Novum]", _
    ReplaceWith:="", Replace:=wdReplaceAll

End Sub

