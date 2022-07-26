Attribute VB_Name = "ModuleFR"
Sub FindReplacePrefix()

Set MyRange = ActiveDocument.Content

With MyRange.Find
      .text = "[POS]"
      .Replacement.text = ""
      .MatchWildcards = False
      .Forward = True
      .Execute Replace:=wdReplaceAll
End With

End Sub


Sub FindReplacePrefix2()

Set MyRange = ActiveDocument.Content

With MyRange.Find
      .text = "[SSP]"
      .Replacement.text = ""
      .MatchWildcards = False
      .Forward = True
      .Execute Replace:=wdReplaceAll
End With

End Sub

Sub ChangeHeading1()

Dim mykeywords
mykeywords = Array("Version History", "Glossary of Terms", "Contents", "Related Documents")
Dim myword As Integer

For myword = LBound(mykeywords) To UBound(mykeywords)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 1 No Numbers")
    
    With Selection.Find
        .text = mykeywords(myword)
        .Style = "Heading 1"
        .Replacement.text = mykeywords(myword)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll

Next

End Sub


