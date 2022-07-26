Attribute VB_Name = "ModuleFR"
Sub ChangeVH()

Dim mykeywords
mykeywords = Array("Version History")
Dim myword As Integer

For myword = LBound(mykeywords) To UBound(mykeywords)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Table Caption")
    
    With Selection.Find
        .text = mykeywords(myword)
        .Style = "Table Title Large"
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

Sub ChangeDistrGlos()

Dim mykeywords
mykeywords = Array("Glossary of Terms", "Distributions List", "Document References")
Dim myword As Integer

For myword = LBound(mykeywords) To UBound(mykeywords)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Table Caption")
    
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

