Attribute VB_Name = "FIndReplace"
Sub ReplacePrefixes()

Dim mykeywords
mykeywords = Array("\[FDS.CRM.*\]")
Dim myword As Integer

For myword = LBound(mykeywords) To UBound(mykeywords)

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    'Selection.Find.Replacement.Style = ActiveDocument.Styles("HEADING TO SET")
    
    With Selection.Find
        .Text = mykeywords(myword)
        '.Style = "Heading 1"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll

Next

End Sub
