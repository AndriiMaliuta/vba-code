Attribute VB_Name = "ModuleFR"
' Created by Andrii Maliuta on Nov 19, 2019


Sub FindReplacePrefix()

    Dim rngStory As Word.Range
    
      Dim lngJunk As Long
    
      lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    
      For Each rngStory In ActiveDocument.StoryRanges
    
        Do
    
          With rngStory.Find
    
            .text = "[SSP]"
            .Replacement.text = ""
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
    
          End With
    
          Set rngStory = rngStory.NextStoryRange
    
        Loop Until rngStory Is Nothing
    
      Next
  
End Sub


Sub FindReplacePrefix2()

    Dim rngStory As Word.Range
    
      Dim lngJunk As Long
    
      lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    
      For Each rngStory In ActiveDocument.StoryRanges
    
        Do
    
          With rngStory.Find
    
            .text = "[POS]"
            .Replacement.text = ""
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
    
          End With
    
          Set rngStory = rngStory.NextStoryRange
    
        Loop Until rngStory Is Nothing
    
      Next

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


