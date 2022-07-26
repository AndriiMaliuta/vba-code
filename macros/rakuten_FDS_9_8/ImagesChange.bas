Attribute VB_Name = "ImagesChange"
Sub resize()
'
' resize Macro
'
Dim i As Integer

    With ActiveDocument
    
        For i = 1 To .InlineShapes.Count
           
            Set ishp = .InlineShapes(i)
            
                If ishp.Height > InchesToPoints(7) Then
                    ishp.LockAspectRatio = True
                    ishp.Height = InchesToPoints(7)
                End If
                    
        Next i

    End With
End Sub

Sub centerPictures()
  Dim shpIn As InlineShape, shp As shape
  For Each shpIn In ActiveDocument.InlineShapes
    shpIn.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  Next shpIn
  For Each shp In ActiveDocument.Shapes
    shp.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  Next shp
End Sub

Sub deleteAltText()

Dim shape As InlineShape

 For Each shape In ActiveDocument.InlineShapes
    'shape.Select
    shape.AlternativeText = ""
 Next

End Sub
