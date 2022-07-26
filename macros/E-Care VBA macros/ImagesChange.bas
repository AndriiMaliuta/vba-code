Attribute VB_Name = "ImagesChange"
Sub resize()
'
' resize Macro
'
'
Dim i As Long
With ActiveDocument
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
            .ScaleHeight = 15
            .ScaleWidth = 15

        End With
    Next i
End With
End Sub

Sub centerPictures()
  Dim shpIn As InlineShape, shp As Shape
  For Each shpIn In ActiveDocument.InlineShapes
    shpIn.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  Next shpIn
  For Each shp In ActiveDocument.Shapes
    shp.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  Next shp
End Sub
