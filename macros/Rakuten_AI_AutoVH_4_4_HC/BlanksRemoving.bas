Attribute VB_Name = "BlanksRemoving"
' Andrii Maliuta
' July 24, 2018
' Macro to remove empty paragraphs before tables

Sub RemoveBlanks()

Dim MyRange As Range, myTable As Table

For Each myTable In ActiveDocument.Tables

If myTable.Style <> "NC-2" Then

    Set MyRange = myTable.Range
    
    MyRange.Collapse wdCollapseStart
    MyRange.Move wdParagraph, -1
     'if paragraph before table empty, delete it
     
    If MyRange.Paragraphs(1).Range.text = vbCr Then
        MyRange.Paragraphs(1).Range.Delete
    End If

End If

Next myTable

End Sub


