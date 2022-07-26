Attribute VB_Name = "FormatModule"
'@author    Dmitry Romenskiy
'@date      12 Oct 2015
'@author    Gennadii Berezin
'@date      11 Dec 2016
Option Explicit

Private entries As LogEntries

Private templateVersionHistoryTable As Table

Private sourceVersionHistoryTable As Table
Private sourceVersionHistoryHeader As Paragraph

Private sourceDocumentHistoryTable As Table
Private sourceDocumentHistoryHeader As Paragraph

Private sourceRelatedDocumentsTable As Table
Private sourceRelatedDocumentsHeader As Paragraph

Private historyGenerated As Boolean

Function need2runMacro() As Boolean
   Dim result As Boolean
   On Error GoTo handleErrors
    result = ActiveDocument.Variables("need2runMacro")
exitHere:
    On Error Resume Next
    need2runMacro = result
    If ActiveDocument.Variables("need2runMacro") <> False Then
        ActiveDocument.Variables("need2runMacro") = False
    End If
    Exit Function
handleErrors:
    If Err.Number = 5825 Then '5825 - Object deleted
        result = True
        ActiveDocument.Variables.Add Name:="need2runMacro", Value:=False
    End If
    Resume exitHere
End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
'
'19 Oct 2015 - Dmitry Romenskiy - Added functionality to remove all links from the documen
'23 Nov 2015 - Dmitry Romenskiy - Added setting variables to default values before start

Sub FDSFormatter()
'    If need2runMacro = False Then
'        Exit Sub
'    End If
    
    Set entries = Nothing
    
    Set templateVersionHistoryTable = Nothing
    Set sourceVersionHistoryTable = Nothing
    
    Set sourceRelatedDocumentsTable = Nothing
    Set sourceRelatedDocumentsHeader = Nothing
    
    Set sourceDocumentHistoryTable = Nothing
    Set sourceDocumentHistoryHeader = Nothing
    
    historyGenerated = False

    Application.ScreenUpdating = False
    ContextTwoProgressBar.Show
    
'    While ActiveDocument.Hyperlinks.Count > 0
'        ActiveDocument.Hyperlinks(1).Delete
'    Wend
    
    collectLogEntries
    ContextTwoProgressBar.updateOverall 1, 3
    
    formatTables
    ContextTwoProgressBar.updateOverall 2, 3
    
    'formatParagraphs
    'ContextTwoProgressBar.updateOverall 3, 4
    
    formatPictures
    ContextTwoProgressBar.updateOverall 3, 3
    
'    ActiveDocument.TablesOfContents(1).Update
'    ActiveDocument.TablesOfContents(1).Range.Font.Name = "Calibri"
    Unload ContextTwoProgressBar
    Application.ScreenUpdating = True
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Sub collectLogEntries()
    ContextTwoProgressBar.OperationName.Caption = "Collecting Version History..."
    
    ' ~==================== History of changes
    Dim currentSourceVersionHistoryTableCandidate As Table
    Dim currentSourceVersionHistoryHeaderCandidate As Paragraph

    For Each currentSourceVersionHistoryTableCandidate In ActiveDocument.Tables
    'currentSourceVersionHistoryTableCandidate.Select
        Set currentSourceVersionHistoryHeaderCandidate = currentSourceVersionHistoryTableCandidate _
            .Range.GoTo( _
                what:=wdGoToHeading, _
                which:=wdGoToPrevious, _
                Count:=1). _
                Paragraphs(1)
                
                'currentSourceVersionHistoryHeaderCandidate.Range.Select
                
        If LCase(Trim(Replace(currentSourceVersionHistoryHeaderCandidate.Range.Text, Chr(13), Chr(32)))) = "version history" Then
            Set sourceVersionHistoryTable = currentSourceVersionHistoryTableCandidate
            Set sourceVersionHistoryHeader = currentSourceVersionHistoryHeaderCandidate
            sourceVersionHistoryHeader.Range.Style = "Table Title Large"
            sourceVersionHistoryHeader.Range.Font.Name = "Calibri"
            sourceVersionHistoryTable.Range.Font.Name = "Calibri"
            sourceVersionHistoryTable.Range.Font.size = 10
            Exit For
        End If
        
    Next currentSourceVersionHistoryTableCandidate
    
    If sourceVersionHistoryTable Is Nothing Then Exit Sub
  
    Set entries = Utils.extractEntriesFromTable(sourceVersionHistoryTable, 2, 1)
    currentSourceVersionHistoryTableCandidate.Delete
    currentSourceVersionHistoryHeaderCandidate.Range.Delete
    
'If True Then Exit Sub
   
    
    ' ~==================== History of the document
    Dim currentSourceDocumentHistoryTableCandidate As Table
    Dim currentSourceDocumentHistoryHeaderCandidate As Paragraph
    
    For Each currentSourceDocumentHistoryTableCandidate In ActiveDocument.Tables
        Set currentSourceDocumentHistoryHeaderCandidate = currentSourceDocumentHistoryTableCandidate _
            .Range.GoTo( _
                    what:=wdGoToLine, _
                    which:=wdGoToPrevious, _
                    Count:=1). _
                    Paragraphs(1)
        
        If currentSourceDocumentHistoryHeaderCandidate.OutlineLevel = wdOutlineLevel1 _
                And LCase(Trim(Replace(currentSourceDocumentHistoryHeaderCandidate.Range.Text, Chr(13), Chr(32)))) = "history of the document" Then
            Set sourceDocumentHistoryTable = currentSourceDocumentHistoryTableCandidate
            Set sourceDocumentHistoryHeader = currentSourceDocumentHistoryHeaderCandidate
            sourceDocumentHistoryHeader.Range.Style = "Table Title Large"
            sourceDocumentHistoryHeader.Range.Font.Name = "Calibri"
            sourceDocumentHistoryTable.Range.Font.Name = "Calibri"
            sourceDocumentHistoryTable.Range.Font.size = 10
            'sourceDocumentHistoryTable.Style = "Table Grid"
            Exit For
        End If
        
    Next currentSourceDocumentHistoryTableCandidate
    
    ' ~==================== Related documents
    Dim currentSourceRelatedDocumentsTableCandidate As Table
    Dim currentSourceRelatedDocumentsHeaderCandidate As Paragraph
    
    For Each currentSourceRelatedDocumentsTableCandidate In ActiveDocument.Tables
        Set currentSourceRelatedDocumentsHeaderCandidate = currentSourceRelatedDocumentsTableCandidate _
            .Range.GoTo( _
                    what:=wdGoToLine, _
                    which:=wdGoToPrevious, _
                    Count:=1). _
                    Paragraphs(1)
        
        If currentSourceRelatedDocumentsHeaderCandidate.OutlineLevel = wdOutlineLevel1 _
                And LCase(Trim(Replace(currentSourceRelatedDocumentsHeaderCandidate.Range.Text, Chr(13), Chr(32)))) = "related documents" Then
            Set sourceRelatedDocumentsTable = currentSourceRelatedDocumentsTableCandidate
            Set sourceRelatedDocumentsHeader = currentSourceRelatedDocumentsHeaderCandidate
            sourceRelatedDocumentsHeader.Range.Style = "Table Title Large"
            sourceRelatedDocumentsHeader.Range.Font.Name = "Calibri"
            sourceRelatedDocumentsTable.Range.Font.Name = "Calibri"
            sourceRelatedDocumentsTable.Range.Font.size = 10
            'sourceRelatedDocumentsTable.Style = "Table Grid"
            Exit For
        End If
        
    Next currentSourceRelatedDocumentsTableCandidate
    
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
'
'@date      26 Oct 2015 - Dmitry Romenskiy - fixed dropping list formatting within table cells.
Private Sub formatTables()
    ContextTwoProgressBar.OperationName.Caption = "Formatting tables..."

    Dim currentTable As Table
    Dim index As Long
    index = 1

    For Each currentTable In ActiveDocument.Tables
        If index = 1 Then GoTo next_table               'Ignoring the table at the title page
        
        If index = 2 Then
            If Not sourceDocumentHistoryTable Is Nothing Or Not sourceRelatedDocumentsTable Is Nothing Then
                currentTable.Select
                Selection.MoveUp Unit:=wdLine, Count:=1
                Selection.TypeParagraph
                Selection.MoveUp Unit:=wdLine, Count:=1
                
                If Not sourceRelatedDocumentsTable Is Nothing Then
                    sourceRelatedDocumentsHeader.Range.Cut
                    Selection.Paste
                    Selection.TypeParagraph
                    sourceRelatedDocumentsTable.Range.Cut
                    Selection.Paste
                End If
                
                If Not sourceDocumentHistoryTable Is Nothing Then
                    Selection.TypeParagraph
                    sourceDocumentHistoryHeader.Range.Cut
                    Selection.Paste
                    Selection.TypeParagraph
                    sourceDocumentHistoryTable.Range.Cut
                    Selection.Paste
                End If
            End If
        End If
        
'        currentTable.Range.Font.Name = "Calibri"
'        currentTable.Range.Font.size = 10
        'currentTable.Style = "Table Grid"
        
        ' remove line after row in table
'        Dim iRow, iCol As Integer
'        Dim Str, endSymbol As String
'        If index > 2 Then
'            With currentTable
'                For iRow = 1 To .Rows.Count
'                    For iCol = 1 To .Columns.Count
'                        On Error Resume Next
'                        Str = .Cell(iRow, iCol).Range.Text
'                        If Asc(Left(Right(Str, 3), 1)) = 11 Then
'                            endSymbol = Right(Str, 1)
'                            Str = Left(Str, Len(Str) - 3)
'                            .Cell(iRow, iCol).Range.Text = Str + endSymbol
'                        End If
'                        On Error GoTo 0
'                    Next iCol
'                Next iRow
'            End With
'        End If
        
        If entries Is Nothing Or historyGenerated Then
            GoTo next_table
        Else
            Dim vhHeadingCandidate As Paragraph
            Set vhHeadingCandidate = currentTable _
                .Range.GoTo( _
                    what:=wdGoToLine, _
                    which:=wdGoToPrevious, _
                    Count:=1). _
                    Paragraphs(1)
                    
            If vhHeadingCandidate.Range.Style = "Table Title Large" _
                    And LCase(Trim(Replace(vhHeadingCandidate.Range.Text, Chr(13), Chr(32)))) = "version history" Then
                Set templateVersionHistoryTable = currentTable
                fillInTemplateVersionhistoryTable
            End If
        End If
        
next_table:
        index = index + 1
        
        If index Mod 10 = 0 Then
            DoEvents
        End If
        ContextTwoProgressBar.updateCurrent index, ActiveDocument.Tables.Count
    
    Next currentTable
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Sub fillInTemplateVersionhistoryTable()
    Dim currentHistoryEntry As LogEntry
    Dim row As Long
    row = 2
    Dim dateStr As String
    
    For Each currentHistoryEntry In entries.LogEntries
        If currentHistoryEntry.ChangeDate = vbNull Then
            dateStr = ""
        Else
            dateStr = Format(currentHistoryEntry.ChangeDate, "dd-mm-yyyy")
        End If
    
        With templateVersionHistoryTable
            .Rows.Add
            .Cell(row, 1).Range.Text = currentHistoryEntry.Version & Chr(7)
            .Cell(row, 2).Range.Text = dateStr & Chr(7)
            .Cell(row, 3).Range.Text = currentHistoryEntry.Author & Chr(7)
            .Cell(row, 4).Range.Text = currentHistoryEntry.Section & Chr(7)
            .Cell(row, 5).Range.Text = currentHistoryEntry.Description & Chr(7)
        End With
        row = row + 1
    Next currentHistoryEntry
    templateVersionHistoryTable.Range.Font.TextColor = RGB(0, 0, 0)
    templateVersionHistoryTable.Range.Font.Name = "Calibri"
    templateVersionHistoryTable.Range.Font.size = 10
    setColumnWidth
    historyGenerated = True
End Sub

Private Sub setColumnWidth()
    templateVersionHistoryTable.AutoFitBehavior wdAutoFitFixed
    templateVersionHistoryTable.Columns.Item(1).Width = Application.InchesToPoints(0.5)
    templateVersionHistoryTable.Columns.Item(2).Width = Application.InchesToPoints(0.8)
    templateVersionHistoryTable.Columns.Item(3).Width = Application.InchesToPoints(1.2)
    templateVersionHistoryTable.Columns.Item(4).Width = Application.InchesToPoints(2)
    templateVersionHistoryTable.Columns.Item(5).Width = Application.InchesToPoints(2.15)
    templateVersionHistoryTable.PreferredWidthType = wdPreferredWidthPoints
    templateVersionHistoryTable.Rows.LeftIndent = 0
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Sub formatParagraphs()
    ContextTwoProgressBar.OperationName.Caption = "Formatting paragraphs..."
    Dim index As Long
    Dim tmp_header As String
    Dim j As Integer
    index = 1
    
    Dim currentParagraph As Paragraph
    'For Each currentParagraph In ActiveDocument.Paragraphs
    For j = 1 To ActiveDocument.Paragraphs.Count
    'currentParagraph = ActiveDocument.Paragraphs(j)
        ActiveDocument.Paragraphs(j).Range.Select
       
        If Selection.Information(wdActiveEndPageNumber) = 1 Then
            GoTo endFor
        End If

        ActiveDocument.Paragraphs(j).Range.Font.Name = "Calibri"
        'If (ActiveDocument.Paragraphs(j).Range.ListFormat.ListType = wdListBullet) Then
        If (ActiveDocument.Paragraphs(j).Style = "Scroll List Bullet 1" Or _
            ActiveDocument.Paragraphs(j).Style = "Scroll List Bullet 3") Then
            If Not (ActiveDocument.Paragraphs(j).Range.ListFormat.ListTemplate Is Nothing) Then
                ActiveDocument.Paragraphs(j).Range.ListFormat.ListTemplate.ListLevels(1).Font.Color = wdAuto
                ActiveDocument.Paragraphs(j).Range.ListFormat.ApplyListTemplate ListTemplate:=Word.Application.ListGalleries(wdBulletGallery).ListTemplates(1), _
                continuepreviouslist:=False, applyto:=wdListApplyToSelection, defaultlistbehavior:=wdWord9ListBehavior
            End If
        End If
        'Heading 1
        If ActiveDocument.Paragraphs(j).Style = "Heading 1" Then
            With ActiveDocument.Paragraphs(j).Range
                .Text = formatHeader(ActiveDocument.Paragraphs(j).Range.Text)
                .Style = "Heading 1"
                .Font.size = 14
                .Font.Name = "Calibri"
            End With
                GoTo endFor
        End If
    
        'Heading 2
        If ActiveDocument.Paragraphs(j).Style = "Heading 2" Then
            With ActiveDocument.Paragraphs(j).Range
                .Style = "Heading 2"
                .Font.size = 12
                .Font.Name = "Calibri"
            End With
            GoTo endFor
        End If

        'Heading 3
        If ActiveDocument.Paragraphs(j).Style = "Heading 3" Then
            With ActiveDocument.Paragraphs(j).Range
                .Style = "Heading 4"      'fix bug of ScrollOffice plugin
                .Style = "Heading 3"
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'Heading 4
        If ActiveDocument.Paragraphs(j).Style = "Heading 4" Then
            With ActiveDocument.Paragraphs(j).Range
                .Style = "Heading 4"
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'Heading 5
        If ActiveDocument.Paragraphs(j).Style = "Heading 5" Then
            With ActiveDocument.Paragraphs(j).Range
                .Style = "Heading 5"
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'Normal
        If ActiveDocument.Paragraphs(j).Style = "Normal" And ActiveDocument.Paragraphs(j).Range.Tables.Count = 0 Then
            With ActiveDocument.Paragraphs(j).Range
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'Normal (Web)
        If ActiveDocument.Paragraphs(j).Style = "Normal (Web)" And ActiveDocument.Paragraphs(j).Range.Tables.Count = 0 Then
            With ActiveDocument.Paragraphs(j).Range
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'body
        If ActiveDocument.Paragraphs(j).Style = "body" And ActiveDocument.Paragraphs(j).Range.Tables.Count = 0 Then
            With ActiveDocument.Paragraphs(j).Range
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'default
        If ActiveDocument.Paragraphs(j).Style = "default" And ActiveDocument.Paragraphs(j).Range.Tables.Count = 0 Then
            With ActiveDocument.Paragraphs(j).Range
                .Font.Name = "Calibri"
                '.Font.size = 11
            End With
            GoTo endFor
        End If

        'normal1
        If ActiveDocument.Paragraphs(j).Style = "normal1" And ActiveDocument.Paragraphs(j).Range.Tables.Count = 0 Then
            With ActiveDocument.Paragraphs(j).Range
                .Font.Name = "Calibri"
                .Font.size = 11
            End With
            GoTo endFor
        End If
endFor:
        index = index + 1
        If index Mod 10 = 0 Then
            DoEvents
        End If
        ContextTwoProgressBar.updateCurrent index, ActiveDocument.Paragraphs.Count
    Next j 'currentParagraph
End Sub

'remove spaces between words (<span> tag)
Function formatHeader(Str As String) As String
    Dim testArray() As String
    Dim resultArray() As String
    ReDim resultArray(0)
    Dim element As Variant
    testArray = Split(Str)
    For Each element In testArray
        If (element <> "") Then
            If (resultArray(0) <> "") Then ReDim Preserve resultArray(UBound(resultArray) + 1)
            resultArray(UBound(resultArray)) = element
        End If
    Next
    formatHeader = Join(resultArray, " ")
End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Sub formatPictures()
    ContextTwoProgressBar.OperationName.Caption = "Formatting pictures..."
    
    Dim index As Long
    Dim currentImg As InlineShape
    index = 1
    For Each currentImg In ActiveDocument.InlineShapes
        If currentImg.Range.ListParagraphs.Count > 0 Then
            If currentImg.Range.ListFormat.ListLevelNumber = 1 Then
                currentImg.Width = 6 * 72
            ElseIf currentImg.Range.ListFormat.ListLevelNumber = 2 Then
                currentImg.Width = 5.5 * 72
            ElseIf currentImg.Range.ListFormat.ListLevelNumber = 3 Then
                currentImg.Width = 5 * 72
            ElseIf currentImg.Range.ListFormat.ListLevelNumber = 4 Then
                currentImg.Width = 4.5 * 72
            ElseIf currentImg.Range.ListFormat.ListLevelNumber > 4 Then
                currentImg.Width = (4.5 - ((currentImg.Range.ListFormat.ListLevelNumber - 4) * 0.5)) * 72
            End If
            
        End If
        
        index = index + 1
        If index Mod 10 = 0 Then
            DoEvents
        End If
        ContextTwoProgressBar.updateCurrent index, ActiveDocument.InlineShapes.Count
    Next currentImg
End Sub

Private Sub abc()
    ActiveDocument.Variables("need2runMacro") = True
End Sub


