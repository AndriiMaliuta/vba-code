Attribute VB_Name = "Utils"
'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Option Explicit

Private authorLastCharIndex As Long

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Function extractEntriesFromTable(Table As Table, ChangeDataColIndex As Long, SectioncolIndex As Long) As LogEntries
    Dim currentCell As Cell
    Dim index As Long
    index = 1
    Dim entries As New LogEntries
    
    Dim cellContent As String
    Dim cellEntries() As String
    
    Dim sectionName As String
    
    For Each currentCell In Table.Columns(ChangeDataColIndex).Cells
        If index = 1 Then GoTo next_cell
        
        sectionName = Table.Cell(currentCell.RowIndex, SectioncolIndex).Range.Text
        sectionName = Replace(sectionName, Chr(7), "")                      'Removing the special character at the end of table cell content
        sectionName = Replace(sectionName, Chr(160), Chr(32))               'Replacing non-breaking spaces with regular ones
        sectionName = Replace(sectionName, Chr(13), "")
        sectionName = Trim(sectionName)                                     'Removing leading and trailing spaces
        
        cellContent = currentCell.Range.Text
        cellContent = Replace(cellContent, Chr(7), "")                      'Removing the special character at the end of table cell content
        cellContent = Replace(cellContent, Chr(160), Chr(32))               'Replacing non-breaking spaces with regular ones
        cellContent = Trim(cellContent)                                     'Removing leading and trailing spaces
        cellEntries = Split(cellContent, Chr(13))                           'Splitting cell content by carriage return character
        
        Dim i As Long
        Dim cellEntry As String
        
        For i = 0 To UBound(cellEntries)
            cellEntry = cellEntries(i)
            cellEntry = Trim(cellEntry)
            
            If cellEntry = "" Then GoTo next_i
            
            entries.Add sectionName, _
                        parseDate(cellEntry), _
                        parseAuthor(cellEntry), _
                        parseDescription(cellEntry)
            
next_i:
        Next i

next_cell:
        index = index + 1
    Next currentCell
    entries.Sort
    Set extractEntriesFromTable = entries
End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Function parseDate(Expression As String) As Date
    Dim strDate As String
    Dim dateDate As Date
    
    If Len(Expression) < 10 Then
        Err.Raise vbObjectError + 1, "Utils::parseDate", "History entry is too short. Failed to parse date."
    End If
    
    strDate = Left(Expression, 10)
    
    On Error GoTo DateParserErrorHandler:
    dateDate = DateValue(strDate)
    
    parseDate = dateDate
    Exit Function
    
DateParserErrorHandler:
    MsgBox "Invalid date format occured: [" & strDate & "]. Setting the date now to Jan 1 2020. Please correct the date in source BASS Design Item"
    dateDate = DateValue("2020-01-01")
    Resume Next
End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Function parseAuthor(Expression As String) As String
    Dim strAuthor As String
    
    Dim firstStickIndex As Long
    Dim secondStickIndex As Long
    
    firstStickIndex = InStr(Expression, "|")
    secondStickIndex = InStr(firstStickIndex + 1, Expression, "|")
    authorLastCharIndex = secondStickIndex
    strAuthor = Mid(Expression, firstStickIndex + 1, secondStickIndex - firstStickIndex - 1)
    
'    Dim firstBraceIndex As Long
'    Dim secondBraceIndex As Long
'
'    firstBraceIndex = InStr(Expression, "(")
'    secondBraceIndex = InStr(Expression, ")")
'    authorLastCharIndex = secondBraceIndex
'
'    strAuthor = Mid(Expression, firstBraceIndex + 1, secondBraceIndex - firstBraceIndex - 1)
    strAuthor = Trim(strAuthor)
    
    parseAuthor = strAuthor

End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Function parseDescription(Expression As String) As String
    Dim strDescription As String
    Dim thirdStickIndex As Long
    thirdStickIndex = InStr(authorLastCharIndex + 1, Expression, "|")
    
    If thirdStickIndex = 0 Then thirdStickIndex = Len(Expression) + 1
    
    strDescription = Mid(Expression, authorLastCharIndex + 1, thirdStickIndex - authorLastCharIndex - 1)
    strDescription = Trim(strDescription)
    
    parseDescription = strDescription

End Function

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Function fixString(sourceString As String) As String
    fixString = Replace(sourceString, Chr(160), Chr(32))
    fixString = Trim(fixString)
End Function


