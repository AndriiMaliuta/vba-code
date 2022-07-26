Attribute VB_Name = "Utils"
'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Option Explicit

Private lastStickIndex As Long
Private nextStickIndex As Long

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

' 1|06.04.2018|Anatolii Rostovtsev|First version
            If InStr(cellEntry, "|") > 0 Then
                entries.Add sectionName, parseVersion(cellEntry), parseDate(cellEntry), parseAuthor(cellEntry), parseDescription(cellEntry)
            Else
                entries.Add sectionName, "", vbNull, "", cellEntry
            End If
            
next_i:
        Next i

next_cell:
        index = index + 1
    Next currentCell
    entries.SortByDate
    entries.Sort
    Set extractEntriesFromTable = entries
End Function

Private Function parseVersion(expression As String) As String
    lastStickIndex = InStr(expression, "|")
    parseVersion = Trim(Mid(expression, 1, lastStickIndex - 1))
End Function



Private Function parseDate(expression As String) As Date
    Dim strDate As String
    Dim dateDate As Date
    Dim formattedDateString As String
            
    strDate = getNextString(expression)
    
    If Len(strDate) < 10 Then
        Err.Raise vbObjectError + 1, "Utils::parseDate", "History entry is too short. Failed to parse date."
    End If
    
        ' convert 28.03.2018 to yyyy-MM-dd
    formattedDateString = Right(strDate, 4) & "-" & Mid(strDate, 4, 2) & "-" & Left(strDate, 2)
    On Error GoTo DateParserErrorHandler:
    dateDate = DateValue(formattedDateString)
    
    parseDate = dateDate
    Exit Function
    
DateParserErrorHandler:
    MsgBox "Invalid date format occured: [" & strDate & "]. Setting the date now to Jan 1 2020. Please correct the date in source BASS Design Item"
    dateDate = DateValue("2020-01-01")
    Resume Next
End Function

Private Function parseAuthor(expression As String) As String
    parseAuthor = getNextString(expression)
End Function

Private Function parseDescription(expression As String) As String
    parseDescription = getNextString(expression)
End Function

Function getNextString(sourceString As String) As String
    Dim strValue As String
    
    nextStickIndex = InStr(lastStickIndex + 1, sourceString, "|")
    If nextStickIndex = 0 Then nextStickIndex = Len(sourceString) + 1
    
    strValue = Mid(sourceString, lastStickIndex + 1, nextStickIndex - lastStickIndex - 1)
    lastStickIndex = nextStickIndex
    
    getNextString = Trim(strValue)
End Function

Function fixString(sourceString As String) As String
    fixString = Replace(sourceString, Chr(160), Chr(32))
    fixString = Trim(fixString)
End Function

