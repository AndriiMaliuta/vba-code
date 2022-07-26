Attribute VB_Name = "DesignItemsUtils"
Option Explicit
    
    ' for getDescription functions
Dim nextIndex As Long

Function collectDesignItems() As Collection
    Dim tbl As Table
    Dim hp As Paragraph
    Dim row As row
    Dim items As New Collection
    Dim item As DesignItem
'    Dim Version As String

    For Each tbl In ActiveDocument.Tables
        If tbl.Columns.Count = 2 Then
            Set hp = tbl.Range.GoTo(what:=wdGoToHeading, which:=wdGoToPrevious, Count:=1).Paragraphs(1)
            If hp.Style = "Heading 2" And LCase(trimString(tbl.Rows(1).Cells(2).Range.text)) = "history of changes new" Then
                For Each row In tbl.Rows
                    If Not row.IsFirst Then
                        Set item = New DesignItem
                        item.Name = trimString(row.Cells(1).Range.text)
                        item.History = row.Cells(2).Range.text
                        items.Add item
                    End If
                Next row
                Exit For
            End If
        End If
    Next tbl
    
    Set collectDesignItems = items
End Function

Function parseDesignItems(designItems As Collection) As logEntries
    Dim entries As New logEntries
    Dim entrs As logEntries
    Dim entry As LogEntry
    Dim DesignItem As DesignItem
    
    For Each DesignItem In designItems
        Set entrs = parseDesignItemText(DesignItem.Name, DesignItem.History)
        For Each entry In entrs.logEntries
            entries.AddEntry entry
        Next entry
        Set entrs = Nothing
    Next DesignItem
    
    Set parseDesignItems = entries
End Function


Private Function parseDesignItemText(Section As String, text As String) As logEntries
    Dim entrs As New logEntries
    Dim di As DesignItem
    
    text = Trim(Replace(text, Chr(7), ""))
    
    Dim lines() As String
    Dim line As Variant
    Dim lineStr As String
    Dim dateStr As String
    Dim diDate As Date
    Dim curDate As Date
    Dim description As String
    Dim hasDate As Boolean
    Dim version As String
    
    lines = Split(text, Chr(13))
    
    description = ""
    diDate = vbNull
    curDate = vbNull
    hasDate = False
    
    For Each line In lines
        If Len(line) > 0 Then
            lineStr = CStr(line)
            dateStr = getDate(lineStr)
            
            On Error Resume Next
            diDate = DateValue(dateStr)
            If Err.Number <> 0 Then
                diDate = vbNull
                Err.Clear
            End If
                
            If diDate <> vbNull Then
                If hasDate And Len(description) > 0 Then
                    entrs.Add Section, getVersion(description), curDate, getAuthor(description), getDescription(description)
                End If
                
                curDate = diDate
                hasDate = True
                description = Replace(Replace(lineStr, Chr(10), ""), Chr(13), "")
            Else
                If hasDate Then
                    description = description + Chr(13) + Replace(Replace(lineStr, Chr(10), ""), Chr(13), "")
                End If
            End If
        End If
    Next line
    
    If hasDate And Len(description) > 0 Then
        entrs.Add Section, getVersion(description), curDate, getAuthor(description), getDescription(description)
    End If
    
    Set parseDesignItemText = entrs
End Function

Function getVersion(sourceString As String) As String
    Dim strValue As String
    Dim firstIndex As Long
    
    firstIndex = InStr(sourceString, "|")
    If firstIndex > 0 Then
        strValue = Left(sourceString, firstIndex - 1)
    Else
        strValue = ""
    End If
    
    getVersion = Trim(strValue)
End Function

Function getDate(sourceString As String) As String
    Dim strValue As String
    Dim firstIndex As Long
    
    firstIndex = InStr(sourceString, "|")
    If firstIndex > 0 Then
        firstIndex = firstIndex + 1
        nextIndex = InStr(firstIndex, sourceString, "|")
        strValue = Mid(sourceString, firstIndex, nextIndex - firstIndex)
    Else
        strValue = ""
    End If
    
    getDate = Trim(strValue)
End Function

Function getAuthor(sourceString As String) As String
    Dim strValue As String
    Dim firstIndex As Long
    
    strValue = ""
    firstIndex = InStr(sourceString, "|")
    If firstIndex > 0 Then
        firstIndex = firstIndex + 1
        nextIndex = InStr(firstIndex, sourceString, "|")
    
        If nextIndex > 0 Then
            firstIndex = nextIndex + 1
            nextIndex = InStr(firstIndex, sourceString, "|")
            strValue = Mid(sourceString, firstIndex, nextIndex - firstIndex)
        End If
    End If
    
    getAuthor = Trim(strValue)
End Function

' should be used only after getAuthor function call, so nextIndex is already initialized
Function getDescription(sourceString As String) As String
    Dim strValue As String
    
    If nextIndex > 0 Then
        strValue = Mid(sourceString, nextIndex + 1, Len(sourceString) - nextIndex)
    Else
        strValue = ""
    End If
    
    getDescription = Trim(strValue)
End Function

Function trimString(sourceString As String) As String
    trimString = Replace(Replace(Replace(sourceString, Chr(10), ""), Chr(13), ""), Chr(7), "")
End Function



