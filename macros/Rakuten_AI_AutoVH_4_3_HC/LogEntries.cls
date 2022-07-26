VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private entries As New Collection

Function AddEntry(entry As LogEntry) As Boolean
    entries.Add entry
End Function

Function Add(Section As String, version As String, ChangeDate As Date, Author As String, description As String) As Boolean
    Dim entry As New LogEntry
    
'Debug.Print ""
'Debug.Print "ADD: " & Section & "   " & Version & "   " & ChangeDate & "   " & Author & "   " & description
'Debug.Print ""
    
    entry.Section = Section
    entry.version = version
    entry.ChangeDate = ChangeDate
    entry.Author = Author
    entry.description = description
    
    entries.Add entry
End Function

Property Get logEntries() As Collection
    Set logEntries = entries
End Property


Sub SortByDate()
    Dim sortedArray() As LogEntry
    
    If entries.Count = 0 Then Exit Sub
    
    ReDim sortedArray(1 To entries.Count)
    Dim currentEntry As LogEntry
    Dim size As Long
    size = 0
    Dim inserted As Boolean
    
    For Each currentEntry In entries
        inserted = False
        If size = 0 Then
            Set sortedArray(1) = currentEntry
            inserted = True
            size = size + 1
            GoTo next_entry
        End If
        
        Dim i As Long
        For i = 1 To size
            If currentEntry.ChangeDate < sortedArray(i).ChangeDate Then
                sortedArray = stepArray(sortedArray, i)
                Set sortedArray(i) = currentEntry
                size = size + 1
                inserted = True
                GoTo next_entry
            End If
            
        Next i
        
        If Not inserted Then
            Set sortedArray(size + 1) = currentEntry
            size = size + 1
        End If
        
next_entry:
    
    Next currentEntry
    
    Do While entries.Count >= 1
        entries.Remove (entries.Count)
    Loop
    
    Dim k As Long
    For k = 1 To UBound(sortedArray)
        entries.Add sortedArray(k)
    Next k
    
End Sub


Sub Sort()
    Dim sortedArray() As LogEntry
    
    If entries.Count = 0 Then Exit Sub
    
    ReDim sortedArray(1 To entries.Count)
    Dim currentEntry As LogEntry
    Dim size As Long
    size = 0
    Dim inserted As Boolean
    
    For Each currentEntry In entries
        inserted = False
        If size = 0 Then
            Set sortedArray(1) = currentEntry
            inserted = True
            size = size + 1
            GoTo next_entry
        End If
        
        Dim i As Long
        For i = 1 To size
'            If currentEntry.ChangeDate < sortedArray(i).ChangeDate Then
'                sortedArray = stepArray(sortedArray, i)
'                Set sortedArray(i) = currentEntry
'                size = size + 1
'                inserted = True
'                GoTo next_entry
'            End If
            
            If currentEntry.VersionNumber < sortedArray(i).VersionNumber Then
                sortedArray = stepArray(sortedArray, i)
                Set sortedArray(i) = currentEntry
                size = size + 1
                inserted = True
                GoTo next_entry
            End If
            
            
        Next i
        
        If Not inserted Then
            Set sortedArray(size + 1) = currentEntry
            size = size + 1
        End If
        
next_entry:
    
    Next currentEntry
    
    Do While entries.Count >= 1
        entries.Remove (entries.Count)
    Loop
    
    Dim k As Long
    For k = 1 To UBound(sortedArray)
        entries.Add sortedArray(k)
    Next k
    
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Private Function stepArray(sorted() As LogEntry, startIndex As Long) As LogEntry()
    Dim i As Long
    For i = UBound(sorted) - 1 To 1 Step -1
        If i >= startIndex Then
            Set sorted(i + 1) = sorted(i)
        End If
    Next i
    
    stepArray = sorted
End Function

Sub updateVersions()
    Dim currentVersion As String
    currentVersion = "-9999"
    
    If entries.Count = 0 Then Exit Sub
    
    Dim currentEntry As LogEntry
    Dim index As Long
    index = 1
    
    For Each currentEntry In entries
        If currentVersion <> currentEntry.version Then
            index = 1
            currentVersion = currentEntry.version
        End If
        currentEntry.version = currentEntry.version & "." & FormatNumber(index, 0)
        index = index + 1
    Next currentEntry
End Sub


