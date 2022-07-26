VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Option Explicit

Private pSection As String
Private pChangeDate As Date
Private pAuthor As String
Private pDescription As String

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Get Section() As String
    Section = Trim(pSection)
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Let Section(Section As String)
    Dim value As String
    
    value = fixString(Section)
    pSection = value
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Get ChangeDate() As Date
    ChangeDate = pChangeDate
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Let ChangeDate(ChangeDate As Date)
    pChangeDate = ChangeDate
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Get Author() As String
    Author = Trim(pAuthor)
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Let Author(Author As String)
    Dim value As String
    
    value = fixString(Author)
    pAuthor = value
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Get Description() As String
    Description = Trim(pDescription)
End Property

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Property Let Description(Description As String)
    Dim value As String
    
    value = fixString(Description)
    pDescription = value
End Property

