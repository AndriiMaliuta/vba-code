VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSection As String
Private pVersion As String
Private pChangeDate As Date
Private pAuthor As String
Private pDescription As String

Property Get Section() As String
    Section = Trim(pSection)
End Property

Property Let Section(Section As String)
    pSection = fixString(Section)
End Property

Property Get version() As String
    version = Trim(pVersion)
End Property

Property Let version(version As String)
    pVersion = fixString(version)
End Property

Property Get VersionNumber() As String
    If IsNumeric(Trim(pVersion)) Then
        VersionNumber = CDbl(Trim(pVersion))
    ElseIf IsNumeric(Replace(Trim(pVersion), ".", ",")) Then
        VersionNumber = CDbl(Replace(Trim(pVersion), ".", ","))
    Else
'        MsgBox "Value: " + Trim(pVersion)
        VersionNumber = -1
    End If
End Property

Property Get ChangeDate() As Date
    ChangeDate = pChangeDate
End Property

Property Let ChangeDate(ChangeDate As Date)
    pChangeDate = ChangeDate
End Property

Property Get Author() As String
    Author = Trim(pAuthor)
End Property

Property Let Author(Author As String)
    pAuthor = fixString(Author)
End Property

Property Get description() As String
    description = Trim(pDescription)
End Property

Property Let description(description As String)
    pDescription = fixString(description)
End Property

