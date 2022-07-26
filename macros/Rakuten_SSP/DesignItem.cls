VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private diName As String
Private diNumber As String
Private diHistory As String

Property Get Name() As String
    Name = Trim(diName)
End Property

Property Let Name(Name As String)
    diName = Name
End Property

Property Get Number() As String
    Number = Trim(diNumber)
End Property

Property Let Number(Number As String)
    diNumber = Number
End Property

Property Get History() As String
    History = Trim(diHistory)
End Property

Property Let History(History As String)
    diHistory = History
End Property

