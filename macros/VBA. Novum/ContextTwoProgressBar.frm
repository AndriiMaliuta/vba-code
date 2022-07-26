VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContextTwoProgressBar 
   Caption         =   "Now processing..."
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   OleObjectBlob   =   "ContextTwoProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContextTwoProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Option Explicit

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Sub updateOverall(done As Long, overall As Long)
    Dim bar_width As Long
    Dim index As Long
    Dim var As Long
    
    bar_width = OverallBarFrame.Width
    index = bar_width / 100
    
    OverallCompletedBar.Width = (bar_width * done) / overall
    OverallToDoBar.Left = OverallCompletedBar.Width
    
    var = bar_width - OverallCompletedBar.Width
    If var < 0 Then
        var = 0
    End If
    
    OverallToDoBar.Width = var
    
    OverallBarFrame.Repaint
    Repaint
    
    '100 * mltpl * done / overall
End Sub

'@author    Dmitry Romenskiy
'@date      12 Oct 2015
Sub updateCurrent(done As Long, overall As Long)
    Dim bar_width As Long
    Dim index As Long
    Dim var As Long
    
    bar_width = CurrentBarFrame.Width
    index = bar_width / 100
    
    CurrentCompletedBar.Width = (bar_width * done) / overall
    CurrentToDoBar.Left = CurrentCompletedBar.Width
    
    var = bar_width - CurrentCompletedBar.Width
    If var < 0 Then
        var = 0
    End If
    
    CurrentToDoBar.Width = var
    
    CurrentBarFrame.Repaint
    Repaint
    
    '100 * mltpl * done / overall
End Sub
