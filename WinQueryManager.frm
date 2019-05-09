VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinQueryManager 
   Caption         =   "Query Manager"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   OleObjectBlob   =   "WinQueryManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinQueryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Terminated()

Private Sub UserForm_Activate()
    DrawPage
    IssueWarning
End Sub

Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub
Private Sub DrawPage()
    Me.ListBox1.list = DataAccess.GetQueries
End Sub
Private Sub IssueWarning()
    Dim boxResult As VbMsgBoxResult
    boxResult = MsgBox("STOP" & vbCrLf & vbCrLf & "Changing data listed here can result in fatal errors throughout the application, if not performed properly." & vbCrLf & vbCrLf & _
                        "It is advised that if you are not meant to be here, click cancel now.", vbOKCancel, "!!!!WARNING!!!!")
    
    If boxResult = vbCancel Then
        Me.Hide
        UserForm_Terminate
    End If
End Sub

