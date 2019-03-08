VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinPrintDialogue 
   Caption         =   "Dialogue"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "WinPrintDialogue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinPrintDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonCAQuit_Click()
    Functions.ClearLog
    ButtonQuit_Click
End Sub

Private Sub ButtonCAReturn_Click()
    Functions.ClearLog
    WinLogNav.DrawPage
    
    ButtonReturn_Click
    WinPrintPreview.Hide
End Sub
Private Sub ButtonQuit_Click()
    Me.Hide
    WinPrintPreview.Hide
    WinLogNav.Hide
End Sub
Private Sub ButtonReturn_Click()
    Me.Hide
End Sub

