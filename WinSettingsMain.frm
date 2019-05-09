VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinSettingsMain 
   Caption         =   "Settings"
   ClientHeight    =   2595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "WinSettingsMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinSettingsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PreferedNameViewRequested()
Public Event FileManagerViewRequested()
Public Event QueryManagerViewRequested()
Public Event Terminated()
Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub

Private Sub ButtonContactMe_Click()
    Dim mail As MailItem
    
    Set mail = Outlook.CreateItem(olMailItem)
    
    mail.To = "ContactMe@peterjkingston.com"
    mail.Subject = "Request for Development Support: iGenerate:Labels"
    
    mail.Display
End Sub

Private Sub ButtonFileManager_Click()
    RaiseEvent FileManagerViewRequested
End Sub

Private Sub ButtonPreferredNames_Click()
    RaiseEvent PreferedNameViewRequested
End Sub

Private Sub ButtonQueryManager_Click()
    RaiseEvent QueryManagerViewRequested
End Sub

