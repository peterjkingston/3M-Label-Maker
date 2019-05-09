VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinSplash 
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10110
   OleObjectBlob   =   "WinSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event OnActivate()
Public Event Terminated()
Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub
Private Sub UserForm_Activate()
    DoEvents ''<---Required to paint the splash screen content during this operation
    RaiseEvent OnActivate
End Sub

