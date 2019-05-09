VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinMainNav 
   Caption         =   "iGenerate:Labels"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
   OleObjectBlob   =   "WinMainNav.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinMainNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event SearchViewRequested()
Public Event ManualEntryViewRequested(REQUEST As iGenerateLabels.ENTRY_MODE)
Public Event SettingsViewRequested()
Public Event Terminated()
Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub

Private Sub ButtonLabelsFromFile_Click()
    RaiseEvent ManualEntryViewRequested(LOAD_OUTLOOK_ENTRIES)
End Sub

Private Sub ButtonManual_Click()
    RaiseEvent ManualEntryViewRequested(MANUAL)
End Sub

Private Sub ButtonPictureLoadFromFile_Click()
    ButtonLabelsFromFile_Click
End Sub

Private Sub ButtonPictureManual_Click()
    ButtonManual_Click
End Sub

Private Sub ButtonPictureSearch_Click()
    ButtonSearch_Click
End Sub

Private Sub ButtonPictureSettings_Click()
    ButtonSettings_Click
End Sub

Private Sub ButtonSearch_Click()
    RaiseEvent SearchViewRequested
End Sub

Private Sub ButtonSettings_Click()
    RaiseEvent SettingsViewRequested
End Sub

Private Sub UserForm_Activate()

    DataAccess.SetPicture Me.ButtonPictureSearch.picture, "PATH_SearchIcon"
    DataAccess.SetPicture Me.ButtonPictureManual.picture, "PATH_EditIcon"
    DataAccess.SetPicture Me.ButtonPictureSettings.picture, "PATH_SettingsIcon"
    DataAccess.SetPicture Me.ButtonLabelsFromFile.picture, "PATH_LoadFromFileIcon"
    
End Sub
