VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinPreferredEntry 
   Caption         =   "Preferred Name Entry"
   ClientHeight    =   1395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "WinPreferredEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinPreferredEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditMode As Boolean
Public Event Terminated()
Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub

Private Sub UserForm_Activate()
    DrawPage
End Sub
Private Sub ButtonSubmit_Click()
    If EditMode Then
        DataAccess.ChangePreferredName Trim(Me.TextBoxSoldTo), Trim(Me.TextBoxPreferredName)
        Me.Hide
    Else
        If Not DataAccess.AddPreferredName(Trim(Me.TextBoxSoldTo), Trim(Me.TextBoxPreferredName)) Then
            MsgBox "Add entry failed."
        Else
            Me.Hide
        End If
    End If
    DrawPage
End Sub
Private Sub DrawPage()
    If EditMode Then
        Me.Caption = "Edit"
        Me.TextBoxSoldTo.Enabled = False
        Me.TextBoxSoldTo.BackStyle = fmBackStyleTransparent
    Else
        Me.Caption = "Add"
        Me.TextBoxSoldTo.Enabled = True
        Me.TextBoxSoldTo.BackStyle = fmBackStyleOpaque
    End If
End Sub

