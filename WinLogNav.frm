VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinLogNav 
   Caption         =   "Log Navigation"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "WinLogNav.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinLogNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Win64 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
               (ByVal hWnd As LongPtr) As LongPtr
#Else
    Private Declare Function SetForegroundWindow Lib "user32" _
               (ByVal hWnd As Long) As Long
#End If
Private Sub ButtonClearAll_Click()
    If DataAccess.LabelCount = 0 Then Exit Sub
    Functions.ClearLog
    DrawPage
End Sub

Private Sub ButtonEdit_Click()
    If DataAccess.LabelCount <> 0 Then
        WinManualEntry.Caption = "Edit"
        WinManualEntry.EditMode = True
        WinManualEntry.Show
        DrawPage
    End If
End Sub

Private Sub ButtonManualEntry_Click()
    WinManualEntry.Caption = "Manual Entry"
    WinManualEntry.EditMode = False
    WinManualEntry.Show
    DrawPage
End Sub

Private Sub ButtonPrintPreview_Click()
    WinPrintPreview.Show
End Sub

Private Sub ButtonRemoveSelection_Click()
    If WinLogNav.ListBox1.ListIndex <> -1 Then
        Functions.RemoveItem (WinLogNav.ListBox1.ListIndex)
        DrawPage
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ButtonEdit_Click
End Sub

Private Sub ListBox1_Click()
    If Me.ListBox1.ListIndex > -1 Then
        Me.ButtonEdit.Enabled = True
        Me.ButtonRemoveSelection.Enabled = True
    End If
End Sub

Private Sub UserForm_Activate()
    DrawPage
    If Me.ListBox1.ListIndex = -1 Then: Me.ButtonEdit.Enabled = False
    If Not Functions.IsQueryValid Then: MsgBox "The dataset is out of date. Please update the table from SAP." & vbCrLf & " //INSTRUCTIONS TO BE ADDED"
    
End Sub

Public Sub DrawPage()
    If DataAccess.DataIsEmpty Then
        WinLogNav.ListBox1.list = Array("EMPTY")
        ButtonPrintPreview.Enabled = False
        Me.ButtonRemoveSelection.Enabled = False
    Else
        WinLogNav.ListBox1.list = DataAccess.labelArray
        ButtonPrintPreview.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Application.Visible = True
    Sheet1.Activate
    SetForegroundWindow Application.hWnd
    Application.Visible = False
    
    DataAccess.AssembleGlobalDict
    DataAccess.AssembleCorrectionDict
End Sub

Private Sub UserForm_Terminate()
    Functions.TurnOn
End Sub
