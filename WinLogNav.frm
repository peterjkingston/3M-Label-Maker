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
Private WithEvents mManualEntry As WinManualEntry
Attribute mManualEntry.VB_VarHelpID = -1

Public Event RequestViewManualEntry(EditMode As Boolean, dataArgs() As String, index As Integer)
Public Event RequestViewPrintPreview()
Public Event Terminated()

Private Sub ButtonClearAll_Click()
    If DataAccess.LabelCount = 0 Then Exit Sub
    DataAccess.ClearLabels
    DrawPage
End Sub

Private Sub ButtonEdit_Click()
    Dim labelArgs(3) As String, column As Integer
    With Me.ListBox1
        For column = 0 To UBound(labelArgs)
            labelArgs(column) = .list(.ListIndex, column)
        Next column
    End With
    RaiseEvent RequestViewManualEntry(True, labelArgs, ListBox1.ListIndex)
    DrawPage
End Sub

Private Sub ButtonManualEntry_Click()
    Dim labelArgs(3) As String, column As Integer
    With Me.ListBox1
        For column = 0 To UBound(labelArgs)
            labelArgs(column) = "" ''
        Next column
    End With
    RaiseEvent RequestViewManualEntry(False, labelArgs, 0)
    DrawPage
End Sub

Private Sub ButtonPrintPreview_Click()
    ''WinPrintPreview.Show
    RaiseEvent RequestViewPrintPreview
    DrawPage
End Sub

Private Sub ButtonRemoveSelection_Click()
    Dim index As Integer
    With Me.ListBox1
        If .ListIndex <> -1 Then
            index = .ListIndex
            DataAccess.RemoveLabel index
            DrawPage
            If index = .ListCount Then: index = UBound(.list)
            .ListIndex = index
        End If
    End With
    DrawPage
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

Private Sub mManualEntry_OnSubmit()
    DrawPage
End Sub

Private Sub mManualEntry_Terminated()
    Set mManualEntry = Nothing
End Sub

Private Sub UserForm_Activate()
    DrawPage
    If Me.ListBox1.ListIndex = -1 Then: Me.ButtonEdit.Enabled = False
    If Not Functions.IsQueryValid Then: MsgBox "The dataset is out of date. Please update the table from SAP." & vbCrLf & " //INSTRUCTIONS TO BE ADDED"
    
End Sub

Private Sub DrawPage()
    If DataAccess.DataIsEmpty Then
        Me.ListBox1.list = Array("EMPTY")
        Me.ButtonEdit.Enabled = False
        Me.ButtonRemoveSelection.Enabled = False
        Me.ButtonPrintPreview.Enabled = False
    Else
        Me.ListBox1.list = DataAccess.labelArray
        Me.ButtonEdit.Enabled = True
        Me.ButtonRemoveSelection.Enabled = True
        Me.ButtonPrintPreview.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Application.Visible = True
    Sheet1.Activate
    Application.Visible = False
    
End Sub

Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub

Public Sub FillEntries(ENTRY As iGenerateLabels.ENTRY_MODE)

    Select Case ENTRY
        Case ENTRY_MODE.MANUAL:
            ''Do Nothing
        Case ENTRY_MODE.LOAD_OUTLOOK_ENTRIES
            DataAccess.SetUserLabelsOutlook
        Case ENTRY_MODE.LOAD_SAVED_SET
        
    End Select

End Sub
Public Sub ListenTo(EventProvider As Object)
    If TypeOf EventProvider Is WinManualEntry Then: Set mManualEntry = WinManualEntry
End Sub
