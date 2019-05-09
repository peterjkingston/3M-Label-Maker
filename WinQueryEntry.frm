VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinQueryEntry 
   Caption         =   "Query Entry"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "WinQueryEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinQueryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditMode As Boolean, EditIndex As Integer
Public Event Terminated()
Public Event OnSubmit(args() As String, edit As Boolean, EditIndex As Integer)
Private Sub ButtonSubmit_Click()
    Dim args(4) As String
    
    args(0) = Me.ComboBoxWHERE.Value
    args(1) = Me.ComboBoxOPERATOR.Value
    args(2) = Me.TextBoxVALUE.Value
    args(3) = Me.ComboBoxWHERE.ListIndex
    args(4) = Me.ComboBoxOPERATOR.ListIndex
    
    RaiseEvent OnSubmit(args, EditMode, EditIndex)
    
End Sub

Private Sub UserForm_Initialize()

    Me.ComboBoxWHERE.list = GetColumnList
    Me.ComboBoxOPERATOR.list = GetOperatorList
    Me.ComboBoxOPERATOR.ListIndex = 0
    
End Sub
Private Function GetOperatorList() As String()
    Dim oList(3) As String
    
    oList(0) = "Equal to(=)"
    oList(1) = "Like(~)"
    oList(2) = "Greater than(>)"
    oList(3) = "Less than(<)"
    
    GetOperatorList = oList
End Function
Private Function GetColumnList() As String()
    Dim cList() As String
    
    cList = DataAccess.GetColumnNames
    
    GetColumnList = cList
End Function

Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub
