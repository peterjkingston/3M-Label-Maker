VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinUserQuery 
   Caption         =   "Query"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9030
   OleObjectBlob   =   "WinUserQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinUserQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonExecute_Click()
    Dim activeQueries As Collection
    
End Sub

Private Sub UserForm_Activate()
    ComboBox1.list = PKLib.ToVarArray(PKLib.GetQueryHeaders(Sheet4))
End Sub

