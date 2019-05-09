VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinUserQuery 
   Caption         =   "Query"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9300
   OleObjectBlob   =   "WinUserQuery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinUserQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_addeditwindow As WinQueryEntry
Attribute m_addeditwindow.VB_VarHelpID = -1
Private m_QueryLines As Collection

Public Event LogViewRequested(dataView As CCursorReader)
Public Event Terminated()


Private Sub ButtonAdd_Click()
    Set m_addeditwindow = New WinQueryEntry
    m_addeditwindow.EditMode = False
    m_addeditwindow.Show
End Sub

Private Sub ButtonEdit_Click()
    Set m_addeditwindow = New WinQueryEntry
    m_addeditwindow.EditMode = True
    m_addeditwindow.EditIndex = ListBox1.ListIndex
    m_addeditwindow.ComboBoxWHERE.ListIndex = CInt(ListBox1.list(ListBox1.ListIndex, 3))
    m_addeditwindow.ComboBoxOPERATOR.ListIndex = CInt(ListBox1.list(ListBox1.ListIndex, 4))
    m_addeditwindow.TextBoxVALUE = ListBox1.list(ListBox1.ListIndex, 2)
    m_addeditwindow.Show
End Sub

Private Sub ButtonExecute_Click()
    Dim rCursor As CCursorReader, row As Integer, rqColumns(0) As String, query() As String, temp() As String
    
    If m_QueryLines.count = 0 Then
        MsgBox "No queries specified."
        Exit Sub
    End If
    
    Set rCursor = New CCursorReader
    For row = 1 To m_QueryLines.count
        temp = m_QueryLines(row)
        query = ExtractQuery(temp)
        Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_USOrders").Value, "|", rqColumns, query)
    Next row
    
    RaiseEvent LogViewRequested(rCursor)
End Sub

Private Sub ButtonRemove_Click()
    Dim index As Integer
    index = Me.ListBox1.ListIndex + 1
    m_QueryLines.Remove (index)
    UpdateList
End Sub

Private Sub ListBox1_Click()
    DrawPage
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ButtonEdit_Click
End Sub

Private Sub m_addeditwindow_OnSubmit(args() As String, edit As Boolean, index As Integer)
    Dim result As Collection, row As Integer
    
    If edit Then
        Set result = New Collection
        For row = 1 To m_QueryLines.count
            If row = index + 1 Then
                result.add args
            Else
                result.add m_QueryLines(row)
            End If
        Next row
        Set m_QueryLines = result
    Else
        ''Just add it
        m_QueryLines.add args
    End If
    UpdateList
End Sub

Private Sub m_addeditwindow_Terminated()
    Set m_addeditwindow = Nothing
End Sub

Private Sub UserForm_Activate()
    DrawPage
End Sub

Private Sub UserForm_Initialize()
    Dim emptyStr(0, 0) As String
    
    emptyStr(0, 0) = "EMPTY"
    
    Set m_QueryLines = New Collection
    ListBox1.list = emptyStr
End Sub

Private Sub UserForm_Terminate()
    RaiseEvent Terminated
    
End Sub

Private Sub DrawPage()
    
    If ListBox1.list(0, 0) = "EMPTY" Or ListBox1.ListIndex = -1 Then
        Me.ButtonEdit.Enabled = False
        Me.ButtonRemove.Enabled = False
    Else
        Me.ButtonEdit.Enabled = True
        Me.ButtonRemove.Enabled = True
    End If
    
End Sub

Private Sub UpdateList()
    Dim emptyStr(0, 0) As String
    
    If m_QueryLines.count > 0 Then
        ListBox1.list = PKLib.ToStrArray2D(m_QueryLines)
    Else
        emptyStr(0, 0) = "EMPTY"
        Me.ListBox1.list = emptyStr
    End If
End Sub

Private Function ExtractQuery(query() As String) As String()
    Dim result() As String
    
    ReDim result(CInt(query(3))) As String
    result(CInt(query(3))) = GetIdentifier(CInt(query(4))) & query(2)
    
    ExtractQuery = result
End Function
Private Function GetIdentifier(index As Integer) As String
    Dim result As String
    
    Select Case index
        Case 0: result = "="
        Case 1: result = "~"
        Case 2: result = ">"
        Case 3: result = "<"
    End Select
    
    GetIdentifier = result
End Function
