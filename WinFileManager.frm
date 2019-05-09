VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinFileManager 
   Caption         =   "File Manager"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   OleObjectBlob   =   "WinFileManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Terminated()

Private Sub CommandButton1_Click()
    DialogFind "LoadFromFileIcon", Me.TextBox18
End Sub

Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub

Private Sub ButtonCAOrders_Click()
    DialogFind "CAOrders", Me.TextBoxCAOrders
End Sub

Private Sub ButtonEditIcon_Click()
    DialogFind "EditIcon", Me.TextBoxEdit
End Sub

Private Sub ButtonNameFix_Click()
    DialogFind "NameFix", Me.TextBoxNameFix
End Sub

Private Sub ButtonQueries_Click()
    DialogFind "Queries", Me.TextBoxQueries
End Sub

Private Sub ButtonSearchIcon_Click()
    DialogFind "SearchIcon", Me.TextBoxSearch
End Sub

Private Sub ButtonSettingsIcon_Click()
    DialogFind "SettingsIcon", Me.TextBoxSettings
End Sub

Private Sub ButtonUSOrders_Click()
    DialogFind "USOrders", Me.TextBoxUSOrders
End Sub
Private Sub UpdateFileManager(file As String, path As String, textbox As MSForms.textbox)
    DataAccess.UpdateManifest file, path
    textbox.text = path
End Sub
Private Sub DialogFind(file As String, textbox As MSForms.textbox)
    Dim openDialog As FileDialog
    Set openDialog = Application.FileDialog(msoFileDialogFilePicker)
    If openDialog.Show = -1 Then
        UpdateFileManager file, openDialog.SelectedItems(1), textbox
    End If
End Sub
Private Sub UserForm_Activate()
    Dim rCursor As CCursorReader, rqColumns(0) As String, queries(0) As String, results() As String, row As Integer
    Set rCursor = New CCursorReader
    
    Set rCursor = rCursor.GetCursorReader(ThisWorkbook.path & "\AppManifest.txt", "|", rqColumns, queries)
    results = rCursor.GetRecords2DArray
    
    For row = 0 To UBound(results)
        Select Case Trim(results(row, 0))
            Case "Queries": Me.TextBoxQueries = results(row, 1)
            Case "SearchIcon": Me.TextBoxSearch = results(row, 1)
            Case "SettingsIcon": Me.TextBoxSettings = results(row, 1)
            Case "EditIcon": Me.TextBoxEdit = results(row, 1)
            Case "FileIcon": Me.TextBoxFile = results(row, 1)
            Case "CAOrders": Me.TextBoxCAOrders = results(row, 1)
            Case "USOrders": Me.TextBoxUSOrders = results(row, 1)
            Case "LoadFromFileIcon": Me.TextBox18 = results(row, 1)
            Case "NameFix": Me.TextBoxNameFix = results(row, 1)
        End Select
    Next row
    
    IssueWarning
End Sub
Private Sub IssueWarning()
    Dim boxResult As VbMsgBoxResult
    boxResult = MsgBox("STOP" & vbCrLf & vbCrLf & "Changing data listed here can result in fatal errors throughout the application, if not performed properly." & vbCrLf & vbCrLf & _
                        "It is advised that if you are not meant to be here, click cancel now.", vbOKCancel, "!!!!WARNING!!!!")
    
    If boxResult = vbCancel Then
        Me.Hide
        UserForm_Terminate
    End If
End Sub
