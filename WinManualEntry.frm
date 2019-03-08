VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinManualEntry 
   Caption         =   "Manual Entry"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "WinManualEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinManualEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditMode As Boolean
Private Sub AutomaticLookupEnabled_Click()
    DrawPage
End Sub
Private Sub DrawPage()
    If Not AutomaticLookupEnabled Then
        TextBoxPO.Enabled = True
        TextBoxPO.BackStyle = fmBackStyleOpaque
        TextBoxCName.Enabled = True
        TextBoxCName.BackStyle = fmBackStyleOpaque
        TextBoxCSRep.Enabled = True
        TextBoxCSRep.BackStyle = fmBackStyleOpaque
    Else
        TextBoxPO.Enabled = False
        TextBoxPO.BackStyle = fmBackStyleTransparent
        TextBoxCName.Enabled = False
        TextBoxCName.BackStyle = fmBackStyleTransparent
        TextBoxCSRep.Enabled = False
        TextBoxCSRep.BackStyle = fmBackStyleTransparent
    End If
    If Trim(Me.TextBoxSO) <> "" Then
        Me.ButtonCorrection.Enabled = True
    Else
        Me.ButtonCorrection.Enabled = False
    End If
End Sub
Private Sub ButtonSubmit_Click()
    Dim valArray(3) As String
    
    valArray(0) = TextBoxSO.Value
    valArray(2) = TextBoxPO.Value
    valArray(1) = TextBoxCName.Value
    valArray(3) = TextBoxCSRep.Value
    
    If EditMode Then
        Functions.ReplaceData valArray
    Else
        Functions.InputData valArray
    End If
    WinLogNav.DrawPage
End Sub

Private Sub ButtonCorrection_Click()
    Dim yesno As Integer
    yesno = MsgBox("Change the default name for" & vbCrLf & "Sold-to number " & Trim(DataAccess.GetSoldTo(Trim(Me.TextBoxSO))) & vbCrLf & " to " & vbCrLf & Trim(Me.TextBoxCName) & "?", vbYesNo, "Customer Name Correction")

    If yesno = vbYes Then
        DataAccess.WriteCorrection DataAccess.GetSoldTo(Trim(Me.TextBoxSO)), Trim(Me.TextBoxCName)
    End If
End Sub

Private Sub TextBoxSO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ''When any other object gains focus...
    Dim so As String
    If TextBoxSO.TextLength >= 7 Then
        If AutomaticLookupEnabled Then
            ''...autofill the form
            so = TextBoxSO.text
            TextBoxPO.Value = DataAccess.GetPO(so)
            TextBoxCName.Value = DataAccess.GetCustomerName(so)
            TextBoxCSRep.Value = DataAccess.GetCSRep(so)
        End If
    End If
    DrawPage
    
End Sub
Private Sub UserForm_Activate()
    Dim record As Integer, records() As String
    
    record = WinLogNav.ListBox1.ListIndex
    
    If EditMode Then
        
        With Me
            records = DataAccess.labelArray
            .TextBoxSO.text = records(record, 0)
            .TextBoxPO.text = records(record, 1)
            .TextBoxCName.text = records(record, 2)
            .TextBoxCSRep.text = records(record, 3)
             
            .AutomaticLookupEnabled = False
        End With
    Else
        Me.AutomaticLookupEnabled = True
    End If
    
    DrawPage
End Sub
