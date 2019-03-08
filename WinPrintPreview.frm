VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WinPrintPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "WinPrintPreview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WinPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_LabelsThisPage(17) As Boolean, m_Pages As Collection, m_ActivePage As Integer
Private Sub ButtonBack_Click()
    Alternate.ButtonBack_Click
End Sub
Private Sub ButtonPageLeft_Click()
    Alternate.ButtonPageLeft_Click
End Sub
Private Sub ButtonPageRight_Click()
    Alternate.ButtonPageRight_Click
End Sub

Private Sub ButtonPrint_Click()
    Dim msg As String
    PrintClient.Run GetPrintingPages
    
    msg = "Printing finished." & vbCrLf & vbCrLf & "Please review the printed documents before selecting an option below."
    
    WinPrintDialogue.Label1 = msg
    WinPrintDialogue.Show
End Sub
Private Sub Label1_Click()
    Alternate.Toggle 1
End Sub
Private Sub Label2_Click()
    Alternate.Toggle 2
End Sub
Private Sub Label3_Click()
    Alternate.Toggle 3
End Sub
Private Sub Label4_Click()
    Alternate.Toggle 4
End Sub
Private Sub Label5_Click()
    Alternate.Toggle 5
End Sub
Private Sub Label6_Click()
    Alternate.Toggle 6
End Sub
Private Sub Label7_Click()
    Alternate.Toggle 7
End Sub
Private Sub Label8_Click()
    Alternate.Toggle 8
End Sub
Private Sub Label9_Click()
    Alternate.Toggle 9
End Sub
Private Sub Label10_Click()
    Alternate.Toggle 10
End Sub
Private Sub Label11_Click()
    Alternate.Toggle 11
End Sub
Private Sub Label12_Click()
    Alternate.Toggle 12
End Sub
Private Sub Label13_Click()
    Alternate.Toggle 13
End Sub
Private Sub Label14_Click()
    Alternate.Toggle 14
End Sub
Private Sub Label15_Click()
    Alternate.Toggle 15
End Sub
Private Sub Label16_Click()
    Alternate.Toggle 16
End Sub
Private Sub Label17_Click()
    Alternate.Toggle 17
End Sub
Private Sub Label18_Click()
    Alternate.Toggle 18
End Sub
Private Sub DrawLabels()
    Dim stickers() As CSticker, openPositions() As Integer, position As Variant, stickerNum As Integer, activeLabel As MSForms.label, errorCount As Integer
    
    AdjustPageCount
    
    stickers = GetStickers
    stickerNum = GetStartingSticker
    openPositions = GetOpenPrintingPositionsAlternate
    
    Do While stickerNum > m_Pages.count * 18
        If errorCount > 300 Then
            MsgBox "A program  internal error has occurred."
            Exit Sub
        End If
        errorCount = errorCount + 1
        stickerNum = GetStartingSticker
    Loop
    
    For Each position In openPositions
        Set activeLabel = GetActiveLabel(CStr(position) & " KEY")
            activeLabel.Caption = stickers(stickerNum - 1).CustomerName & vbCrLf & _
                              stickers(stickerNum - 1).SalesOrderNumber
        stickerNum = stickerNum + 1
    Next position
    
    For position = 1 To 18
        Set activeLabel = GetActiveLabel(CStr(position) & " KEY")
        AssignColorAlternate activeLabel, CInt(position)
    Next position
    
    DrawPageLabel
    
End Sub
Private Sub Toggle(position As Integer)
    Dim activeLabel As MSForms.label, boolCollection As Collection, bool As CEncapsulation
    
    Set activeLabel = GetActiveLabel(CStr(position) & " KEY")
    Set boolCollection = m_Pages(m_ActivePage)
    Set bool = New CEncapsulation
    
    If IsPrintableAlternate(m_ActivePage, position) Then
        ''Turn it off
        bool.Value = False
        m_Pages(m_ActivePage).Remove CStr(position) & " KEY"
        m_Pages(m_ActivePage).Add bool, CStr(position) & " KEY"
    Else
        ''Turn it on
        bool.Value = True
        m_Pages(m_ActivePage).Remove CStr(position) & " KEY"
        m_Pages(m_ActivePage).Add bool, CStr(position) & " KEY"
    End If
    
    DrawLabels
    
End Sub
Private Function IsPrintable(position As Integer) As Boolean
    Dim table As Range
    
    Set table = Names("Printing_Positions").RefersToRange
    
    IsPrintable = table(position, 2)
    
End Function
Private Function IsPrintableAlternate(Page As Integer, position As Integer) As Boolean
    
    IsPrintableAlternate = m_Pages(Page)(position).Value
    
End Function
Private Function GetActiveLabel(position As String) As MSForms.label
     Dim activeLabel As MSForms.label
     
    Select Case position
        Case "1 KEY"
            Set activeLabel = Label1
        Case "2 KEY"
            Set activeLabel = Label2
        Case "3 KEY"
            Set activeLabel = Label3
        Case "4 KEY"
            Set activeLabel = Label4
        Case "5 KEY"
            Set activeLabel = Label5
        Case "6 KEY"
            Set activeLabel = Label6
        Case "7 KEY"
            Set activeLabel = Label7
        Case "8 KEY"
            Set activeLabel = Label8
        Case "9 KEY"
            Set activeLabel = Label9
        Case "10 KEY"
            Set activeLabel = Label10
        Case "11 KEY"
            Set activeLabel = Label11
        Case "12 KEY"
            Set activeLabel = Label12
        Case "13 KEY"
            Set activeLabel = Label13
        Case "14 KEY"
            Set activeLabel = Label14
        Case "15 KEY"
            Set activeLabel = Label15
        Case "16 KEY"
            Set activeLabel = Label16
        Case "17 KEY"
            Set activeLabel = Label17
        Case "18 KEY"
            Set activeLabel = Label18
    End Select
    
    Set GetActiveLabel = activeLabel
End Function
Public Function GetOpenPrintingPositions() As Integer()
    Dim ints As Collection, intArray() As Integer, table As Range, i As Variant
    
    Set ints = New Collection
    Set table = Names("Printing_Positions").RefersToRange
    
    For i = 1 To 18
        If table(CInt(i), 2) Then
            ints.Add (CInt(i))
        End If
    Next i
    
    If ints.count <> 0 Then
        ReDim intArray(ints.count - 1) As Integer
        
        j = 0
        For Each i In ints
            intArray(j) = CInt(i)
            j = j + 1
        Next i
        
        GetOpenPrintingPositions = intArray
    Else
        ReDim intArray(20) As Integer
        GetOpenPrintingPositions = intArray
    End If
    
End Function
Public Function GetOpenPrintingPositionsAlternate() As Integer()
    Dim val As Boolean, i As Variant, ints As Collection, intArray() As Integer, data As Object
    
    Set ints = New Collection
    
    For i = 1 To 18
        If m_Pages(m_ActivePage)(i & " KEY").Value Then
            ints.Add i, i & " KEY"
        End If
    Next i
    
    If ints.count <> 0 Then
        ReDim intArray(ints.count - 1) As Integer
        
        j = 0
        For Each i In ints
            intArray(j) = ints(CStr(i) & " KEY")
            j = j + 1
        Next i
    Else
        ReDim intArray(20) As Integer
    End If
    
    GetOpenPrintingPositionsAlternate = intArray
    
End Function
Private Function GetConditionalPrintCell(position As Integer) As Range
    Dim table As Range
    
    Set table = Names("Printing_Positions").RefersToRange
    
    Set GetConditionalPrintCell = table(position, 2)
End Function
Private Sub AssignColor(activeLabel As MSForms.label, position As Integer)
    Dim table As Range, boolVal As Boolean
    
    Set table = Names("Printing_Positions").RefersToRange
    boolVal = table(position, 2)
    
    Select Case boolVal
        Case True
            activeLabel.BackColor = &H8000000B
        Case False
            activeLabel.BackColor = &H80000007
    End Select
    
End Sub
Private Sub AssignColorAlternate(activeLabel As MSForms.label, position As Integer)
    Dim boolVal As Boolean
    
    boolVal = m_Pages(m_ActivePage)(position & " KEY").Value
    
    Select Case boolVal
        Case True
            activeLabel.BackColor = &H8000000B
        Case False
            activeLabel.BackColor = &H80000007
    End Select
    
End Sub
Private Function RoundUpTo18(val As Integer) As Integer
    Dim num As Integer
    num = Abs(val - 18) + val
    
    If num > 18 Then: num = val + (val Mod 18)
        
    RoundUpTo18 = num
End Function

Private Sub UserForm_Activate()
    Alternate.InitializeData
    'Alternate.DrawLabels
End Sub
Private Sub InitializeData()
    Dim newCollection As Collection
    Set m_Pages = New Collection
    
    Set newCollection = GetBlankCollection(18, True)
    m_ActivePage = 1
    m_Pages.Add newCollection, "1 KEY"
    
End Sub
Private Function GetBlankArray(indexCount As Integer, defaultValue As Variant) As Variant
    Dim result() As Variant, i As Integer
    
    ReDim result(indexCount - 1) As Variant
    
    For i = 0 To UBound(result)
        result(i) = defaultValue
    Next i
    
    GetBlankArray = result
End Function
Private Function GetBlankCollection(indexCount As Integer, defaultValue As Variant) As Collection
    Dim result As Collection, i As Integer, bool As CEncapsulation
    
    Set result = New Collection
    
    For i = 1 To indexCount
        Set bool = New CEncapsulation
        bool.Value = defaultValue
        result.Add bool, CStr(i) & " KEY"
    Next i
    
    Set GetBlankCollection = result
End Function
Private Sub DrawPageLabel()
    ''Toggle availability of the page left button
    If m_ActivePage = 1 Then
        ButtonPageLeft.Enabled = False
    Else
        ButtonPageLeft.Enabled = True
    End If
    
    ''Toggle the availability of the page right button
    If m_Pages.count = m_ActivePage Then
        ButtonPageRight.Enabled = False
    Else
        ButtonPageRight.Enabled = True
    End If
    
    ''Fill in the label at the bottom
    PageLabel.Caption = "Page " & CStr(m_ActivePage) & "/" & CStr(m_Pages.count)
End Sub
Private Function GetStartingSticker() As Integer
    Dim Page As Integer, sticker As Integer, count As Integer
    
    If m_ActivePage <> 1 Then
        For Page = 1 To m_ActivePage
            For sticker = 1 To 18
                If IsPrintableAlternate(Page, sticker) Then
                    count = count + 1
                End If
            Next sticker
        Next Page
    End If
    
    GetStartingSticker = count + 1
End Function
Private Function GetPrintablesCount(Page As Integer) As Integer
    Dim count As Integer, i As Integer
    
    For i = 1 To 18
        If IsPrintableAlternate(Page, i) Then
            count = count + 1
        End If
    Next i
    
    GetPrintablesCount = count
End Function
Private Function GetPrintablesFromTo(fromPage As Integer, toPage As Integer) As Integer
    Dim i As Integer, count As Integer
    
    For i = fromPage To toPage
        count = count + GetPrintablesCount(i)
    Next i
    
    GetPrintablesFromTo = count
End Function
Private Function GetTotalPrintablesCount() As Integer
    Dim count As Integer, i As Integer
    
    For i = 1 To m_Pages.count
        count = count + GetPrintablesCount(i)
    Next i
    
    GetTotalPrintablesCount = count
End Function
Private Sub AdjustPageCount()
    
    ''If there aren't enough pages for all the labels, add one until there are enough pages
    Do While m_Pages.count > (GetTotalPrintablesCount / 18)
        m_Pages.Add GetFreshPage
    Loop
    
    ''If there all the labels fit on one page less that what there is, remove a page from the end, loop
    Do While DataAccess.LabelCount < GetPrintablesFromTo(1, m_Pages.count - 1)
       m_Pages.Remove CStr(m_Pages.count) & " KEY"
    Loop
    
End Sub
