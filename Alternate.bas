Attribute VB_Name = "Alternate"
Public m_Pages As Collection
Private m_ActivePage As Integer, m_labelDisplayers As Collection
Public Sub InitializeData()
    Dim newPage As CPrintablePage
    Set m_Pages = New Collection
    
    Set newPage = GetFreshPage
    m_ActivePage = 1
    m_Pages.Add newPage
    
    Set m_labelDisplayers = GetDisplayers
    
    DrawLabels
End Sub
Public Sub ButtonBack_Click()
    WinLogNav.DrawPage
    WinPrintPreview.Hide
End Sub
Public Sub ButtonPageLeft_Click()
    m_ActivePage = m_ActivePage - 1
    DrawLabels
End Sub
Public Sub ButtonPageRight_Click()
    m_ActivePage = m_ActivePage + 1
    DrawLabels
End Sub
Private Function GetFreshPage() As CPrintablePage
    Dim result As CPrintablePage
    Set result = New CPrintablePage
    
    Set result.Printables = GetBlankCollection(18, True)
    
    Set GetFreshPage = result
End Function

Private Sub DrawLabels()
    Dim stickers As Collection, printable As CPrintable, stickerNum As Integer, activeDisplayer As MSForms.label, printableID As Integer, printablePage As CPrintablePage
    
    AdjustPageCount
    
    Set stickers = GetStickers
    stickerNum = GetStartingSticker
    
    ''If a spot on the page is marked for usage, print the next available label
    printableID = 1
    For Each activeDisplayer In m_labelDisplayers
        Set printable = m_Pages(m_ActivePage).Printables(printableID & " KEY")
        If printable.IsPrintable Then
            activeDisplayer.Caption = stickers(stickerNum).Body
            stickerNum = stickerNum + 1
        End If
        printableID = printableID + 1
    Next activeDisplayer
    
    printableID = 1
    For Each activeDisplayer In m_labelDisplayers
        Set printablePage = m_Pages(m_ActivePage)
        AssignColorAlternate activeDisplayer, printableID & " KEY"
        printableID = printableID + 1
    Next activeDisplayer
    
    DrawPageLabel
    
End Sub
Public Sub Toggle(position As Integer)
    Dim activeDisplayer As MSForms.label, boolCollection As Collection, bool As CPrintable
    
    Set activeDisplayer = GetActiveLabelDisplayer(CStr(position) & " KEY")
    Set boolCollection = m_Pages(m_ActivePage).Printables
    Set bool = New CPrintable
    bool.key = CStr(position) & " KEY"
    
    
    If IsPrintableAlternate(m_ActivePage, bool.key) Then
        ''Turn it off
        bool.IsPrintable = False
        m_Pages(m_ActivePage).Printables.Remove bool.key
        m_Pages(m_ActivePage).Printables.Add bool, bool.key
    Else
        ''Turn it on
        bool.IsPrintable = True
        m_Pages(m_ActivePage).Printables.Remove bool.key
        m_Pages(m_ActivePage).Printables.Add bool, bool.key
    End If
    
    DrawLabels
    
End Sub
Private Function IsPrintableAlternate(Page As Integer, printableKEY As String) As Boolean
    
    IsPrintableAlternate = m_Pages(Page).Printables(printableKEY).IsPrintable
    
End Function
Private Function GetActiveLabelDisplayer(position As String) As MSForms.label
     Dim activeDisplayer As MSForms.label
     
    Select Case position
        Case "1 KEY"
            Set activeDisplayer = WinPrintPreview.Label1
        Case "2 KEY"
            Set activeDisplayer = WinPrintPreview.Label2
        Case "3 KEY"
            Set activeDisplayer = WinPrintPreview.Label3
        Case "4 KEY"
            Set activeDisplayer = WinPrintPreview.Label4
        Case "5 KEY"
            Set activeDisplayer = WinPrintPreview.Label5
        Case "6 KEY"
            Set activeDisplayer = WinPrintPreview.Label6
        Case "7 KEY"
            Set activeDisplayer = WinPrintPreview.Label7
        Case "8 KEY"
            Set activeDisplayer = WinPrintPreview.Label8
        Case "9 KEY"
            Set activeDisplayer = WinPrintPreview.Label9
        Case "10 KEY"
            Set activeDisplayer = WinPrintPreview.Label10
        Case "11 KEY"
            Set activeDisplayer = WinPrintPreview.Label11
        Case "12 KEY"
            Set activeDisplayer = WinPrintPreview.Label12
        Case "13 KEY"
            Set activeDisplayer = WinPrintPreview.Label13
        Case "14 KEY"
            Set activeDisplayer = WinPrintPreview.Label14
        Case "15 KEY"
            Set activeDisplayer = WinPrintPreview.Label15
        Case "16 KEY"
            Set activeDisplayer = WinPrintPreview.Label16
        Case "17 KEY"
            Set activeDisplayer = WinPrintPreview.Label17
        Case "18 KEY"
            Set activeDisplayer = WinPrintPreview.Label18
    End Select
    
    Set GetActiveLabelDisplayer = activeDisplayer
End Function
Private Sub AssignColorAlternate(activeDisplayer As MSForms.label, printableKEY As String)
    Dim boolVal As Boolean
    
    boolVal = m_Pages(m_ActivePage).Printables(printableKEY).IsPrintable
    
    Select Case boolVal
        Case True
            activeDisplayer.BackColor = &H8000000B
        Case False
            activeDisplayer.BackColor = &H80000007
    End Select
    
End Sub
Public Function GetStickers() As Collection
    Dim i As Integer, currentSticker As CSticker, stickersCollection As Collection, stickersArray() As CSticker, totalStickers As Integer, labelArray() As String
    
    Set stickersCollection = New Collection
    totalStickers = RoundUpTo18(DataAccess.LabelCount)
    labelArray = DataAccess.labelArray
    
    For i = 0 To totalStickers - 1
        Set currentSticker = New CSticker
        
        If i >= DataAccess.LabelCount Then
            currentSticker.CustomerName = ""
            currentSticker.SalesOrderNumber = ""
            currentSticker.CSName = ""
        Else
            currentSticker.CustomerName = labelArray(i, 1)
            currentSticker.SalesOrderNumber = labelArray(i, 0)
            currentSticker.CSName = labelArray(i, 3)
        End If
        
        stickersCollection.Add currentSticker
    Next i
    
    Do While (stickersCollection.count / 18) < m_Pages.count
        Set currentSticker = New CSticker
        
        stickersCollection.Add currentSticker
    Loop
    
    Set GetStickers = stickersCollection
End Function
Public Function GetStickersNoExcess() As Collection
    Dim table As Range, i As Integer, currentSticker As CSticker, stickersCollection As Collection, stickersArray() As CSticker, totalStickers As Integer
    
    Set table = Names(Globals.dataTableName).RefersToRange
    Set stickersCollection = New Collection
    totalStickers = RoundUpTo18(table.count / 2)
    
    For i = 1 To totalStickers
        Set currentSticker = New CSticker
        currentSticker.CustomerName = "" ''TODO ''table(i, Globals.dataColumnCustomerName)
        currentSticker.SalesOrderNumber = "" ''TODO ''table(i, Globals.dataColumnSO)
        stickersCollection.Add currentSticker
    Next i
    
    Set GetStickersNoExcess = stickersCollection
End Function
Private Function RoundUpTo18(val As Integer) As Integer
    Dim num As Integer
    num = Abs(val - 18) + val
    
    If num > 18 Then: num = val + (val Mod 18)
        
    RoundUpTo18 = num
End Function
Private Function GetBlankCollection(indexCount As Integer, defaultValue As Variant) As Collection
    Dim result As Collection, i As Integer, printable As CPrintable
    
    Set result = New Collection
    
    For i = 1 To indexCount
        Set printable = New CPrintable
        printable.IsPrintable = defaultValue
        printable.key = CStr(i) & " KEY"
        result.Add printable, printable.key
    Next i
    
    Set GetBlankCollection = result
End Function
Private Sub DrawPageLabel()
    ''Toggle availability of the page left button
    If m_ActivePage = 1 Then
        WinPrintPreview.ButtonPageLeft.Enabled = False
    Else
        WinPrintPreview.ButtonPageLeft.Enabled = True
    End If
    
    ''Toggle the availability of the page right button
    If m_Pages.count = m_ActivePage Then
        WinPrintPreview.ButtonPageRight.Enabled = False
    Else
        WinPrintPreview.ButtonPageRight.Enabled = True
    End If
    
    ''Fill in the label at the bottom
    WinPrintPreview.PageLabel.Caption = "Page " & CStr(m_ActivePage) & "/" & CStr(m_Pages.count)
End Sub
Private Function GetStartingSticker() As Integer
    Dim Page As Integer, sticker As Integer, count As Integer
    
    For Page = 1 To m_ActivePage - 1
        count = count + m_Pages(Page).TrueCount
    Next Page
    
    GetStartingSticker = count + 1
End Function
Private Function GetPrintablesFromTo(fromPage As Integer, toPage As Integer) As Integer
    Dim Page As Integer, count As Integer
    
    For Page = fromPage To toPage
        count = count + m_Pages(Page).TrueCount
    Next Page
    
    GetPrintablesFromTo = count
End Function
Private Function GetTotalPrintablesCount() As Integer
    Dim count As Integer, printablePage As CPrintablePage
    
    For Each printablePage In m_Pages
        count = count + printablePage.TrueCount
    Next printablePage
    
    GetTotalPrintablesCount = count
End Function
Private Sub AdjustPageCount()
    
    ''If there aren't enough pages for all the labels, add one until there are enough pages
    Do While DataAccess.LabelCount > GetTotalPrintablesCount
        m_Pages.Add GetFreshPage
    Loop
    
    ''If there all the labels fit on one page less that what there is, remove a page from the end, loop
    Do While DataAccess.LabelCount < GetPrintablesFromTo(1, m_Pages.count - 1)
       m_Pages.Remove m_Pages.count
    Loop
    
End Sub
Private Function GetDisplayers() As Collection
    Dim result As Collection
    Set result = New Collection
    
    With WinPrintPreview
        result.Add .Label1
        result.Add .Label2
        result.Add .Label3
        result.Add .Label4
        result.Add .Label5
        result.Add .Label6
        result.Add .Label7
        result.Add .Label8
        result.Add .Label9
        result.Add .Label10
        result.Add .Label11
        result.Add .Label12
        result.Add .Label13
        result.Add .Label14
        result.Add .Label15
        result.Add .Label16
        result.Add .Label17
        result.Add .Label18
    End With
    
    Set GetDisplayers = result
    
End Function


