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
Public Event Terminated()
Public Event RequestViewPrintDialog()
Private m_labelDisplayers As Collection
Private Sub UserForm_Terminate()
    RaiseEvent Terminated
End Sub
Private Sub ButtonBack_Click()
    Me.Hide
    RaiseEvent Terminated
End Sub
Private Sub ButtonPageLeft_Click()
    Dim adjusted As Integer
    If m_ActivePage > 1 Then: adjusted = 1
    m_ActivePage = m_ActivePage - adjusted
    DrawLabels
End Sub
Private Sub ButtonPageRight_Click()
    Dim adjusted As Integer
    If m_ActivePage < m_Pages.count Then: adjusted = 1
    m_ActivePage = m_ActivePage + adjusted
    DrawLabels
End Sub

Private Sub ButtonPrint_Click()
    Dim msg As String
    If MsgBox("Please ensure that ULINE S-19346 labels are loaded in your default printer and that your default printer is set appropriately.", vbOKCancel, "Before Printing...") = vbOK Then
        PrintClient.Run GetPrintingPages(GetStickers)
        RaiseEvent RequestViewPrintDialog
        If Me.Visible Then: DrawLabels
    End If
End Sub
Private Sub Label1_Click()
    Toggle 1
End Sub
Private Sub Label2_Click()
    Toggle 2
End Sub
Private Sub Label3_Click()
    Toggle 3
End Sub
Private Sub Label4_Click()
    Toggle 4
End Sub
Private Sub Label5_Click()
    Toggle 5
End Sub
Private Sub Label6_Click()
    Toggle 6
End Sub
Private Sub Label7_Click()
    Toggle 7
End Sub
Private Sub Label8_Click()
    Toggle 8
End Sub
Private Sub Label9_Click()
    Toggle 9
End Sub
Private Sub Label10_Click()
    Toggle 10
End Sub
Private Sub Label11_Click()
    Toggle 11
End Sub
Private Sub Label12_Click()
    Toggle 12
End Sub
Private Sub Label13_Click()
    Toggle 13
End Sub
Private Sub Label14_Click()
    Toggle 14
End Sub
Private Sub Label15_Click()
    Toggle 15
End Sub
Private Sub Label16_Click()
    Toggle 16
End Sub
Private Sub Label17_Click()
    Toggle 17
End Sub
Private Sub Label18_Click()
    Toggle 18
End Sub
Private Sub UserForm_Activate()
    InitializeData
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
        m_Pages(m_ActivePage).Printables.add bool, bool.key
    Else
        ''Turn it on
        bool.IsPrintable = True
        m_Pages(m_ActivePage).Printables.Remove bool.key
        m_Pages(m_ActivePage).Printables.add bool, bool.key
    End If
    
    DrawLabels
    
End Sub
Private Sub InitializeData()
    Dim newPage As CPrintablePage
    Set m_Pages = New Collection
    
    Set newPage = GetFreshPage
    m_ActivePage = 1
    m_Pages.add newPage
    
    Set m_labelDisplayers = GetDisplayers
    
    DrawLabels
End Sub
Private Function GetDisplayers() As Collection
    Dim result As Collection
    Set result = New Collection
    
    With Me
        result.add .Label1
        result.add .Label2
        result.add .Label3
        result.add .Label4
        result.add .Label5
        result.add .Label6
        result.add .Label7
        result.add .Label8
        result.add .Label9
        result.add .Label10
        result.add .Label11
        result.add .Label12
        result.add .Label13
        result.add .Label14
        result.add .Label15
        result.add .Label16
        result.add .Label17
        result.add .Label18
    End With
    
    Set GetDisplayers = result
    
End Function
Private Function GetFreshPage() As CPrintablePage
    Dim result As CPrintablePage
    Set result = New CPrintablePage
    
    Set result.Printables = GetBlankCollection(18, True)
    
    Set GetFreshPage = result
End Function
Private Function GetBlankCollection(indexCount As Integer, defaultValue As Variant) As Collection
    Dim result As Collection, i As Integer, printable As CPrintable
    
    Set result = New Collection
    
    For i = 1 To indexCount
        Set printable = New CPrintable
        printable.IsPrintable = defaultValue
        printable.key = CStr(i) & " KEY"
        result.add printable, printable.key
    Next i
    
    Set GetBlankCollection = result
End Function
Private Function GetActiveLabelDisplayer(position As String) As MSForms.label
     Dim activeDisplayer As MSForms.label
     
    Select Case position
        Case "1 KEY"
            Set activeDisplayer = Me.Label1
        Case "2 KEY"
            Set activeDisplayer = Me.Label2
        Case "3 KEY"
            Set activeDisplayer = Me.Label3
        Case "4 KEY"
            Set activeDisplayer = Me.Label4
        Case "5 KEY"
            Set activeDisplayer = Me.Label5
        Case "6 KEY"
            Set activeDisplayer = Me.Label6
        Case "7 KEY"
            Set activeDisplayer = Me.Label7
        Case "8 KEY"
            Set activeDisplayer = Me.Label8
        Case "9 KEY"
            Set activeDisplayer = Me.Label9
        Case "10 KEY"
            Set activeDisplayer = Me.Label10
        Case "11 KEY"
            Set activeDisplayer = Me.Label11
        Case "12 KEY"
            Set activeDisplayer = Me.Label12
        Case "13 KEY"
            Set activeDisplayer = Me.Label13
        Case "14 KEY"
            Set activeDisplayer = Me.Label14
        Case "15 KEY"
            Set activeDisplayer = Me.Label15
        Case "16 KEY"
            Set activeDisplayer = Me.Label16
        Case "17 KEY"
            Set activeDisplayer = Me.Label17
        Case "18 KEY"
            Set activeDisplayer = Me.Label18
    End Select
    
    Set GetActiveLabelDisplayer = activeDisplayer
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
Private Sub DrawPageLabel()
    ''Toggle availability of the page left button
    If m_ActivePage = 1 Then
        Me.ButtonPageLeft.Enabled = False
    Else
        Me.ButtonPageLeft.Enabled = True
    End If
    
    ''Toggle the availability of the page right button
    If m_Pages.count = m_ActivePage Then
        Me.ButtonPageRight.Enabled = False
    Else
        Me.ButtonPageRight.Enabled = True
    End If
    
    ''Fill in the label at the bottom
    Me.PageLabel.Caption = "Page " & CStr(m_ActivePage) & "/" & CStr(m_Pages.count)
End Sub
Private Sub AdjustPageCount()
    
    ''If there aren't enough pages for all the labels, add one until there are enough pages
    Do While DataAccess.LabelCount > GetTotalPrintablesCount
        m_Pages.add GetFreshPage
    Loop
    
    ''If there all the labels fit on one page less that what there is, remove a page from the end, loop
    Do While DataAccess.LabelCount < GetPrintablesFromTo(1, m_Pages.count - 1)
       m_Pages.Remove m_Pages.count
    Loop
    
End Sub
Private Function GetStickers() As Collection
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
        
        stickersCollection.add currentSticker
    Next i
    
    Do While (stickersCollection.count / 18) < m_Pages.count
        Set currentSticker = New CSticker
        
        stickersCollection.add currentSticker
    Loop
    
    Set GetStickers = stickersCollection
End Function
Private Function GetStickersNoExcess() As Collection
    Dim table As Range, i As Integer, currentSticker As CSticker, stickersCollection As Collection, stickersArray() As CSticker, totalStickers As Integer
    
    Set table = Names(Globals.dataTableName).RefersToRange
    Set stickersCollection = New Collection
    totalStickers = RoundUpTo18(table.count / 2)
    
    For i = 1 To totalStickers
        Set currentSticker = New CSticker
        currentSticker.CustomerName = "" ''TODO ''table(i, Globals.dataColumnCustomerName)
        currentSticker.SalesOrderNumber = "" ''TODO ''table(i, Globals.dataColumnSO)
        stickersCollection.add currentSticker
    Next i
    
    Set GetStickersNoExcess = stickersCollection
End Function
Private Function RoundUpTo18(val As Integer) As Integer
    Dim num As Integer
    num = Abs(val - 18) + val
    
    If num > 18 Then: num = val + (val Mod 18)
        
    RoundUpTo18 = num
End Function
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
Private Function IsPrintableAlternate(Page As Integer, printableKEY As String) As Boolean
    
    IsPrintableAlternate = m_Pages(Page).Printables(printableKEY).IsPrintable
    
End Function
Private Function GetPrintingPages(stickers As Collection) As Collection
    
    Dim ablePage As CPrintablePage, printable As CPrintable, sticker As CSticker, stickerNum As Integer, printableNum As Integer
    ''Out
    Dim printingPage As CPrintingPage, printingPages As Collection
    
    Set printingPages = New Collection
    stickerNum = 1
    
    For Each ablePage In m_Pages
        Set printingPage = New CPrintingPage
        
        For printableNum = 1 To ablePage.Printables.count
                
            Set printable = ablePage.Printables(printableNum & " KEY")
                
            If printable.IsPrintable Then
                Set sticker = stickers(stickerNum)
                stickerNum = stickerNum + 1
            Else
                Set sticker = New CSticker
            End If
            
            printingPage.stickers.add sticker
                
        Next printableNum
        printingPages.add printingPage
    Next ablePage
    
    Set GetPrintingPages = printingPages
End Function
