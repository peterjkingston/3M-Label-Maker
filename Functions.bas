Attribute VB_Name = "Functions"
Public g_Lookups As Collection
Public Sub ShowMe(Optional loadFromEmail As Boolean = True)
    ''ThisWorkbook.RefreshAll //Results in the application being unable to close.
    If loadFromEmail Then: ''LoadLabels
    WinLogNav.Show
End Sub
Public Function IsQueryValid() As Boolean

End Function
Public Sub InputData(inputArray() As String)
    DataAccess.AddLabel inputArray
End Sub
Public Function ArrayIndexOfString(strVal As String, strAry() As String) As Integer
    Dim i As Integer
    For i = 0 To UBound(strAry)
        If strVal = strAry(i) Then
            ArrayIndexOfString = i
            Exit Function
        End If
    Next i
End Function
Public Function GetPrintingPages(stickers As Collection) As Collection
    
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
Public Sub TurnOn()
    Application.Visible = True
End Sub
