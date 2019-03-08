Attribute VB_Name = "Functions"
Public g_Lookups As Collection
Private Const m_LogPath As String = "C:\Users\reception\Documents\P Drive\Projects\NewLabels\sample.txt"
Private Const m_LogPathCA As String = "C:\Users\reception\Documents\P Drive\Projects\NewLabels\sampleCanada.txt"
Private Const m_SOFilePath As String = "C:\Users\reception\Documents\P Drive\Projects\NewLabels\LabelsFromEmail.txt"
Public Sub ShowMe(Optional loadFromEmail As Boolean = True)
    ''ThisWorkbook.RefreshAll //Results in the application being unable to close.
    If loadFromEmail Then: ''LoadLabels
    WinLogNav.Show
End Sub
Private Sub LoadLabels()
    Dim records As Collection, tStream As TextStream, fso As FileSystemObject, line As String, i As Integer
    
    AssembleGlobalDict
    Set records = New Collection
    Set fso = New FileSystemObject
    Set tStream = fso.OpenTextFile("C:\Users\reception\Documents\P Drive\Projects\NewLabels\LabelsFromEmail.txt", ForReading)
    
    tStream.SkipLine ''Skip the column header
    
    While Not tStream.AtEndOfStream
        line = Trim(tStream.ReadLine)
        ''records.Add m_DictSORecord(line) ''TODO FIX LATER
        Raise Error("NOT IMPLEMENTED")
    Wend
    
    tStream.Close
    
    Set tStream = fso.OpenTextFile("C:\Users\reception\Documents\P Drive\Projects\NewLabels\Labels.txt", ForAppending)
    
    For i = 1 To records.count
        If IsEmpty(records(i)) Then
            tStream.WriteLine "||||"
        Else
            tStream.WriteLine "|" & CStr(records(i)(0)) & "|" & _
                              CStr(records(i)(1)) & "|" & _
                              CStr(records(i)(2)) & "|"
        End If
    Next i
    
    tStream.Close
End Sub
Public Function IsQueryValid() As Boolean

End Function
Public Sub ReplaceData(inputArray() As String)
    DataAccess.UpdateLabelArray WinLogNav.ListBox1.ListIndex, inputArray
End Sub
Public Sub ClearLog()
    DataAccess.ClearLabels
End Sub
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
Public Sub RemoveItem(index As Integer)
    DataAccess.RemoveLabel (index)
End Sub
Public Function GetPrintingPages() As Collection
    
    Dim ablePage As CPrintablePage, printable As CPrintable, sticker As CSticker, stickers As Collection, stickerNum As Integer, printableNum As Integer
    ''Out
    Dim printingPage As CPrintingPage, printingPages As Collection
    
    Set printingPages = New Collection
    Set stickers = Alternate.GetStickers
    stickerNum = 1
    
    For Each ablePage In Alternate.m_Pages
        Set printingPage = New CPrintingPage
        
        For printableNum = 1 To ablePage.Printables.count
                
            Set printable = ablePage.Printables(printableNum & " KEY")
                
            If printable.IsPrintable Then
                Set sticker = stickers(stickerNum)
                stickerNum = stickerNum + 1
            Else
                Set sticker = New CSticker
            End If
            
            printingPage.stickers.Add sticker
                
        Next printableNum
        printingPages.Add printingPage
    Next ablePage
    
    Set GetPrintingPages = printingPages
End Function
Public Sub TurnOn()
    Application.Visible = True
End Sub
