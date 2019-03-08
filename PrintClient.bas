Attribute VB_Name = "PrintClient"
Public Sub Run(printingPages As Collection)
    Dim spot As Integer, printingPage As CPrintingPage, printingRanges, currentRange As Range
    
    Set printingRanges = GetPageRanges
    
    For Each printingPage In printingPages
        ''Map the labels to the sheet positions
        For spot = 1 To 18
            Set currentRange = printingRanges(spot)
            currentRange(1, 1) = printingPage.stickers(spot).CustomerName
            currentRange(2, 1) = printingPage.stickers(spot).SalesOrderNumber
            currentRange(3, 1) = printingPage.stickers(spot).CSName
        Next spot
        PrintPage
        ClearPage
    Next printingPage
End Sub
Private Function GetPageRanges() As Collection
    Dim tempRange As Range, ws As Worksheet, result As Collection
    
    Set result = New Collection
    Set ws = ThisWorkbook.Worksheets("LABELS 3x1")
    
    For lookuprow = 0 To 8
        row = (lookuprow * 4) + 1
        column = 1
        Set tempRange = Range(ws.Cells(row, column), ws.Cells(row + 3, column + 3))
        
        result.Add tempRange
        
        Set tempRange = Range(ws.Cells(row, column + 5), ws.Cells(row + 3, column + 8))
    
        result.Add tempRange
        
    Next lookuprow
    
    Set GetPageRanges = result
    
End Function
Private Sub PrintPage()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("LABELS 3x1")
    
    ws.PrintOut
End Sub
Private Sub ClearPage()
    Dim clearRange As Range, ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("LABELS 3x1")
    Set clearRange = Range(ws.Cells(1, 1), ws.Cells(36, 9))
    
    clearRange.ClearContents
End Sub
