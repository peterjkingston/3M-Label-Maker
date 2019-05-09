Attribute VB_Name = "Dictionaries"
Option Explicit
Private Function AssembleDict(arrayOfCursors As Variant) As Dictionary
    Dim result As Dictionary, cursor As Integer, record As Integer, rCursor As CCursorReader, currentKey As String
    
    Set result = New Dictionary
    
    For cursor = 0 To UBound(arrayOfCursors)
        Set rCursor = arrayOfCursors(cursor)
        For record = 1 To rCursor.recordCount
            currentKey = Trim(rCursor.GetRecord(record)(0))
            If Not result.Exists(currentKey) Then: result.add currentKey, rCursor.GetRecord(record)
        Next record
    Next cursor
    
    Set AssembleDict = result
End Function
Public Sub AssembleAll(datastore As Collection)
    Dim rCursor As CCursorReader, rqColumns(1) As String, queries(0) As String, manifestPath As String, records() As String, path As String, key As String, row As Integer, pathObj As CEncapsulation
    Dim rCursorUS As CCursorReader, rCursorCA As CCursorReader
    manifestPath = ThisWorkbook.path & "\AppManifest.txt"
    
    rqColumns(0) = "Object_Name"
    rqColumns(1) = "File_Path"
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(manifestPath, "|", rqColumns, queries)
    
    records = rCursor.GetRecords2DArray
    
    For row = 0 To UBound(records, 1)
        ''Store the path for later
        path = records(row, 1)
        key = Trim(records(row, 0))
        Set pathObj = New CEncapsulation
        pathObj.Value = path
        datastore.add pathObj, "PATH_" & key
        
        ''Run file specific actions
        Select Case key
            Case "NameFix"
                datastore.add AssembleNameFix(path), "DICT_" & key
            Case "USOrders"
                Set rCursorUS = GetCursorUS(path)
            Case "CAOrders"
                Set rCursorCA = GetCursorCA(path)
            Case "Queries"
                datastore.add AssembleQueries(path), "DICT_" & key
        End Select
    Next row
    
    datastore.add AssembleDict(Array(rCursorUS, rCursorCA)), "DICT_Sales"
End Sub
Private Function AssembleNameFix(path As String) As Dictionary
    Dim rCursor As CCursorReader, rqColumns() As String, queries(0) As String
    
    rqColumns = GetColumns("CorrectionDict")
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(path, "|", rqColumns, queries)
    
    Set AssembleNameFix = AssembleDict(Array(rCursor))
End Function
Private Function GetCursorUS(path As String) As CCursorReader
    Dim rCursor As CCursorReader, rqColumns() As String, queries(0) As String
    
    rqColumns = GetColumns("USData")
    Set rCursor = New CCursorReader
    Set GetCursorUS = rCursor.GetCursorReader(path, "|", rqColumns, queries)
    
End Function
Private Function GetCursorCA(path As String) As CCursorReader
    Dim rCursor As CCursorReader, rqColumns() As String, queries(0) As String
    
    rqColumns = GetColumns("CAData")
    Set rCursor = New CCursorReader
    Set GetCursorCA = rCursor.GetCursorReader(path, "|", rqColumns, queries)
    
End Function
Private Function AssembleQueries(path As String) As Dictionary
    Dim rCursor As CCursorReader, rqColumns(5) As String, queries(0) As String
    
    rqColumns(0) = "QName"
    rqColumns(1) = "Arg1"
    rqColumns(2) = "Arg2"
    rqColumns(3) = "Arg3"
    rqColumns(4) = "Arg4"
    rqColumns(5) = "Arg5"
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(path, "|", rqColumns, queries)
    
    Set AssembleQueries = AssembleDict(Array(rCursor))
    
End Function
Private Function GetColumns(qName As String) As String()
    Dim queriesDict As Dictionary, newStrings() As String, i As Integer, lastIndex As Integer, finalStrings() As String
    Set queriesDict = Main.Program.StoreObject("DICT_Queries")
    If queriesDict.Exists(qName) Then
        newStrings = queriesDict(qName)
        
        For i = 1 To UBound(newStrings)
            If Not Trim(newStrings(i)) = "" Then
                newStrings(i) = Trim(newStrings(i))
                lastIndex = i
            Else
                lastIndex = i - 1
                Exit For
            End If
        Next i
        
        ReDim finalStrings(lastIndex - 1) As String
        
        For i = 1 To lastIndex
            finalStrings(i - 1) = newStrings(i)
        Next i
        
        GetColumns = finalStrings
    End If
End Function
''TODO
