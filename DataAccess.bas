Attribute VB_Name = "DataAccess"
Private m_LabelList As Collection, m_CorrectionDict As Dictionary, m_DictSORecord As Dictionary
Public Function labelArray() As String()
    labelArray = PKLib.ToStrArray2D(m_LabelList)
End Function
Public Function LabelCount() As Integer
    LabelCount = m_LabelList.count
End Function
Public Sub UpdateLabelArray(indexBase0 As Integer, info() As String)
    Dim i As Integer, j As Integer, newCollection As Collection
    
    Set newCollection = New Collection
    
    For i = 1 To m_LabelList.count
        If i - 1 = indexBase0 Then
            newCollection.add info
        Else
            newCollection.add m_LabelList(i)
        End If
    Next i
    
    Set m_LabelList = newCollection
End Sub
Public Sub ClearLabels()
    ''Set m_LabelList = New Collection
    Dim row As Integer
    For row = 1 To m_LabelList.count
        m_LabelList.Remove 1
    Next row
End Sub
Public Sub AddLabel(info() As String)
    If m_LabelList Is Nothing Then: Set m_LabelList = New Collection
    m_LabelList.add info
End Sub
Public Sub RemoveLabel(indexBase0 As Integer)
    m_LabelList.Remove indexBase0 + 1
End Sub
Public Function DataIsEmpty() As Boolean
    If m_LabelList Is Nothing Then: Set m_LabelList = New Collection
    DataIsEmpty = CBool(m_LabelList.count = 0)
End Function
Public Function GetPO(key As String) As String
    If m_DictSORecord Is Nothing Then: AssembleDataAccess
    If Not m_DictSORecord.Exists(key) Then
        GetPO = "<NOT FOUND>"
        Exit Function
    End If
    GetPO = Trim(m_DictSORecord.Item(key)(4))

End Function
Public Function GetCustomerName(key As String) As String
    If m_DictSORecord Is Nothing Then: AssembleDataAccess
    If Not m_DictSORecord.Exists(key) Then
        GetCustomerName = "<NOT FOUND>"
        Exit Function
    End If
    GetCustomerName = Truncate(GetCorrectCName(Trim(GetSoldTo(key)), Trim(m_DictSORecord.Item(key)(1))), 25)
End Function
Public Function GetCSRep(key As String) As String
    If m_DictSORecord Is Nothing Then: AssembleDataAccess
    If Not m_DictSORecord.Exists(key) Then
        GetCSRep = "<NOT FOUND>"
        Exit Function
    End If
    GetCSRep = Trim(m_DictSORecord.Item(key)(2))

End Function
Private Function GetCorrectCName(key As String, default As String) As String
    If m_CorrectionDict Is Nothing Then: AssembleDataAccess
    If m_CorrectionDict.Exists(key) Then
        GetCorrectCName = Trim(m_CorrectionDict(key)(1))
    Else
        GetCorrectCName = default
    End If
End Function
Public Function GetSoldTo(key As String) As String
    If m_DictSORecord Is Nothing Then: AssembleDataAccess
    If Not m_DictSORecord.Exists(key) Then
        GetSoldTo = "<NOT FOUND>"
        Exit Function
    End If
    GetSoldTo = m_DictSORecord.Item(key)(3)
    
End Function
Public Function AddPreferredName(soldTo As String, preferredName As String) As Boolean
    Dim rCursor As CCursorReader, result() As String, rqColumns(0) As String, query(0) As String, success As Boolean
    Dim newData As Collection, wCursor As CCursorWriter
    On Error GoTo ErrorHandler
    Set newData = New Collection
    
    newData.add soldTo
    newData.add preferredName
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_NameFix").value, "|", rqColumns, query)
    
    RemoveIfExists newData(1), rCursor
    
    success = rCursor.TryAddRecord(newData)
         
    Set wCursor = New CCursorWriter
    Set wCuror = wCursor.GetCursorWriter(rCursor)
    
    wCursor.WriteToFile Main.Program.StoreObject("PATH_NameFix").value, "|"
    AddPreferredName = success
Exit Function
ErrorHandler:
    Stop
    Resume
    AddPreferredName = False
End Function
Private Sub RemoveIfExists(value As String, rCursor As CCursorReader)
    Dim recordIndex As Integer
    For recordIndex = 0 To rCursor.recordCount - 1
        If rCursor.GetRecord(recordIndex)(0) = value Then
            rCursor.TryRemoveRecord recordIndex
        Exit Sub
        End If
    Next recordIndex
    
End Sub

Public Function ChangePreferredName(soldTo As String, preferredName As String) As Boolean
    Dim rCursor As CCursorReader, result() As String, rqColumns(0) As String, query(0) As String
    Dim row As Integer
    Dim newData As Collection, wCursor As CCursorWriter
    
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_NameFix").value, "|", rqColumns, query)
    
    For row = 0 To rCursor.recordCount - 1
        If Trim(rCursor.GetRecord(row)(0)) = soldTo Then
            RemovePreferredName (row)
            Exit For
        End If
    Next row
    
    Set newData = New Collection
    newData.add soldTo
    newData.add preferredName
    rCursor.AddRecord newData
    
    Set wCursor = New CCursorWriter
    Set wCursor = wCursor.GetCursorWriter(rCursor)
    
    wCursor.WriteToFile Main.Program.StoreObject("PATH_NameFix").value, "|"
End Function

Public Function GetPreferredNames() As String()
    Dim rCursor As CCursorReader, result() As String, rqColumns(0) As String, query(0) As String
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_NameFix").value, "|", rqColumns, query)
    
    GetPreferredNames = rCursor.GetRecords2DArray
End Function
Public Sub RemovePreferredName(index As Integer)
    Dim rCursor As CCursorReader, result() As String, rqColumns(0) As String, query(0) As String
    Dim wCursor As CCursorWriter
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_NameFix").value, "|", rqColumns, query)
    If rCursor.TryRemoveRecord(index) Then
    
        Set wCursor = New CCursorWriter
        Set wCursor = wCursor.GetCursorWriter(rCursor)
        wCursor.WriteToFile Main.Program.StoreObject("PATH_NameFix").value, "|"
        
    Else
        Error "Invalid index requested for removal"
    End If
End Sub
Public Function WriteCorrection(soldToNum As String, preferedName As String)
    Dim newRecord As Collection, rCursor As CCursorReader, wCursor As CCursorWriter, rqColumns(1) As String, query(0) As String, correctionPath As String
    Set newRecord = New Collection
    correctionPath = Main.Program.StoreObject("PATH_NameFix").value
    
    newRecord.add soldToNum
    newRecord.add preferedName
    
    rqColumns(0) = "Sold-to pt"
    rqColumns(1) = "Name 1"
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(correctionPath, "|", rqColumns, query)
    rCursor.TryAddRecord newRecord
    
    Set wCursor = New CCursorWriter
    Set wCursor = wCursor.GetCursorWriter(rCursor)
    wCursor.WriteToFile correctionPath, "|"
End Function
Private Sub AssembleDataAccess()
    Set m_CorrectionDict = Main.Program.StoreObject("DICT_NameFix")
    Set m_DictSORecord = Main.Program.StoreObject("DICT_Sales")
End Sub
Public Sub SetPicture(picture As StdPicture, storeArg As String)
    Dim picPath As String, picDisp As IPictureDisp
    picPath = Main.Program.StoreObject(storeArg).value
    Set picture = StdFunctions.LoadPicture(picPath)
    
End Sub
Public Sub SetUserLabelsOutlook()
    Dim usrPath As String, fso As FileSystemObject, tStream As TextStream, line As String, fileDate As Date, label(3) As String
    usrPath = "C:\Users\" & Environ$("Username") & "\Documents\Today's Outlook SOs.txt"
    
    Set fso = New FileSystemObject
    Set tStream = fso.OpenTextFile(usrPath, ForReading, True)
    
    If tStream.AtEndOfStream Then
        MsgBox "No entries found."
        tStream.Close
        Exit Sub
    End If
    
    fileDate = CDate(Trim(tStream.ReadLine))
    tStream.Close
    
    If fileDate = Date Then
        Set tStream = fso.OpenTextFile(usrPath, ForReading): tStream.SkipLine ''<---Skip the first line(it's the date)
        Do Until tStream.AtEndOfStream
            line = tStream.ReadLine
            label(0) = line
            label(1) = GetCustomerName(line)
            label(2) = GetPO(line)
            label(3) = GetCSRep(line)
            AddLabel label
        Loop
        tStream.Close
    Else
        MsgBox "No entries found."
    End If
    
End Sub
Public Sub SetUserLabelsQuery(labels() As String)
    Dim row As Integer, label(3) As String, key As String
    
    For row = 0 To UBound(labels)
        key = Trim(labels(row, 0))
        label(0) = key
        label(1) = GetCustomerName(key)
        label(2) = GetPO(key)
        label(3) = GetCSRep(key)
        AddLabel label
    Next row
    
End Sub
Public Sub UpdateManifest(fileName As String, filePath As String)
    Dim rCursor As CCursorReader, wCursor As CCursorWriter, rqColumns(1) As String, query(0) As String, record As Integer, data As Collection
    
    Set data = New Collection
    data.add fileName
    data.add filePath
    
    rqColumns(0) = "Object_Name"
    rqColumns(1) = "File_Path"
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(ThisWorkbook.path & "\AppManifest.txt", "|", rqColumns, query)
    For record = 0 To rCursor.recordCount - 1
        If Trim(rCursor.GetRecord(record)(0)) = fileName Then
            rCursor.TryRemoveRecord record
            rCursor.TryAddRecord data
            Exit For
        End If
    Next record
    
End Sub
Public Function GetQueries() As String()
    Dim rCursor As CCursorReader, rqColumns(0) As String, query(0) As String
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReader(Main.Program.StoreObject("PATH_Queries").value, "|", rqColumns, query)
    
    GetQueries = rCursor.GetRecords2DArray
End Function
Public Function GetColumnNames() As String()
    Dim rCursor As CCursorReader, query(0) As String
    
    Set rCursor = New CCursorReader
    Set rCursor = rCursor.GetCursorReaderOnlyHeaders(Main.Program.StoreObject("PATH_USOrders").value, "|", query)
    
    GetColumnNames = rCursor.GetColumnNames
End Function
Private Function Truncate(strVal As String, numMaxChars As Integer) As String
    If Len(strVal) > numMaxChars Then
        Truncate = Left$(strVal, numMaxChars) & "..."
    Else
        Truncate = strVal
    End If
End Function
