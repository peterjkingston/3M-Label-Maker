Attribute VB_Name = "DataAccess"
Private m_LabelList As Collection, m_CorrectionDict As Dictionary, m_DictSORecord As Dictionary
Public m_LogPath As String
Public m_LogPathCA As String
Public m_SOFilePath As String
Public m_CorrectionPath As String
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
            newCollection.Add info
        Else
            newCollection.Add m_LabelList(i)
        End If
    Next i
    
    Set m_LabelList = newCollection
End Sub
Public Sub ClearLabels()
    Set m_LabelList = New Collection
End Sub
Public Sub AddLabel(info() As String)
    m_LabelList.Add info
End Sub
Public Sub RemoveLabel(indexBase0 As Integer)
    m_LabelList.Remove indexBase0 + 1
End Sub
Public Function AssembleDict(arrayOfCursors As Variant) As Dictionary
    Dim result As Dictionary, cursor As Integer, record As Integer, aCursor As CCursor, currentKey As String
    
    Set result = New Dictionary
    
    For cursor = 0 To UBound(arrayOfCursors)
        Set aCursor = arrayOfCursors(cursor)
        For record = 1 To aCursor.recordCount
            currentKey = Trim(aCursor.GetRecord(record)(0))
            If Not result.Exists(currentKey) Then: result.Add currentKey, aCursor.GetRecord(record)
        Next record
    Next cursor
    
    Set AssembleDict = result
End Function
Public Sub AssembleGlobalDict()
    Dim aCursorWD As CCursor, aCursorCA As CCursor, rqColumns(4) As String, query(0) As String
    Set aCursorWD = New CCursor
    
    rqColumns(0) = "Document"
    rqColumns(1) = "Name 1"
    rqColumns(2) = "Created"
    rqColumns(3) = "Sold-to pt"
    rqColumns(4) = "Purchase order number"
    
    Set aCursorWD = aCursorWD.GetCursor(m_LogPath, "|", rqColumns, query)
    
    rqColumns(0) = "Document"
    rqColumns(1) = "Name 1"
    rqColumns(2) = "Created"
    rqColumns(3) = "Sold-to pt"
    rqColumns(4) = "PO number"
    
    Set aCursorCA = aCursorWD.GetCursor(m_LogPathCA, "|", rqColumns, query)
    
    Set m_DictSORecord = AssembleDict(Array(aCursorWD, aCursorCA))
End Sub
Public Sub AssembleCorrectionDict()
    Dim aCursor As CCursor, rqColumns(1) As String, query(0) As String
    Set aCursor = New CCursor
    
    If Not m_CorrectionDict Is Nothing Then: Exit Sub
    
    rqColumns(0) = "Sold-to pt"
    rqColumns(1) = "Name 1"
    
    Set aCursor = aCursor.GetCursor(m_CorrectionPath, "|", rqColumns, query)
    
    Set m_CorrectionDict = AssembleDict(Array(aCursor))
End Sub
Public Function DataIsEmpty() As Boolean
    If m_LabelList Is Nothing Then: Set m_LabelList = New Collection
    DataIsEmpty = CBool(m_LabelList.count = 0)
End Function
Public Function GetPO(key As String) As String
    If Not m_DictSORecord.Exists(key) Then
        GetPO = "<NOT FOUND>"
        Exit Function
    End If
    GetPO = m_DictSORecord.Item(key)(4)

End Function
Public Function GetCustomerName(key As String) As String
    If Not m_DictSORecord.Exists(key) Then
        GetCustomerName = "<NOT FOUND>"
        Exit Function
    End If
    GetCustomerName = GetCorrectCName(key)

End Function
Public Function GetCSRep(key As String) As String
    If Not m_DictSORecord.Exists(key) Then
        GetCSRep = "<NOT FOUND>"
        Exit Function
    End If
    GetCSRep = m_DictSORecord.Item(key)(2)

End Function
Private Function GetCorrectCName(key As String) As String
    If m_CorrectionDict.Exists(key) Then
        GetCorrectCName = m_CorrectionDict(key)
    Else
        GetCorrectCName = m_DictSORecord.Item(key)(1)
    End If
End Function
Public Function GetSoldTo(key As String) As String
    If Not m_DictSORecord.Exists(key) Then
        GetSoldTo = "<NOT FOUND>"
        Exit Function
    End If
    GetSoldTo = m_DictSORecord.Item(key)(3)
    
End Function
Public Function WriteCorrection(soldToNum As String, preferedName As String)
    Dim newRecord As Collection, aCursor As CCursor, rqColumns(1) As String, query(0) As String
    Set newRecord = New Collection
    newRecord.Add soldToNum
    newRecord.Add preferedName
    
    rqColumns(0) = "Sold-to pt"
    rqColumns(1) = "Name 1"
    
    Set aCursor = New CCursor
    Set aCursor = aCursor.GetCursor(m_CorrectionPath, "|", rqColumns, query)
    aCursor.TryAddRecord newRecord
    
    aCursor.WriteToText m_CorrectionPath, "|"
End Function
