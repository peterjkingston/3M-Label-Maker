VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCursorReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim myCursor As CCursorReader
Private m_IsColumnSet As Boolean, m_IsSeparatorset, m_ColumnCount As Integer, m_RecordCount As Integer
Private m_ColumnNames As Dictionary, m_Records As Collection, m_Buckets As Collection
Private m_Separators As Variant, m_Lengths As Variant
Private Const M_LIKE_IDENTIFIER As String = "~"
Private Const M_EXACT_IDENTIFIER As String = "="
Private Const M_GREATER_IDENTIFIER As String = ">"
Private Const M_LESS_IDENTIFIER As String = "<"
Public Function TryAddRecord(newCollection As Collection) As Boolean

    If Not (newCollection Is Nothing) Then
        If m_IsColumnSet Then
            
            Set newCollection = MatchLengths(newCollection)
            
            m_Records.add newCollection
            m_RecordCount = m_RecordCount + 1
            TryAddRecord = True
            Exit Function

        End If
    End If
    
    TryAddRecord = False
End Function
Public Sub AddRecord(newCollection As Collection)
    ''For use on initial file read
    
    If Not (newCollection Is Nothing) Then
        If m_IsColumnSet Then
        
            m_Records.add newCollection
            m_RecordCount = m_RecordCount + 1
            
        End If
    End If
    
End Sub
Public Function TryRemoveRecord(index As Integer) As Boolean
    
    If Not (m_Records Is Nothing) Then
        If Not index > m_Records.count Then
        
            m_Records.Remove (index)
            m_RecordCount = m_RecordCount - 1
            TryRemoveRecord = True
            Exit Function
            
        End If
    End If
    
    TryRemoveRecord = False
End Function
Public Function GetRecords2DArray() As String()
    Dim result() As String, row As Integer, column As Integer
    
    If m_RecordCount > 0 Then
    ReDim result(m_RecordCount - 1, m_ColumnCount - 1) As String
        For row = 1 To m_RecordCount
            For column = 1 To m_ColumnCount
                result(row - 1, column - 1) = m_Records(row)(column)
            Next column
        Next row
    Else
        ReDim result(0, 0) As String
    End If
    GetRecords2DArray = result
End Function
Public Function GetColumnNames() As String()
    Dim result() As String, column As Integer
    
    ReDim result(m_ColumnCount - 1) As String
    
    For column = 0 To m_ColumnCount - 1
        result(column) = Trim(m_ColumnNames.Items(column))
    Next column
    
    GetColumnNames = result
End Function
Private Function MatchLengths(newCollection As Collection) As Collection
    Dim column As Integer, record As Integer, spaceNeeded As Integer, adjusted As Collection, columnLength As Integer
    
    Set adjusted = New Collection
    
    For column = 1 To m_ColumnCount
        ''if the same index length is less then buff it up, else, make this one larger
        columnLength = m_Separators(column) - m_Separators(column - 1)
        spaceNeeded = Abs(columnLength - Len(newCollection(column)))
        
        If Len(newCollection(column)) > columnLength Then
            ''Buff up the column header, and the separators
            ''COLUMN HEADER
            FixColumnHeaders column, columnLength
            
            ''SEPARATORS
            m_Separators(column) = m_Separators(column - 1) + columnLength + spaceNeeded
            
        ElseIf Len(newCollection(column)) < columnLength Then
            ''Buff up the record column
            adjusted.add newCollection(column) & Space(spaceNeeded)
            
        Else
            adjusted.add newCollection(column)
            
        End If
    Next column
    
    Set MatchLengths = adjusted
End Function
Private Sub FixColumnHeaders(column As Integer, spaceNeeded As Integer)
    ''Brain checked out here... code no good.. maybe?
    Dim index As Integer, newDict As Dictionary
    
    Set newDict = New Dictionary
    
    For index = 0 To m_ColumnCount - 1
    
        If index = column - 1 Then
            newDict.add m_ColumnNames.Items(index), m_ColumnNames.Items(index) & Space(spaceNeeded)
        ElseIf index > column - 1 Then Exit For
        Else
            newDict.add m_ColumnNames.Items(index), m_ColumnNames.Items(index)
        End If
        
    Next index
    
    Set m_ColumnNames = newDict
End Sub
Public Sub SetColumns(newCollection As Collection)
    Set m_ColumnNames = AssignKeys(newCollection)
    
    m_ColumnCount = m_ColumnNames.count
    
    Set m_Records = New Collection
    m_IsColumnSet = True
End Sub
Public Property Get lengths() As Variant
    lengths = m_Lengths
End Property
Public Property Get isColumnSet() As Boolean
    isColumnSet = m_IsColumnSet
End Property
Public Property Get IsSeparatorset() As Boolean
    IsSeparatorset = m_IsSeparatorset
End Property
Public Property Let IsSeparatorset(IsSet As Boolean)
    m_IsSeparatorset = IsSet
End Property
Public Property Get Separators() As Variant
    Separators = m_Separators
End Property
Public Property Let Separators(SeparatorsArray As Variant)
    Dim i As Integer
    
    m_Separators = SeparatorsArray
    m_Lengths = Array()
    ReDim m_Lengths(UBound(m_Separators) - 1) As Variant
    
    
    If Not IsSeparatorset Then
        ''Only do this the first time
        For i = 0 To UBound(m_Separators) - 2
            ''Length = (The character in front of this separator) to
            ''         (The character before the next seperator)
            m_Lengths(i) = m_Separators(i + 1) - m_Separators(i) - 1
        Next i
        
    End If
    
    m_IsSeparatorset = True
End Property
Public Property Get ColumnCount() As Integer
    ColumnCount = m_ColumnCount
End Property
Public Function GetColumnName(columnIndex As Integer) As String
    If columnIndex > m_ColumnCount Or columnIndex < 1 Then Exit Function ''Return an empty string if out of bounds
    GetColumnName = Trim$(m_ColumnNames(columnIndex))
End Function
Public Function GetRecord(recordIndex As Integer) As String()
    Dim newArray() As String, i As Integer, record As Collection
    
    ReDim newArray(m_ColumnCount - 1) As String
    
    Set record = m_Records(recordIndex)
    
    For i = 1 To m_ColumnCount
        newArray(i - 1) = record(i)
    Next i
    
    GetRecord = newArray
End Function
Public Property Get recordCount() As Integer
    recordCount = m_RecordCount
End Property
Private Function AssignKeys(newCollection As Collection) As Dictionary
    Dim i As Integer, result As Dictionary
    On Error GoTo ErrorHandler
    Set result = New Dictionary
    result.CompareMode = BinaryCompare
    
    For i = 1 To newCollection.count
        result.add newCollection(i), newCollection(i)
    Next i
    
    Set AssignKeys = result
    
Exit Function
ErrorHandler:
    Stop
End Function
Public Sub TrimToColumns(newColumnNames() As String)
    Dim i As Integer, j As Integer, newColumns As Dictionary, newSeparators As Variant, newLengths As Variant
    
    On Error GoTo ErrorHandler
    Set newColumns = New Dictionary
    newColumns.CompareMode = BinaryCompare
    
    newSeparators = Array()
    newLengths = Array()
    ReDim newSeparators(UBound(newColumnNames) + 1) As Variant
    ReDim newLengths(UBound(newColumnNames)) As Variant
    
    For i = 0 To UBound(newColumnNames)
        For j = 0 To m_ColumnCount - 1
            If newColumnNames(i) Like ("*" & Trim(m_ColumnNames.Items(j)) & "*") Then
                newColumns.add m_ColumnNames.Items(j), m_ColumnNames.Items(j)
                newSeparators(i) = m_Separators(j)
                newLengths(i) = m_Lengths(j)
                Exit For
            End If
        Next j
    Next i
    
    newSeparators(UBound(newSeparators)) = newSeparators(UBound(newColumnNames)) + newLengths(UBound(newLengths))
    m_Separators = newSeparators
    m_Lengths = newLengths
    
    Set m_ColumnNames = newColumns
    m_ColumnCount = m_ColumnNames.count
Exit Sub
ErrorHandler:
    Stop
    Resume
End Sub
Public Function GetCursorReader(path As String, delimiter As String, requestedColumns() As String, query() As String) As CCursorReader
    If UBound(requestedColumns) = 0 And requestedColumns(0) = "" Then
        Set GetCursorReader = GetCursorReaderAllColumns(path, delimiter, query)
    Else
        Set GetCursorReader = GetCursorReaderSpecificColumns(path, delimiter, requestedColumns, query)
    End If
End Function
Private Function GetCursorReaderSpecificColumns(path As String, delimiter As String, requestedColumns() As String, query() As String) As CCursorReader
    ''INFO: Use this function to initialize the CCursorReader, or to return a new one.
    
    Dim tStream As TextStream, fso As FileSystemObject, line As String, newCursor As CCursorReader, blankArray(0) As String, isColumnSet As Boolean
    
    ''Temporary data to be put as method parameters/arguments later
    Dim time As Double
    time = Timer
    
    ''Get the TextStream, open for reading
    Set fso = New FileSystemObject
    Set tStream = fso.GetFile(path).OpenAsTextStream(ForReading)
    
    ''Start with an empty CCursorReader
    Set newCursor = New CCursorReader
    
    ''Read lines until the end of the stream, check the static IsColumnSet property of the CCursorReader each time to
        ''differentiate between the column headers and line data
    While Not tStream.AtEndOfStream
        line = tStream.ReadLine
        If Left(line, 1) = delimiter Then
            If Not isColumnSet Then
                
                ''Set the columns for the CCursorReader
                newCursor.SetColumns GetRecordFromText(line, delimiter, blankArray, newCursor)
                newCursor.TrimToColumns requestedColumns
                isColumnSet = newCursor.isColumnSet
                
            Else
            
                ''Add each line as a record to the CCursorReader
                
                newCursor.AddRecord GetRecordFromText(line, delimiter, query, newCursor)
                
            End If
        End If
    Wend
    tStream.Close
    Set GetCursorReaderSpecificColumns = newCursor
    
    Debug.Print "GetCursorReader took " & Timer - time & " seconds"
End Function
Private Function GetCursorReaderAllColumns(path As String, delimiter As String, query() As String) As CCursorReader
    ''INFO: Use this function to initialize the CCursorReader, or to return a new one.
    
    Dim tStream As TextStream, fso As FileSystemObject, line As String, newCursor As CCursorReader, blankArray(0) As String, isColumnSet As Boolean
    
    ''Temporary data to be put as method parameters/arguments later
    Dim time As Double
    time = Timer
    
    ''Get the TextStream, open for reading
    Set fso = New FileSystemObject
    Set tStream = fso.GetFile(path).OpenAsTextStream(ForReading)
    
    ''Start with an empty CCursorReader
    Set newCursor = New CCursorReader
    
    ''Read lines until the end of the stream, check the static IsColumnSet property of the CCursorReader each time to
        ''differentiate between the column headers and line data
    While Not tStream.AtEndOfStream
        line = tStream.ReadLine
        If Left(line, 1) = delimiter Then
            If Not isColumnSet Then
                
                ''Set the columns for the CCursorReader
                newCursor.SetColumns GetRecordFromText(line, delimiter, blankArray, newCursor)
                isColumnSet = newCursor.isColumnSet
                
            Else
            
                ''Add each line as a record to the CCursorReader
                
                newCursor.AddRecord GetRecordFromText(line, delimiter, query, newCursor)
                
            End If
        End If
    Wend
    tStream.Close
    Set GetCursorReaderAllColumns = newCursor
    
    Debug.Print "GetCursorReader took " & Timer - time & " seconds"
End Function
Public Function GetCursorReaderOnlyHeaders(path As String, delimiter As String, query() As String) As CCursorReader
    ''INFO: Use this function to initialize the CCursorReader, or to return a new one.
    
    Dim tStream As TextStream, fso As FileSystemObject, line As String, newCursor As CCursorReader, blankArray(0) As String, isColumnSet As Boolean
    
    ''Temporary data to be put as method parameters/arguments later
    Dim time As Double
    time = Timer
    
    ''Get the TextStream, open for reading
    Set fso = New FileSystemObject
    Set tStream = fso.GetFile(path).OpenAsTextStream(ForReading)
    
    ''Start with an empty CCursorReader
    Set newCursor = New CCursorReader
    
    ''Read lines until the end of the stream, check the static IsColumnSet property of the CCursorReader each time to
        ''differentiate between the column headers and line data
    Do While Not tStream.AtEndOfStream
        line = tStream.ReadLine
        If Left(line, 1) = delimiter Then
            If Not isColumnSet Then
                
                ''Set the columns for the CCursorReader
                newCursor.SetColumns GetRecordFromText(line, delimiter, blankArray, newCursor)
                isColumnSet = newCursor.isColumnSet
                Exit Do
            Else
            
                ''Add each line as a record to the CCursorReader
                
                newCursor.AddRecord GetRecordFromText(line, delimiter, query, newCursor)
                
            End If
        End If
    Loop
    tStream.Close
    Set GetCursorReaderOnlyHeaders = newCursor
    
    Debug.Print "GetCursorReader took " & Timer - time & " seconds"
End Function
Private Function GetRecordFromText(text As String, delimiter As String, queryMatch() As String, Optional rCursor As CCursorReader = Null, Optional addIndexColumn As Boolean = False, Optional index_ID As Integer = 0) As Collection
    Dim position As Integer, startString As Integer, endString As Integer, index As Integer, record As Collection, indexes As Collection, tempArray As Variant, tempSeparators As Variant, tempLengths As Variant
    
    On Error GoTo ErrorHandler
    
    Set record = New Collection
    If IsNull(rCursor) Then: Set rCursor = New CCursorReader
    
    If Not rCursor.IsSeparatorset Then
        Set indexes = New Collection
        
        For position = 1 To Len(text)
            If startString = 0 Then
                If Mid(text, position, 1) = delimiter Then
                    
                    startString = position
                    indexes.add position ''Save the position for later
                    
                End If
            Else
                If Mid(text, position, 1) = delimiter Then
                
                    endString = position
                    indexes.add position ''Save the position for later
                    
                    record.add Mid(text, startString + 1, endString - startString - 1)
                    startString = position
                End If
            End If
        Next position
        
        indexes.add Len(text)
        tempArray = PKLib.ToVarArray(indexes)
        rCursor.Separators = tempArray ''Box the Separators
        
    Else
    
        tempSeparators = rCursor.Separators ''Unbox the Separators
        tempLengths = rCursor.lengths ''Unbox the Lenghts
        
        For index = 0 To UBound(tempSeparators) - 1
            record.add Mid(text, tempSeparators(index) + 1, tempLengths(index))
        Next index
        
        If Not ExactMatch(queryMatch, record) Then: Exit Function
        If Not LikeMatch(queryMatch, record) Then: Exit Function
        If Not GreaterMatch(queryMatch, record) Then: Exit Function
        If Not LessMatch(queryMatch, record) Then: Exit Function
        
        If addIndexColumn Then: record.add index_ID
        
    End If
    
    Set GetRecordFromText = record
Exit Function
ErrorHandler:
    Stop
    Resume
    
End Function
Private Function ExactMatch(match() As String, record As Collection) As Boolean
    Dim i As Integer, check As String

    For i = 0 To UBound(match)
        If match(i) <> "" And Left(match(i), 1) = M_EXACT_IDENTIFIER Then
            If Trim(record(i + 1)) = "" Then: Exit Function
            check = Mid(match(i), 2, Len(match(i)) - 1)
            If Not Trim(record(i + 1)) Like check Then: Exit Function
        End If
    Next i
    ExactMatch = True
End Function
Private Function LikeMatch(match() As String, record As Collection) As Boolean
    Dim i As Integer, check As String

    For i = 0 To UBound(match)
        If match(i) <> "" And Left(match(i), 1) = M_LIKE_IDENTIFIER Then
            If Trim(record(i + 1)) = "" Then: Exit Function
            check = "*" & Mid(match(i), 2, Len(match(i)) - 1) & "*"
            If Not Trim(record(i + 1)) Like check Then: Exit Function
        End If
    Next i
    LikeMatch = True
End Function
Private Function GreaterMatch(match() As String, record As Collection) As Boolean
    Dim i As Integer, check As String, inVal As Variant, compVal As Variant
    
    For i = 0 To UBound(match)
        If match(i) <> "" And Left(match(i), 1) = M_GREATER_IDENTIFIER Then
            If Trim(record(i + 1)) = "" Then: Exit Function
            inVal = Trim(record(i + 1))
            compVal = Mid(match(i), 2, Len(match(i)) - 1)
            If IsDate(inVal) Then: inVal = CDate(inVal)
            If IsNumeric(inVal) Then: inVal = CLng(inVal)
            If IsDate(compVal) Then: compVal = CDate(compVal)
            If IsNumeric(compVal) Then: compVal = CLng(compVal)
            
            If Not inVal > compVal Then: Exit Function
        End If
    Next i
    
    GreaterMatch = True
End Function
Private Function LessMatch(match() As String, record As Collection) As Boolean
    Dim i As Integer, check As Integer, inVal As Variant, compVal As Variant
    
    For i = 0 To UBound(match)
        If match(i) <> "" And Left(match(i), 1) = M_LESS_IDENTIFIER Then
            If Trim(record(i + 1)) = "" Then: Exit Function
            inVal = Trim(record(i + 1))
            compVal = Mid(match(i), 2, Len(match(i)) - 1)
            If IsDate(inVal) Then: inVal = CDate(inVal)
            If IsNumeric(inVal) Then: inVal = CLng(inVal)
            If IsDate(compVal) Then: compVal = CDate(compVal)
            If IsNumeric(compVal) Then: compVal = CLng(compVal)
            
            If Not inVal < compVal Then: Exit Function
        End If
    Next i
    
    LessMatch = True
End Function
Private Sub SetWidths(ByVal line As String, delimiter As String)
    Dim index As Integer, indexes As Collection, newArray As Variant
    
    Set indexes = New Collection
    index = 1
    
    
        Do While InStr(index, line, delimiter) > 0
            For index = index To Len(line)
                    
                If Mid(line, index, 1) = delimiter Then
                    indexes.add index
                    index = index + 1
                    Exit For
                End If
                    
            Next index
        Loop
    
    newArray = PKLib.ToVarArray(indexes)
    
    
    ''Copy over the largest widths only
    For index = 0 To UBound(m_Separators)
    
        If newArray(index) > m_Separators(index) Then: m_Separators(index) = newArray(index)
    
    Next index
    
End Sub

