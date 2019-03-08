Attribute VB_Name = "PKLib"
Public Sub Wait(seconds As Double)
    Dim time As Double
    
    time = Timer + seconds
    
    Do While Timer < time
    Loop
    
End Sub
Public Function ToVarArray(list As Collection) As Variant
    Dim newArray As Variant, i As Integer
    
    ReDim newArray(list.count - 1) As Variant
    
    For i = 1 To list.count
        newArray(i - 1) = list(i)
    Next i
    
    ToVarArray = newArray
End Function
Public Function ToStrArray(list As Collection) As String()
    Dim newArray() As String, i As Integer
    
    ReDim newArray(list.count - 1) As String
    
    For i = 1 To list.count
        newArray(i - 1) = list(i)
    Next i
    
    ToStrArray = newArray
End Function
Public Function ToStrArray2D(list2D As Collection) As String()
    Dim newArray() As String, r As Integer, isCollectionArray As Boolean, c As Integer, startCol As Integer, offset As Integer
    
    isCollectionArray = isArray(list2D(1))
    
    If isCollectionArray Then
        endCol = UBound(list2D(1))
        ReDim newArray(list2D.count - 1, endCol) As String
        startCol = 0
        offset = 0
    Else
        endCol = list2D(1).count - 1
        ReDim newArray(list2D.count - 1, endCol) As String
        startCol = 1
        offset = 1
    End If
    
    For r = 1 To list2D.count
        For c = startCol To endCol
            newArray(r - 1, c - offset) = list2D(r)(c)
        Next c
    Next r
    
    ToStrArray2D = newArray
End Function
Public Function GetQueryHeaders(ws As Worksheet) As Collection
    Dim listObj As ListObject, listCol As ListColumn, result As Collection
    
    Set result = New Collection
    Set listObj = ws.ListObjects(1)
    
    For Each listCol In listObj.ListColumns
        result.Add Trim(listCol.Name)
    Next listCol
    
    Set GetQueryHeaders = result
End Function
Public Function GetQueryOperators() As Collection
    Dim result As Collection
    
    result.Add "="
    result.Add "CONTAINS"
    result.Add "MATCH COLUMN"
    
    Set GetQueryOperators = result
End Function
Public Function GetSubArrayStr(strArray() As String, Optional startRow As Integer = 1, Optional endRow As Integer = 1, Optional startCol As Integer = 1, Optional endCol As Integer = 1) As String()
    Dim newArray() As String, r As Integer, c As Integer, invalid As Boolean
    
    Select Case True
        Case endRow < startRow:                 invalid = True
        Case endRow >= UBound(strArray, 0):     invalid = True
        Case endCol < startCol:                 invalid = True
        Case endCol >= UBound(strArray, 1):     invalid = True
        Case startRow < 0:                      invalid = True
        Case startCol < 0:                      invalid = True
    End Select
    
    If invalid Then: Exit Function
    
    ReDim newArray(endRow - startRow + 1, endCol - startCol + 1) As String
    
    For startRow = startRow To endRow
        c = 0
        For startCol = startCol To endCol
            newArray(r, c) = strArray(startRow, startCol)
        Next startCol
        r = r + 1
    Next startRow
    
    GetSubArrayStr = newArray
End Function
