VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCursorWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim myCursor As New CCursorWriter
Private m_RecordCount As Integer
Private m_ColumnNames() As String, m_Records2D() As String
Private m_Lengths() As Integer
Public Function GetCursorWriter(fromReader As CCursorReader) As CCursorWriter
    m_Records2D = fromReader.GetRecords2DArray
    
    If Not m_Records2D(0, 0) = "" Then
        m_RecordCount = UBound(m_Records2D) + 1
        m_ColumnNames = fromReader.GetColumnNames
        
        m_Lengths = GetLengths
        
        Set GetCursorWriter = Me
    Else
        Set GetCursorWriter = Nothing
    End If
End Function
Private Function GetLengths() As Integer()
    Dim column As Integer, row As Integer, lengths() As Integer
    
    ReDim lengths(UBound(m_ColumnNames)) As Integer
    
    ''Grab the length from each column
    For column = 0 To UBound(lengths)
        lengths(column) = Len(m_ColumnNames(column))
    Next column
    
    ''For each row, check each data item
    For row = 0 To m_RecordCount - 1
        For column = 0 To UBound(m_ColumnNames)
            
            ''If the data is longer than the existing length for that column, assign it_
            ''the new length.
            If Len(m_Records2D(row, column)) > lengths(column) Then
                lengths(column) = Len(m_Records2D(row, column))
            End If
            
        Next column
    Next row
    
    GetLengths = lengths
End Function
Public Sub WriteToFile(path As String, delimiter As String)
    Dim fso As FileSystemObject, tStream As TextStream, row As Integer
    
    Set fso = New FileSystemObject
    Set tStream = fso.OpenTextFile(path, ForWriting, True)
    
    tStream.WriteLine CreateLine(m_ColumnNames, delimiter)
    
    For row = 0 To m_RecordCount - 1
        tStream.WriteLine CreateLine(GetRecord(row), delimiter)
    Next row
    
    tStream.Close
End Sub
Private Function CreateLine(withStrings() As String, delimiter As String)
    Dim column As Integer, newString As String
    
    newString = delimiter
    For column = 0 To UBound(withStrings)
        newString = newString & _
                    withStrings(column) & _
                    Space(m_Lengths(column) - Len(withStrings(column))) & _
                    delimiter
    Next column
    
    CreateLine = newString
    
End Function
Private Function GetRecord(row As Integer) As String()
    GetRecord = record(row)
End Function
Private Property Get record(row As Integer) As String()
    Dim strArry() As String, column As Integer
    ReDim strArry(UBound(m_Records2D, 2)) As String
    
    For column = 0 To UBound(strArry)
        strArry(column) = m_Records2D(row, column)
    Next column
    record = strArry
End Property
