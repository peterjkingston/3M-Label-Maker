Attribute VB_Name = "Deploy"
Const DELM As String = "|"

''File Extensions
Const FSL As String = "\"
Const FTXT As String = ".txt"
Const FBMP As String = ".bmp"
Const FJPG As String = ".jpg"

'' Manifest
Const LBL_MAIN As String = "Main"
Const LBL_QRY As String = "Queries"
Const LBL_NFIX As String = "NameFix"
Const LBL_1000 As String = "USOrders"
Const LBL_2000 As String = "CAOrders"
Const LBL_COG_JPG As String = "SettingsCog"

''Queries
Const LBL_USTAB As String = "USData"
Const LBL_CADATA As String = "CAData"
Const LBL_CORDICT As String = "CorrectionDict"

Public Sub Manifest()
    Dim fso As FileSystemObject, rqDir As String, tStream As TextStream, dirPath As String
    Dim widths(1) As Integer, labels(1) As String
    
    Set fso = New FileSystemObject
    rqDir = ThisWorkbook.path & "\AppManifest.txt"
    dirPath = ThisWorkbook.path
    
    widths(0) = 50
    widths(1) = Len(ThisWorkbook.FullName) + 30
    
    If Not fso.FileExists(rqDir) Then
    ''Caller should determine if this action should happen or not
    
        ''Create a shell AppManifest at this directory
        Set tStream = fso.OpenTextFile(rqDir, ForWriting, True)
        
        ''Column Headers
        labels(0) = "Object_Name"
        labels(1) = "File_Path"
        tStream.WriteLine WriteTableLine(labels, widths)

        ''Main object (this)
        labels(0) = LBL_MAIN
        labels(1) = ThisWorkbook.FullName
        tStream.WriteLine WriteTableLine(labels, widths)
               
        ''Queries
        labels(0) = LBL_QRY
        labels(1) = dirPath & FSL & LBL_QRY & FTXT
        tStream.WriteLine WriteTableLine(labels, widths)
               
        ''Name Fixer
        labels(0) = LBL_NFIX
        labels(1) = dirPath & FSL & LBL_NFIX & FTXT
        tStream.WriteLine WriteTableLine(labels, widths)
                          
        ''SAP region 1000 order table
        labels(0) = LBL_1000
        labels(1) = dirPath & FSL & LBL_1000 & FTXT
        tStream.WriteLine WriteTableLine(labels, widths)
        
        ''SAP region 2000 order table
        labels(0) = LBL_2000
        labels(1) = dirPath & FSL & LBL_2000 & FTXT
        tStream.WriteLine WriteTableLine(labels, widths)
                          
        ''Graphic 1: Settings Cog
        labels(0) = LBL_COG_JPG
        labels(1) = dirPath & FSL & LBL_COG_JPG & FJPG
        tStream.WriteLine WriteTableLine(labels, widths)
        
        tStream.Close
    End If
End Sub

Public Sub queries()
    Dim fso As FileSystemObject, rqDir As String, tStream As TextStream
    Dim widths(5) As Integer, labels(5) As String, i As Integer
    
    Set fso = New FileSystemObject
    rqDir = ThisWorkbook.path & "\Queries.txt"
    
    For i = 0 To 5
        widths(i) = 30
    Next i
    
    If Not fso.FileExists(rqDir) Then
        ''Create a shell Queries at this directory
        Set tStream = fso.OpenTextFile(rqDir, ForWriting, True)
        
        ''Column Headers
        labels(0) = "QName"
        labels(1) = "Arg1"
        labels(2) = "Arg2"
        labels(3) = "Arg3"
        labels(4) = "Arg4"
        labels(5) = "Arg5"
        tStream.WriteLine WriteTableLine(labels, widths)
        
        ''SAP region 1000 query
        labels(0) = "USData"
        labels(1) = "Document"
        labels(2) = "Name1"
        labels(3) = "Created"
        labels(4) = "Sold-to pt"
        labels(5) = "Purchase order number"
        tStream.WriteLine WriteTableLine(labels, widths)
        
        ''SAP region 2000 query
        labels(0) = "CAData"
        labels(1) = "Document"
        labels(2) = "Name1"
        labels(3) = "Created"
        labels(4) = "Sold-to pt"
        labels(5) = "PO number"
        tStream.WriteLine WriteTableLine(labels, widths)
        
        ''User correction dictionary query
        labels(0) = "CorrectionDict"
        labels(1) = "Sold-to pt"
        labels(2) = "Name1"
        labels(3) = ""
        labels(4) = ""
        labels(5) = ""
        tStream.WriteLine WriteTableLine(labels, widths)
        
        tStream.Close
    End If
End Sub

Public Sub SAP1000()
    Dim fso As FileSystemObject, rqDir As String, tStream As TextStream
    Dim widths(0) As Integer, labels(0) As String, i As Integer
    
    Set fso = New FileSystemObject
    rqDir = ThisWorkbook.path & "\" & LBL_1000 & FTXT
    
    If Not fso.FileExists(rqDir) Then
        ''Create a shell SAP1000 at this directory: Nothing needs to be in the file
        Set tStream = fso.OpenTextFile(rqDir, ForWriting, True)
        tStream.Close
    End If
End Sub

Public Sub SAP2000()
    Dim fso As FileSystemObject, rqDir As String, tStream As TextStream
    Dim widths(0) As Integer, labels(0) As String, i As Integer
    
    Set fso = New FileSystemObject
    rqDir = ThisWorkbook.path & "\" & LBL_2000 & FTXT
    
    If Not fso.FileExists(rqDir) Then
        ''Create a shell SAP2000 at this directory: Nothing needs to be in the file
        Set tStream = fso.OpenTextFile(rqDir, ForWriting, True)
        tStream.Close
    End If
End Sub

Public Sub All()
    Manifest
    queries
    SAP1000
    SAP2000
End Sub

Public Function GetBuildFileNames() As String()
    Dim rqDir As String
    Dim labels(5) As String
    
    rqDir = ThisWorkbook.path
    
    labels(0) = rqDir & FSL & "App"
    labels(1) = "Sold-to pt"
    labels(2) = "Name1"
    labels(3) = ""
    labels(4) = ""
    labels(5) = ""
    
    GetBuildFileNames = labels
End Function

Private Function WriteTableLine(columns() As String, widths() As Integer) As String
    If UBound(columns) <> UBound(widths) Then: Exit Function
    Dim i As Integer, buildStr As String
    
    For i = 0 To UBound(columns)
        buildStr = buildStr & DELM & columns(i) & Space(widths(i) - Len(columns(i)))
    Next i
    
    WriteTableLine = buildStr & DELM
End Function

