VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrintablePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Page As New CPrintablePage
Public Printables As New Collection
Public Function TrueCount() As Integer
    Dim printable As CPrintable, count As Integer
    
    For Each printable In Printables
        If printable.IsPrintable Then: count = count + 1
    Next
    
    If count > 18 Then: count = 18
    TrueCount = count
End Function
Public Function FalseCount() As Integer
    Dim printable As CPrintable, count As Integer
    
    For Each printable In Printables
        If Not printable.IsPrintable Then: count = count + 1
    Next
    
    FalseCount = count
End Function
