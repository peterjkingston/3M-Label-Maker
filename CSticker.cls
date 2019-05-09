VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim sticker As New CSticker
Public CustomerName As String, SalesOrderNumber As String, CSName As String
Public Property Get Body()

    Body = CustomerName & vbCrLf & _
           SalesOrderNumber & vbCrLf & _
           CSName
           
End Property
