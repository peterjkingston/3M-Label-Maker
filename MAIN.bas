Attribute VB_Name = "MAIN"
Option Explicit
Public Program As CProgram_MAIN
Public Sub Main()
    Set Program = New CProgram_MAIN
    
    Program.Main
    
    Set Program = Nothing
End Sub
