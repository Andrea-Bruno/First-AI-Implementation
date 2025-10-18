Attribute VB_Name = "ModuloGioco"
DefInt A-Z
Option Explicit
Dim Application As New Application
Sub Main()
  MDIScacchi.Show
End Sub
Function AP() As Object
  Set AP = Application
End Function
Sub Terminate()
    Set Application = Nothing
End Sub
