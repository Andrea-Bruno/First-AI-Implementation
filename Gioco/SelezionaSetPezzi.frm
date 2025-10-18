VERSION 5.00
Begin VB.Form SelezionaSetPezzi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambia il set delle figure"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Fai doppio click sul set di figure che intendi utilizzare"
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "SelezionaSetPezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Dim AP As Variant
Option Explicit
Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
End Sub
Private Sub Form_Load()
  Dim PthPz As String
  PthPz = MDIScacchi.Casa(0).Path + "\pictures\figure"
  Dim s As String
  s = Dir(PthPz + "\*", 16)
  If s <> "" Then
  Do
    If Left(s, 1) <> "." Then
      List1.AddItem s
    End If
    s = Dir(, 16)
  Loop Until s = ""
  End If
End Sub

Private Sub Form_Terminate()
  Set AP = Nothing
End Sub

Private Sub List1_DblClick()
  AP.CaricaFigure (MDIScacchi.Casa(0).Path + "\pictures\figure\" + List1.List(List1.ListIndex))
  Unload Me
End Sub
