VERSION 5.00
Begin VB.Form Pensiero 
   Caption         =   "Pensiero"
   ClientHeight    =   3750
   ClientLeft      =   300
   ClientTop       =   3420
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4770
   Begin VB.Label Label 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Pensiero del motore d'analisi della mossa in elaborazione"
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Pensiero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Public Function Aggiorna(FormGioco As Gioco)
  Label.Caption = FormGioco.Commento
End Function
Private Sub Form_Resize()
  Dim h As Integer
  Label.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
