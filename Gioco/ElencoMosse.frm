VERSION 5.00
Begin VB.Form ElencoMosse 
   Caption         =   "Mosse"
   ClientHeight    =   7380
   ClientLeft      =   990
   ClientTop       =   1530
   ClientWidth     =   2190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Doppio click per tornare su un'altra mossa"
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
End
Attribute VB_Name = "ElencoMosse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Dim AP As Variant
Dim cA As String * 1
Dim cB As String * 1
Dim cC As String * 1
Dim cD As String * 1
Dim cE As String * 1
Dim cF As String * 1
Dim cG As String * 1
Dim cH As String * 1
Dim r1 As String * 1
Dim r2 As String * 1
Dim r3 As String * 1
Dim r4 As String * 1
Dim r5 As String * 1
Dim r6 As String * 1
Dim r7 As String * 1
Dim r8 As String * 1
Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
 cA = "A"
 cB = "B"
 cC = "C"
 cD = "D"
 cE = "E"
 cF = "F"
 cG = "G"
 cH = "H"
 r1 = "1"
 r2 = "2"
 r3 = "3"
 r4 = "4"
 r5 = "5"
 r6 = "6"
 r7 = "7"
 r8 = "8"
End Sub

Private Sub Form_Resize()
  Dim h As Integer
  h = ScaleHeight - (List.Height + 25)
  If h < 1 Then
    h = 1
  End If
  Text1.Move 0, (List.Height + 25), ScaleWidth, h
  List.Width = ScaleWidth
End Sub

Private Sub Form_Terminate()
  Set AP = Nothing
End Sub

Private Sub List_DblClick()
  If AP.EsisteGiocoSelezionato = True Then
    AP.VaiAlla List.ListIndex + 1
  End If
End Sub
Public Sub Aggiorna()
  Static UnivocoFinestra As String
  Dim N As Integer
  Dim s As Integer
  Dim SC As Variant
  Dim Traccia As Variant
  Dim St As String
  If MDIScacchi.ActiveForm.NomeFile <> UnivocoFinestra Then
    Svuota
    UnivocoFinestra = MDIScacchi.ActiveForm.NomeFile
  End If
  Set SC = MDIScacchi.ActiveForm.Scacchiera
  Traccia = SC.Traccia
  Do While List.ListCount > SC.Quantit‡MosseTraccia
    List.RemoveItem (List.ListCount - 1)
  Loop
  If AP.EsisteGiocoSelezionato = False Then
    Svuota
  Else
    s = List.ListCount
    'elimina le mosse che non devono essere visualizzate
    'If List.ListCount <> 0 Then
    '  For N = List.ListCount To SC.MossaAttuale Step -1
    '    List.RemoveItem (N - 1)
    '  Next
    'End If
    'aggiungi le mosse che mancano
    If SC.Quantit‡MosseTraccia > 0 Then
      For N = List.ListCount + 1 To SC.Quantit‡MosseTraccia
        St = Str(Int((N + 1) / 2)) + ") " + TxtCoo(Traccia(N).Provenienza.Posizione.x, Traccia(N).Provenienza.Posizione.Y) + " - " + TxtCoo(Traccia(N).Destinazione.Posizione.x, Traccia(N).Destinazione.Posizione.Y)
        List.AddItem St
      Next
      List.ListIndex = SC.MossaAttuale - 1
    Else
      Text1.Text = ""
    End If
  End If
End Sub
Public Sub Svuota()
  While List.ListCount
    List.RemoveItem 0
  Wend
End Sub
Private Function TxtCoo(x As Colonna, Y As Riga) As String
  Dim T As String
  Dim T2 As String
  Select Case x
    Case 0
      T = cA
    Case 1
      T = cB
    Case 2
      T = cC
    Case 3
      T = cD
    Case 4
      T = cE
    Case 5
      T = cF
    Case 6
      T = cG
    Case 7
      T = cH
  End Select
  Select Case Y
    Case 0
      T2 = r1
    Case 1
      T2 = r2
    Case 2
      T2 = r3
    Case 3
      T2 = r4
    Case 4
      T2 = r5
    Case 5
      T2 = r6
    Case 6
      T2 = r7
    Case 7
      T2 = r8
  End Select
  TxtCoo = T + T2
End Function

Private Sub List_Click()
  Dim Traccia As Variant
  Dim SC As Variant
  Set SC = MDIScacchi.ActiveForm.Scacchiera
  Traccia = SC.Traccia
  Text1.Text = Traccia(List.ListIndex + 1).Nota
End Sub
Private Sub Text1_Change()
  Dim SC As Variant
  Dim Traccia As Variant
  Set SC = MDIScacchi.ActiveForm.Scacchiera
  If (List.ListIndex + 1) <> 0 Then
    Traccia = SC.Traccia
    Traccia(List.ListIndex + 1).Nota = Text1.Text
    SC.Traccia = Traccia
  End If
End Sub
