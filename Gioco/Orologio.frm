VERSION 5.00
Begin VB.Form Orologio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orologio"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1965
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   1965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1440
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Incremento in secondi del tempo disponibile a ogni mossa effettuata"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "In quante mosse bisogna utilizzare il tempo stabilito; se il valore è 0 i minuti di riflessione si riferiranno all'intera partita"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Tempo di riflessione in minuti per effettuare il numero di mosse stabilito o l'intera partita"
      Top             =   720
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Label Label3 
      Caption         =   "In mosse"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Secondi bonus"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Minuti riflessione"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label OrologioNero 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label OrologioBianco 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Nero"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bianco"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   960
      Y1              =   120
      Y2              =   720
   End
End
Attribute VB_Name = "Orologio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Msg1 = "tepo di piflessione scaduto!"
Const MsgNero = "Nero"
Const MsgBianco = "Bianco"
Sub Aggiorna(FormGioco As Gioco)
  Dim Turno As Colore
  Dim Traccia As Variant
  Turno = ProssimoTurno(FormGioco.Scacchiera.Editabilità)
  If (Turno = Bianco Or Turno = Nero) And FormGioco.Scacchiera.EventoOra <> 0 And FormGioco.Scacchiera.MossaAttuale <> 0 Then
    Dim TB As Variant
    Dim TN As Variant
    Traccia = FormGioco.Scacchiera.Traccia
    If Turno = Bianco Then
      TB = Now - FormGioco.Scacchiera.EventoOra - Traccia(FormGioco.Scacchiera.MossaAttuale).Tempo
      TN = Traccia(FormGioco.Scacchiera.MossaAttuale).Tempo
    Else
      TN = Now - FormGioco.Scacchiera.EventoOra - Traccia(FormGioco.Scacchiera.MossaAttuale).Tempo
      TB = Traccia(FormGioco.Scacchiera.MossaAttuale).Tempo
    End If
    On Local Error Resume Next
    If FormGioco.TempoDiRiflessione Or FormGioco.TempoBonusPerMossa Then
      'il tempo a disposizione si calcola così:
      'TempoDiRiflessione * 1+((MossaAttuale-1)/MossePerTempo)+TempoBonusPerMossa*MossaAttuale
      'La mossaattuale si calcola con Int(SemiMossa/2)
      Dim TempoDisponibile As Long
      TempoDisponibile = (FormGioco.TempoDiRiflessione * 60)
      If FormGioco.MossePerTempo Then
        'Moltiplica per MossePerTempo
        Dim MosseEffettuate As Long
        MosseEffettuate = 1 + Int(FormGioco.Scacchiera.MossaAttuale / 2)
        TempoDisponibile = TempoDisponibile * (1 + Int((MosseEffettuate - 1) / FormGioco.MossePerTempo))
      End If
      If FormGioco.TempoBonusPerMossa Then
        'Aggiungi il bonus per ogni mossa
        TempoDisponibile = TempoDisponibile + FormGioco.TempoBonusPerMossa * Int(FormGioco.Scacchiera.MossaAttuale / 2)
      End If
      If Turno = Bianco Then
        FormGioco.TempoDisponibileBianco = TempoDisponibile
      ElseIf Turno = Nero Then
        FormGioco.TempoDisponibileNero = TempoDisponibile
      End If
      TB = (FormGioco.TempoDisponibileBianco / 86400) - TB
      TN = (FormGioco.TempoDisponibileNero / 86400) - TN
      If TB >= 0 Then
        If OrologioBianco.ForeColor <> &H80000012 Then
          OrologioBianco.ForeColor = &H80000012
        End If
      Else
        If OrologioBianco.ForeColor <> &H8000000E Then
          OrologioBianco.ForeColor = &H8000000E
        End If
        If (FormGioco.BiancoTempoScaduto = False) And (Turno = Bianco) Then
          'Tempo scaduto al bianco
           FormGioco.BiancoTempoScaduto = True
          MsgBox MsgBianco & " " & Msg1, vbInformation
        End If
      End If
      If TN >= 0 Then
        If OrologioNero.ForeColor <> &H80000012 Then
          OrologioNero.ForeColor = &H80000012
        End If
      Else
        If OrologioNero.ForeColor <> &H8000000E Then
          OrologioNero.ForeColor = &H8000000E
        End If
        If (FormGioco.NeroTempoScaduto = False) And (Turno = Nero) Then
          'Tempo scaduto al nero
           FormGioco.NeroTempoScaduto = True
          MsgBox MsgNero & " " & Msg1, vbInformation
        End If
      End If
    Else
      If OrologioBianco.ForeColor <> &H80000012 Then
        OrologioBianco.ForeColor = &H80000012
      End If
      If OrologioNero.ForeColor <> &H80000012 Then
        OrologioNero.ForeColor = &H80000012
      End If
    End If
    OrologioBianco.Caption = Format(TB, "HH:MM:SS")
    OrologioNero.Caption = Format(TN, "HH:MM:SS")
  Else
    OrologioBianco.Caption = 0: OrologioNero.Caption = 0
  End If
End Sub

Public Sub AggiornaLimitiTempo()
  If Text1(0).Text <> LTrim(Str(MDIScacchi.ActiveForm.TempoDiRiflessione)) Then
    Text1(0).Text = LTrim(Str(MDIScacchi.ActiveForm.TempoDiRiflessione))
  End If
  If Text1(1).Text <> LTrim(Str(MDIScacchi.ActiveForm.MossePerTempo)) Then
    Text1(1).Text = LTrim(Str(MDIScacchi.ActiveForm.MossePerTempo))
  End If
  If Text1(2).Text <> LTrim(Str(MDIScacchi.ActiveForm.TempoBonusPerMossa)) Then
    Text1(2).Text = LTrim(Str(MDIScacchi.ActiveForm.TempoBonusPerMossa))
  End If
End Sub

Private Sub Form_Load()
  AggiornaLimitiTempo
End Sub

Private Sub Text1_Change(Index As Integer)
  On Local Error Resume Next
  If Index = 0 Then
    MDIScacchi.ActiveForm.TempoDiRiflessione = Abs(Val(Text1(Index).Text))
  ElseIf Index = 1 Then
    MDIScacchi.ActiveForm.MossePerTempo = Abs(Val(Text1(Index).Text))
  Else
    MDIScacchi.ActiveForm.TempoBonusPerMossa = Abs(Val(Text1(Index).Text))
  End If
  Timer2.Enabled = True
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  On Local Error Resume Next
  AggiornaLimitiTempo
End Sub

Private Sub Timer1_Timer()
  On Local Error GoTo errtim
  Aggiorna MDIScacchi.ActiveForm
  Exit Sub
errtim:
  OrologioBianco.Caption = 0: OrologioNero.Caption = 0
End Sub
Function ProssimoTurno(EditScacchiera As EditorScacchiera) As Colore
  Select Case EditScacchiera
    Case BiancoMossePossibili
      ProssimoTurno = Bianco
    Case NeroMossePossibili
      ProssimoTurno = Nero
    Case Else
      ProssimoTurno = Nessun_Colore
  End Select
End Function

Private Sub Timer2_Timer()
  Timer2.Enabled = False
  MDIScacchi.ActiveForm.NeroTempoScaduto = False
  MDIScacchi.ActiveForm.BiancoTempoScaduto = False
End Sub
