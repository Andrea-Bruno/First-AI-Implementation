VERSION 5.00
Object = "{706DFD08-82E6-11D1-AD79-F03E07C10000}#96.0#0"; "ControlloScacchiera.ocx"
Begin VB.Form Gioco 
   ClientHeight    =   2355
   ClientLeft      =   -3015
   ClientTop       =   -2190
   ClientWidth     =   3120
   ClipControls    =   0   'False
   Icon            =   "Gioco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2355
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ControlloScacchiera.Scacchiera Scacchiera 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3836
      TipoDiSelezione =   3
      ProvenienzaPezzo0=   0
      ProvenienzaPosizioneX0=   0
      ProvenienzaPosizioneY0=   0
      DestinazionePezzo0=   0
      DestinazionePosizioneX0=   0
      DestinazionePosizioneY0=   0
      BiancoArroccoPossibile0=   0
      NeroArroccoPossibile0=   0
      XEnPassant0     =   0
      Editabilit‡0    =   0
      Nota0           =   ""
      Collegamento0   =   ""
      Tempo0          =   0
      Evento          =   ""
      EventoLuogo     =   ""
      EventoData      =   0
      EventoOra       =   0
      EventoSponsor   =   ""
      EventoTipoGara  =   0
      EventoSessione  =   0
      EventoTurno     =   0
      EventoNumeroScacchiera=   0
      EventoIncontro  =   0
      EventoRisultato =   0
      EventoApertura  =   ""
      EventoSottoApertura=   ""
      EventoSottoSottoApertura=   ""
      EventoECO       =   ""
      EventoNIC       =   ""
      EventoTerminato =   0
      EventoNomeAnnotatore=   ""
      EventoModo      =   0
      EventoExtra     =   ""
      BiancoGiocatore =   ""
      NeroGiocatore   =   ""
      BiancoElo       =   0
      NeroElo         =   0
      BiancoUSCF      =   ""
      NeroUSCF        =   ""
      BiancoTitolo    =   0
      NeroTitolo      =   0
      BiancoEm        =   ""
      NeroEm          =   ""
      BiancoTipoGiocatore=   0
      NeroTipoGiocatore=   0
      BiancoNazionalit‡=   ""
      NeroNazionalit‡ =   ""
      MossaAttuale    =   0
      Quantit‡Pezzi   =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2400
      Top             =   120
   End
End
Attribute VB_Name = "Gioco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Dim AP As Variant
Dim m_NomeFile As String
Const Msg1 = "La partita si Ë conclusa!"
Const Msg2 = "Seleziona il pezzo promosso"
Public AggiornamentiAbilitati As Boolean
Public AspettaSelezionePezzoPromosso As Boolean
Public Editabilit‡DurantePromozione As EditorScacchiera
Public NeroTempoScaduto As Boolean
Public BiancoTempoScaduto
'Variabili utilizzate per impostare i limiti di tempo nell'orologio
Public TempoDiRiflessione As Integer
Public MossePerTempo As Integer
Public TempoBonusPerMossa As Integer
Public TempoDisponibileBianco As Long
Public TempoDisponibileNero As Long
Public CalcoloMossaInCorso As Boolean 'Usata per evitare di fargli calcolare una mossa quando la sta gia calcolando
Public Commento As String 'Usata per memorizzare l'output del motore d'elaborazione quando sta calcolando
Public Motore As String
Private Sub Form_Initialize()
   Scacchiera.EventoData = Date
  Set AP = ModuloGioco.AP
  AggiornamentiAbilitati = True
End Sub
Public Property Get NomeFile() As String
    NomeFile = m_NomeFile
End Property
Public Property Let NomeFile(ByVal New_NomeFile As String)
    m_NomeFile = New_NomeFile
End Property
Private Sub Form_Activate()
  Motore = AP.Path + "\Motori\" + Dir(AP.Path + "\Motori\*.exe")
  If AggiornamentiAbilitati = True Then
    AP.AggiornaFinestre
    Orologio.AggiornaLimitiTempo
    AggiornaStatoElaborazione
  End If
End Sub
Private Sub Form_Load()
  Left = -50000
'0K
End Sub

Private Sub Form_Resize()
  If Left >= 0 And Top >= 0 Then
    Scacchiera.Move 0, 0, ScaleWidth, ScaleHeight
    Scacchiera.Visible = True
  End If
End Sub
Public Sub Colloca()
  Dim Max As Long
  Max = MDIScacchi.Width \ 2
  If MDIScacchi.Height \ 2 > Max Then
    Max = MDIScacchi.Height \ 2
  End If
  On Local Error Resume Next
  Move MDIScacchi.Width \ 4, 0, Max, Max
End Sub
Private Sub Form_Terminate()
  AP.AggiornaFinestre
  Set AP = Nothing
End Sub
Private Sub Scacchiera_CambiaQuantit‡MosseTraccia(MosseTraccia As Integer)
  If AggiornamentiAbilitati = True Then
'    AP.AggiornaElencoMosse
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  AspettaMossaElaborata (False)
End Sub
Private Sub Scacchiera_CambiatoMossaAttuale(MossaAttuale As Integer)
  If AggiornamentiAbilitati = True Then
    AP.AggiornaElencoMosse
    AP.AggiornaEditabilit‡
  End If
End Sub
Public Sub MossaEffettuata()
  DoEvents
  If AspettaSelezionePezzoPromosso = False Then
    If AggiornamentiAbilitati = True Then
      AP.AggiornaEditabilit‡
      AP.AggiornaAnalisi
    End If
    If ((Scacchiera.Editabilit‡ = BiancoMossePossibili) And (Computer = Scacchiera.BiancoTipoGiocatore)) Or ((Scacchiera.Editabilit‡ = NeroMossePossibili) And (Computer = Scacchiera.NeroTipoGiocatore)) Then
      Muovi
    End If
    If Scacchiera.Visible = False Then
      Unload Me
      Exit Sub
    End If
    If Scacchiera.ScaccoMatto Then
      Scacchiera.Editabilit‡ = NonEditabile
      MsgBox Msg1, vbInformation
    End If
  End If
End Sub
Private Sub Scacchiera_PezzoPosizionato(pezzo As ControlloScacchiera.pezzo, Provenienza As ControlloScacchiera.IndexCasa, Destinazione As ControlloScacchiera.IndexCasa)
  MossaEffettuata
End Sub
Public Sub Muovi()
  If CalcoloMossaInCorso = False Then
    CalcoloMossaInCorso = True
    Dim Turno As Colore
    If Scacchiera.Editabilit‡ = BiancoMossePossibili Then
      Turno = Bianco
    Else
      Turno = Nero
    End If
    Dim TempoRimanente As Long 'Tempo rimanente in s/100
    If Scacchiera.MossaAttuale = 0 Then
      TempoRimanente = (TempoDiRiflessione * 6000#)
    Else
      TempoRimanente = (TempoDiRiflessione * 6000#) - (Scacchiera.Traccia(Scacchiera.MossaAttuale - 1).Tempo * 8640000#)
    End If
    If TempoRimanente < 0 Then
      TempoRimanente = 0
    End If
    Elaborazione.MossaMigliore Scacchiera, Val(Tag), TempoRimanente, (TempoDiRiflessione), MossePerTempo, TempoBonusPerMossa, AP.Path + "\Archivio\Archivio.Sit", Motore
    If Scacchiera.Visible Then
      AspettaMossaElaborata (True)
    Else
      Unload Me
      Exit Sub
    End If
  End If
End Sub
Private Sub Scacchiera_Promozione(Colore As ControlloScacchiera.Colore, PezzoScelto As ControlloScacchiera.Genere)
  If PezzoScelto = Nessun_Genere Then
    AspettaSelezionePezzoPromosso = True
    Editabilit‡DurantePromozione = Scacchiera.Editabilit‡
    Scacchiera.Editabilit‡ = NonEditabile
    MsgBox Msg2, vbInformation
  End If
End Sub

Function FigureScacchi() As Object
   On Local Error Resume Next
   Set FigureScacchi = MDIScacchi.FigureScacchi
End Function
'Public Property Get AggiornamentiAbilitati() As TipoDiSelezione
'  AggiornamentiAbilitati = AggiornamentiAbilitati
'End Property

Private Sub Timer2_Timer()
  'Timer usato per rilevare quando il MOTORE risponde dando una mossa valida
  Timer2.Enabled = False
  Dim Mossa As Spostamento
  If Val(Tag) Then
    Dim Output As String
    Mossa = Elaborazione.MossaElaborata(Val(Tag), Output)
    If Output <> "" Then
      Commento = Commento + Output
      If Pensiero.Visible <> False Then
        If MDIScacchi.ActiveForm.hDC = Me.hDC Then
          Pensiero.Aggiorna Me
        End If
      End If
    End If
  End If
  If (Mossa.Provenienza.x <> Mossa.Destinazione.x) Or (Mossa.Provenienza.Y <> Mossa.Destinazione.Y) Then
    Output = Commento
    AspettaMossaElaborata False
    CalcoloMossaInCorso = False
    Scacchiera.EffettuaMossa Mossa, Commento
    Exit Sub
  End If
  Timer2.Enabled = True
End Sub
Sub AspettaMossaElaborata(Valore As Boolean)
  If Valore Then
    Commento = ""
  End If
  Timer2.Enabled = Valore
  AggiornaStatoElaborazione
End Sub
Sub AggiornaStatoElaborazione()
  If MDIScacchi.ActiveForm.hDC = Me.hDC Then
    MDIScacchi.Shape1.BackStyle = Abs(Timer2.Enabled)
    MDIScacchi.mnuMuovi.Enabled = Not Timer2.Enabled
  End If
End Sub
