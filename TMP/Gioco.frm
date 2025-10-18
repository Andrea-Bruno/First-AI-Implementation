VERSION 5.00
Object = "{706DFD08-82E6-11D1-AD79-F03E07C10000}#40.0#0"; "ControlloScacchiera.ocx"
Begin VB.Form Gioco 
   ClientHeight    =   1950
   ClientLeft      =   -3015
   ClientTop       =   -2190
   ClientWidth     =   2400
   ClipControls    =   0   'False
   Icon            =   "Gioco.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1950
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ControlloScacchiera.Scacchiera Scacchiera 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2990
      Editabilità     =   2
      ProvenienzaPezzo0=   0
      ProvenienzaPosizioneX0=   0
      ProvenienzaPosizioneY0=   0
      DestinazionePezzo0=   0
      DestinazionePosizioneX0=   0
      DestinazionePosizioneY0=   0
      BiancoArroccoPossibile0=   0
      NeroArroccoPossibile0=   0
      XEnPassant0     =   0
      Editabilità0    =   0
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
      BiancoNazionalità=   ""
      NeroNazionalità =   ""
      MossaAttuale    =   0
      QuantitàPezzi   =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
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
Public AggiornamentiAbilitati As Boolean
Private Sub Form_Initialize()
  Scacchiera.EventoData = DateValue(Date)
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
  If AggiornamentiAbilitati = True Then
    AP.AggiornaFinestre
  End If
End Sub

'Private Declare Function TutteLeMossePossibili Lib "Elaborazione" Alias "#1" (Colore As Colore, DisposizioneScacchiera As Variant, XEnPassant As Colonna, Mosse As Variant, QuantitàPezzi As Integer) As Integer
Private Sub Form_Load()
  Left = -3075
  Top = -2535
  Width = 3074
  Height = 2534
End Sub

Private Sub Form_Resize()
  Timer1.Enabled = False
  Timer1.Enabled = True
End Sub
Public Sub Aggiorna()
  Scacchiera.Move 0, 0, ScaleWidth, ScaleHeight
  Scacchiera.Visible = True
End Sub
Public Sub Colloca()
  Dim Max As Long
  Max = MDIScacchi.Width \ 2
  If MDIScacchi.Height \ 2 > Max Then
    Max = MDIScacchi.Height \ 2
  End If
  On Local Error Resume Next
  Move MDIScacchi.Width \ 4, 0, Max, Max
  'Aggiorna
End Sub
Private Sub Form_Terminate()
  AP.AggiornaFinestre
End Sub

Private Sub Scacchiera_CambiatoMossaAttuale(MossaAttuale As Integer)
  If AggiornamentiAbilitati = True Then
    AP.AggiornaElencoMosse
    AP.AggiornaEditabilità
  End If
End Sub

Private Sub Scacchiera_PezzoPosizionato(pezzo As ControlloScacchiera.pezzo, Provenienza As ControlloScacchiera.IndexCasa, Destinazione As ControlloScacchiera.IndexCasa)
  If AggiornamentiAbilitati = True Then
    AP.AggiornaEditabilità
    AP.AggiornaAnalisi
  End If
  If ((Scacchiera.Editabilità = BiancoMossePossibili) And (Computer = Scacchiera.BiancoTipoGiocatore)) Or ((Scacchiera.Editabilità = NeroMossePossibili) And (Computer = Scacchiera.NeroTipoGiocatore)) Then
    Dim Mossa As Spostamento
    Dim Turno As Colore
    If Scacchiera.Editabilità = BiancoMossePossibili Then
      Turno = Bianco
    Else
      Turno = Nero
    End If
    Mossa = Elaborazione.MossaMigliore(Turno, Scacchiera.DisposizioneScacchiera, Scacchiera.XEnPassant, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, Scacchiera.PosizioneReBianco, Scacchiera.PosizioneReNero, AP.Path + "\Archivio\Archivio.Sit")
    Scacchiera.EffettuaMossa Mossa
    End If
End Sub
Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Aggiorna
End Sub
Function FigureScacchi() As Object
   On Local Error Resume Next
   Set FigureScacchi = MDIScacchi.FigureScacchi
End Function
'Public Property Get AggiornamentiAbilitati() As TipoDiSelezione
'  AggiornamentiAbilitati = AggiornamentiAbilitati
'End Property

