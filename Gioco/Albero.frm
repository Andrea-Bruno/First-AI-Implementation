VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Albero 
   Caption         =   "Albero della partita"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   Icon            =   "Albero.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":0626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":0942
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":0C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":0F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":1296
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":15B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":18CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":1F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":2222
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":253E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":2B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":2E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":31AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":34CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Albero.frx":37E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid ListaRiferimenti 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Doppio click sulla partita per caricarla"
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Albero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Option Base 1
Const FormattazioneGrid = "<Evento                       |<Luogo                                 |>           Data|>Round|<Bianco                      |<Nero                        |>Risultato|>     ECO"
Dim Riferimento As Variant
Dim AP As Variant

Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
End Sub

Private Sub Form_Load()
  ListaRiferimenti.FormatString = FormattazioneGrid
End Sub

Private Sub Form_Terminate()
  Set AP = Nothing
  Riferimento = Empty
End Sub
Public Sub CaricaAlbero(Partita As Variant, Quantit‡Pezzi As Integer, Optional ByVal NomeSituazione As String)
  Dim TopNode As node
  TreeView1.ImageList = ImageList1
  Dim ChiaveRoot As String
  ChiaveRoot = "R"
  If NomeSituazione <> "" Then
    NomeSituazione = " (" & NomeSituazione & ")"
  End If
  Set TopNode = TreeView1.Nodes.Add(, , ChiaveRoot, "Analisi di posizione" & NomeSituazione, 13, 14)
  AddPezzi TopNode, ChiaveRoot, Partita, Quantit‡Pezzi
End Sub
Private Sub AddPezzi(ByVal NodX As Variant, ByVal chiave As String, Partita As Variant, Quantit‡Pezzi As Integer)
  Dim NodX2 As node
  If Quantit‡Pezzi <> 0 Then
    'aggiungi i pezzi del giocatore a cui spotta muovere
    Dim pz As Integer
    For pz = 1 To Quantit‡Pezzi
      Dim etichetta As String
      etichetta = EtichettaPezzo(Partita(pz))
      Dim nome As String
      nome = LTrim$(Str$(Partita(pz).Posizione.x)) + LTrim$(Str$(Partita(pz).Posizione.Y))
      Dim ChiaveNodo As String
      ChiaveNodo = chiave & "\P" & nome
      'Crea il nodo solamente se non Ë esistente
      On Local Error Resume Next
      Set NodX2 = TreeView1.Nodes.Add(NodX, tvwChild, ChiaveNodo, etichetta, NumeroInmagine(Partita(pz)))
      'aggiungi le mosse possibili di ogni pezzo
      If Err.Number = 0 Then
        AddPossibiliMosse NodX2, ChiaveNodo, Partita(pz)
      Else
        Err.Number = 0
      End If
    Next
  End If
End Sub
Private Sub AddPossibiliMosse(ByVal NodX As node, ByVal chiave As String, pezzo As Variant)
  Dim NodX2 As node
  If pezzo.Quantit‡Mosse <> 0 Then
    Dim mp As Integer
    For mp = 1 To pezzo.Quantit‡Mosse
      Dim nome As String
      nome = LTrim$(Str$(pezzo.Mosse(mp).x)) + LTrim$(Str$(pezzo.Mosse(mp).Y))
      Dim etichetta As String
      etichetta = StringaPosizione(pezzo.Mosse(mp).x, pezzo.Mosse(mp).Y)
      Dim ChiaveNodo As String
      ChiaveNodo = chiave & "\A" & nome
      Set NodX2 = TreeView1.Nodes.Add(NodX, tvwChild, ChiaveNodo, etichetta, 15, 16)
    Next
  End If
End Sub
Private Function NumeroInmagine(Nodo As Variant) As Integer
  Select Case Nodo.pezzo
    Case Pedone_Bianco
      NumeroInmagine = 1
    Case Cavallo_Bianco
      NumeroInmagine = 2
    Case Alfiere_Bianco
      NumeroInmagine = 3
    Case Torre_Bianco
      NumeroInmagine = 4
    Case Dama_Bianco
      NumeroInmagine = 5
    Case Re_Bianco
      NumeroInmagine = 6
    Case Pedone_Nero
      NumeroInmagine = 7
    Case Cavallo_Nero
      NumeroInmagine = 8
    Case Alfiere_Nero
      NumeroInmagine = 9
    Case Torre_Nero
      NumeroInmagine = 10
    Case Dama_Nero
      NumeroInmagine = 11
    Case Re_Nero
      NumeroInmagine = 12
  End Select
End Function
Private Function EtichettaPezzo(Nodo As Variant) As String
Dim x As Colonna
Dim Y As Riga

  x = Nodo.Posizione.x
  Y = Nodo.Posizione.Y

  Select Case Nodo.pezzo
    Case Pedone_Bianco
      EtichettaPezzo = "Pedone bianco"
    Case Cavallo_Bianco
      EtichettaPezzo = "Cavallo bianco"
    Case Alfiere_Bianco
      EtichettaPezzo = "Alfiere bianco"
    Case Torre_Bianco
      EtichettaPezzo = "Torre bianca"
    Case Dama_Bianco
      EtichettaPezzo = "Dama bianca"
    Case Re_Bianco
      EtichettaPezzo = "Re bianco"
    Case Pedone_Nero
      EtichettaPezzo = "Pedone nero"
    Case Cavallo_Nero
      EtichettaPezzo = "Cavallo nero"
    Case Alfiere_Nero
      EtichettaPezzo = "Alfiere nero"
    Case Torre_Nero
      EtichettaPezzo = "Torre nera"
    Case Dama_Nero
      EtichettaPezzo = "Dama nera"
    Case Re_Nero
      EtichettaPezzo = "Re nero"
  End Select
  EtichettaPezzo = EtichettaPezzo + " " + StringaPosizione(x, Y)
End Function
Function StringaPosizione(x As Colonna, Y As Riga) As String
  StringaPosizione = "(" + Chr$(x + 65) + LTrim$(Str$(Y + 1)) + ")"
End Function

Private Sub Form_Resize()
  Dim OFS As Long
  If ScaleHeight < ListaRiferimenti.Height Then
    OFS = 0
    ListaRiferimenti.Visible = False
  Else
    OFS = ListaRiferimenti.Height
    ListaRiferimenti.Width = ScaleWidth
    ListaRiferimenti.Visible = True
  End If
  TreeView1.Move 0, 0 + OFS, ScaleWidth, ScaleHeight - OFS
  TreeView1.Visible = True
End Sub


Private Sub ListaRiferimenti_DblClick()
  Carica
End Sub
Public Sub Carica()
  If ListaRiferimenti.Row < ListaRiferimenti.Rows - 1 Then
    Dim Scacchiera As Object
    Dim ScacchieraConAnalogia As Object
    Set Scacchiera = MDIScacchi.ActiveForm.Scacchiera
    Dim Situazione As String
    Dim Posizione As Long
    'Posizione = ListaRiferimenti.ItemData(ListaRiferimenti.ListIndex)
    Posizione = Riferimento(ListaRiferimenti.Row)
    Situazione = Elaborazione.Situazione(Scacchiera.DisposizioneScacchiera, Scacchiera.Editabilit‡, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, Scacchiera.XEnPassant)
    AP.CaricaPartita AP.Path + "\archivio\archivio.pgn", Posizione, ScacchieraConAnalogia
    Set ScacchieraConAnalogia = ScacchieraConAnalogia.Scacchiera
    Dim N As Integer
    MDIScacchi.ActiveForm.AggiornamentiAbilitati = False
    ScacchieraConAnalogia.AbilitaVisualizzazionePezzi = False
    For N = ScacchieraConAnalogia.MossaAttuale To 0 Step -1
      ScacchieraConAnalogia.MossaAttuale = N
      If Situazione = Elaborazione.Situazione(ScacchieraConAnalogia.DisposizioneScacchiera, ScacchieraConAnalogia.Editabilit‡, ScacchieraConAnalogia.BiancoArroccoPossibile, ScacchieraConAnalogia.NeroArroccoPossibile, ScacchieraConAnalogia.XEnPassant) Then
        Exit For
      End If
    Next
    MDIScacchi.ActiveForm.AggiornamentiAbilitati = True
    ScacchieraConAnalogia.AbilitaVisualizzazionePezzi = True
  End If
End Sub

Private Sub ListaRiferimenti_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 32 Then
    Carica
  End If
End Sub

Private Sub TreeView1_NodeClick(ByVal node As MSComctlLib.node)
  Dim v As Integer
  If node.Children = 0 Then
    v = Aggiunginodo(node)
  End If
End Sub
Function Aggiunginodo(node As MSComctlLib.node) As Integer
  Dim percorso As String
  percorso = node.Key
  Dim DS As Variant
  DS = MDIScacchi.ActiveForm.Scacchiera.DisposizioneScacchiera
  
  Dim XEnPassant As Colonna
  XEnPassant = MDIScacchi.ActiveForm.Scacchiera.XEnPassant
  Dim BiancoArroccoPossibile As Lato
  BiancoArroccoPossibile = MDIScacchi.ActiveForm.Scacchiera.BiancoArroccoPossibile
  Dim NeroArroccoPossibile As Lato
  NeroArroccoPossibile = MDIScacchi.ActiveForm.Scacchiera.NeroArroccoPossibile
  Dim PosizioneReBianco As Variant
  PosizioneReBianco = MDIScacchi.ActiveForm.Scacchiera.PosizioneReBianco
  Dim PosizioneReNero As Variant
  PosizioneReNero = MDIScacchi.ActiveForm.Scacchiera.PosizioneReNero
  
  
  Dim N As Variant
  Dim Turno As Colore
  Turno = ProssimoTurno(MDIScacchi.ActiveForm.Scacchiera.Editabilit‡)
  N = 0
  Do
    N = InStr(N + 1, percorso, "P")
    If N = 0 Then
      Exit Do
    Else
      Dim PosizioneInizialeX As Colonna
      Dim PosizioneInizialeY As Riga
      Dim PosizioneFinaleX As Colonna
      Dim PosizioneFinaleY As Riga
      PosizioneInizialeX = Val(Mid$(percorso, N + 1, 1))
      PosizioneInizialeY = Val(Mid$(percorso, N + 2, 1))
      If Mid$(percorso, N + 4, 1) <> "A" Then
        Aggiunginodo = False
        Exit Function
      End If
      PosizioneFinaleX = Val(Mid$(percorso, N + 5, 1))
      PosizioneFinaleY = Val(Mid$(percorso, N + 6, 1))
      'esegui la mossa nel modo indicato dal nodo
      Call Elaborazione.EffettuaMossa(PosizioneInizialeX, PosizioneInizialeY, PosizioneFinaleX, PosizioneFinaleY, DS, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile)
      'Ds(PosizioneFinaleX, PosizioneFinaleY) = Ds(PosizioneInizialeX, PosizioneInizialeY)
      'Ds(PosizioneInizialeX, PosizioneInizialeY) = Nessun_Pezzo
      If Turno = Nero Then
        Turno = Bianco
      Else
        Turno = Nero
      End If
    End If
  Loop
  Dim Quantit‡Pezzi As Integer
  Dim q As Integer
  Dim Mosse() As Variant
  q = Elaborazione.TutteLeMossePossibili(Turno, DS, Mosse, Quantit‡Pezzi, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero)
  Aggiunginodo = True
  AddPezzi ByVal percorso, percorso, Mosse, Quantit‡Pezzi
End Function
Sub ResettaTreeView()
  Dim Nodo As Variant
  Dim Idx() As Variant
  Dim NElementi As Integer
  For Each Nodo In TreeView1.Nodes
    NElementi = NElementi + 1
    ReDim Preserve Idx(NElementi)
    Idx(NElementi) = Nodo.Index
  Next
  If NElementi <> 0 Then
    Dim N As Integer
    For N = NElementi To 1 Step -1
      TreeView1.Nodes.Remove (Idx(N))
    Next
  End If
End Sub
Sub ResettaRiferimenti()
  While ListaRiferimenti.Rows <> 2
    ListaRiferimenti.RemoveItem 1
  Wend
  Riferimento = Empty
End Sub
Sub AggiornaRiferimentiSituazione(Scacchiera As Object)
  Dim File As String
  File = AP.Path + "\archivio\archivio"
  ResettaRiferimenti
  Dim Turno As Colore
  Turno = ProssimoTurno(Scacchiera.Editabilit‡)
  If Turno = Bianco Or Turno = Nero Then
    Riferimento = Elaborazione.RiferimentoArchivio(Turno, Scacchiera.DisposizioneScacchiera, Scacchiera.XEnPassant, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, File + ".sit")
    If Not IsEmpty(Riferimento) Then
      Dim N As Integer
      Dim NumeroFile As Integer
      NumeroFile = FreeFile
      Open File + ".pgn" For Input As #NumeroFile
      For N = 1 To Riferimento(0)
        ListaRiferimenti.AddItem Elaborazione.Partita(File + ".pgn", CLng(Riferimento(N))), N
'        ListaRiferimenti.ItemData(N) = Riferimento(N)
      Next
      Close #NumeroFile
    End If
  End If
End Sub
Sub CreaAlberoDellaPosizione(Scacchiera As Object)
    Dim q As Integer
    Dim Quantit‡Pezzi As Integer
    Dim Mosse() As Variant
    Dim Turno As Colore
    ResettaTreeView
    Turno = ProssimoTurno(Scacchiera.Editabilit‡)
    If Turno = Bianco Or Turno = Nero Then
      q = Elaborazione.TutteLeMossePossibili(Turno, Scacchiera.DisposizioneScacchiera, Mosse, Quantit‡Pezzi, Scacchiera.XEnPassant, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, Scacchiera.PosizioneReBianco, Scacchiera.PosizioneReNero)
      CaricaAlbero Mosse, Quantit‡Pezzi, Scacchiera.Parent.Caption
'    Else
'      ResettaTreeView
    End If
    AggiornaRiferimentiSituazione Scacchiera
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

