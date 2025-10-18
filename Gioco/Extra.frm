VERSION 5.00
Begin VB.Form Extra 
   Caption         =   "Extra"
   ClientHeight    =   585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Crea archivio delle situazioni di gioco"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unifica tutto in archivio unico"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Extra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Const Msg1 = " non contiene LF (verr‡ sistemato!)"
Const Msg2 = "Creazione archivio delle situazioni di partita"
Const Msg3 = "Unione"
Const Msg4 = "Creazione di un indice univoco"
Dim AP As Variant
Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
End Sub
Private Sub Form_Terminate()
  Set AP = Nothing
End Sub
Private Sub Command1_Click()
  'VERIFICA SE I FILE SONO REGOLARI
  Dim File As String
  Dim FF As Integer
  Dim s As String
  Dim FilePg() As String
  Dim Qpg As Integer
  FF = FreeFile
  File = UCase(Dir(AP.Path + "\Archivio\*.pgn"))
  If File <> "" Then
    Do
      File = AP.Path + "\Archivio\" + File
      If InStr(UCase(File), "ARCHIVIO.PGN") = 0 Then
        Qpg = Qpg + 1
        ReDim Preserve FilePg(1 To Qpg)
        FilePg(Qpg) = File
        Open File For Binary As FF
          s = Space(100)
          Get FF, , s
          If InStr(s, Chr(13)) = 0 Then
            MsgBox (File + Msg1)
            s = Space(LOF(FF))
            Get FF, 1, s
            Sostituisci s, Chr(10), Chr(13) + Chr(10)
            Put FF, 1, s
          End If
        Close FF
      End If
      File = Dir
    Loop Until File = ""
  End If
  File = AP.Path + "\archivio\archivio.pgn"
  Dim QPartite As Long
  Dim N As Long
  Dim Codice As String
  Dim Prs As Form
  Dim FRMSCA As Object
  Dim Sca As Object
  Dim CoSt As Variant
  Set CoSt = New CodiceSituazione
  Dim SUnivoco As String
  Dim FileUnivoco As String
  FileUnivoco = AP.Path + "\Archivio\Univoco.tmp"
  If Dir(File) <> "" Then
    If Dir(FileUnivoco) <> "" Then
      FF = FreeFile
      Open FileUnivoco For Binary As FF
        SUnivoco = Space(LOF(FF))
        Get FF, , SUnivoco
      Close FF
    Else
      AP.Apri File
      QPartite = Archivio.ListCount
      If QPartite <> 0 Then
        Set Prs = New Progressione
        Prs.Show 0, Me
        Prs.Caption = Msg4
        Prs.Min = 0
        Prs.Max = QPartite - 1
        Set FRMSCA = MDIScacchi.ActiveForm
        Set Sca = FRMSCA.Scacchiera
        For N = 0 To QPartite - 1
          AP.CaricaPartita Archivio.Caption, Archivio.SeekPartita(N), MDIScacchi.ActiveForm
          Codice = Left(LTrim(Str(Year(Sca.EventoData))) + "    ", 4) + CoSt.Situazione(Sca.DisposizioneScacchiera, Sca.Editabilit‡, Sca.BiancoArroccoPossibile, Sca.NeroArroccoPossibile, Sca.XEnPassant)
          'Aggiungi codice univovo
          'PartiteArchiviate = PartiteArchiviate + 1
          'ReDim Preserve Univoco(1 To PartiteArchiviate)
          'Univoco(PartiteArchiviate) = Codice
          Dim FU As Integer
          FU = FreeFile
          Open FileUnivoco For Binary As FU
            Put FU, LOF(FU) + 1, Codice
          Close FU
          SUnivoco = SUnivoco + Codice
          Prs.Value = N
        Next
        Unload Prs
      End If
      Unload Archivio
    End If
  End If
  On Local Error Resume Next
    Dim StgPGN As String
    StgPGN = AP.Path + "\Archivio\PGN"
    MkDir StgPGN
    Dim StgERR As String
    StgERR = AP.Path + "\Archivio\ERROR"
    MkDir StgERR
  On Local Error GoTo 0
  Dim NF As Integer
  AP.Nuovo False, FRMSCA
  FRMSCA.AggiornamentiAbilitati = False
  Set Sca = FRMSCA.Scacchiera
  Sca.AbilitaVisualizzazionePezzi = False
  For NF = 1 To Qpg
    File = FilePg(NF)
    Archivio.LeggiFileArchivio File, True
    Dim NA As Integer
    Archivio.Visible = False
    QPartite = Archivio.ListCount
    If QPartite <> 0 Then
      Dim N2 As Long
      Set Prs = New Progressione
      Prs.Show 0, Me
      Prs.Caption = Msg3 + " " + File
      Prs.Min = 0
      Prs.Max = QPartite - 1
      For N = 0 To QPartite - 1
        Dim DisabilitaMSG As Boolean
        DisabilitaMSG = True
        Elaborazione.CaricaPartita Archivio.Caption, Archivio.SeekPartita(N), MDIScacchi.ActiveForm, DisabilitaMSG
        If DisabilitaMSG = True Then
          Codice = Left(LTrim(Str(Year(Sca.EventoData))) + "    ", 4) + CoSt.Situazione(Sca.DisposizioneScacchiera, Sca.Editabilit‡, Sca.BiancoArroccoPossibile, Sca.NeroArroccoPossibile, Sca.XEnPassant)
          'Verifica se la partita presa in considerazione Ë gi‡ stata archiviata
          Dim Gi‡Archiviata As Boolean
          If InStr(SUnivoco, Codice) <> 0 Then
            Gi‡Archiviata = True
          Else
            Gi‡Archiviata = False
          End If
          If Gi‡Archiviata = False Then
            'Aggiungi la partita all'archivio
            Elaborazione.CopiaPartita Archivio.Caption, Archivio.SeekPartita(N), AP.Path + "\archivio\archivio.pgn"
            FU = FreeFile
            Open FileUnivoco For Binary As FU
              Put FU, LOF(FU) + 1, Codice
            Close FU
            SUnivoco = SUnivoco + Codice
          Else
            Dim NR As Integer
            'NR = FreeFile
            'DoEvents
            'Open AP.Path + "\Archivio\Rapporto.Txt" For Append As NR
            'Print #NR, "Doppio: " + CStr(Sca.EventoData) + " " + Sca.BiancoGiocatore + " - " + Sca.NeroGiocatore
            'Close NR
          End If
          Prs.Value = N
        Else
          Exit For
        End If
      Next
      Unload Prs
      Unload Archivio
    End If
    If DisabilitaMSG = True Then
      MoveFile File, StgPGN
    Else
      NR = FreeFile
      DoEvents
      Open AP.Path + "\Archivio\Rapporto.Txt" For Append As NR
      Print #NR, "Errore in file: " + File + " Record:" + Str(N)
      Close NR
      MoveFile File, StgERR
    End If
  Next
  AP.Chiudi
  Set CoSt = Nothing
  Unload Me
End Sub
Private Sub MoveFile(File As String, ByVal Destinazione As String)
    Dim NF As Integer
    NF = FreeFile
    Dim P As Integer
    Dim OP As Integer
    If Right(Destinazione, 1) <> "\" Then
      Destinazione = Destinazione + "\"
    End If
    Do
      P = InStr(OP + 1, File, "\")
      If P <> 0 Then
        OP = P
      Else
        Exit Do
      End If
    Loop
    Dim nome As String
    nome = Mid(File, OP + 1)
    Destinazione = Destinazione + nome
    Open File For Binary As NF
      Dim NF2 As Integer
      NF2 = FreeFile
      Dim s As String
      s = Space(LOF(NF))
      Get NF, , s
      If Dir(Destinazione) <> "" Then
        Kill Destinazione
      End If
      Open Destinazione For Binary As NF2
        Put NF2, , s
      Close NF2
    Close NF
    Kill File
    DoEvents
End Sub
Private Sub Sostituisci(Stringa As String, carattere As String, sostituto As String)
  Dim NS As String
  Dim N As Long
  If Len(Stringa) <> 0 Then
    For N = 1 To Len(Stringa)
      Dim C As String
      C = Mid(Stringa, N, 1)
      If C = carattere Then
        C = sostituto
      End If
      NS = NS + C
    Next
  End If
  Stringa = NS
End Sub


Private Sub Command2_Click()
  Archivio.LeggiFileArchivio AP.Path + "\archivio\archivio.pgn", True
  Dim QPartite As Long
  Archivio.Visible = False
  QPartite = Archivio.ListCount
  If QPartite <> 0 Then
    Dim N As Long
    Dim N2 As Integer
    Dim FattoN As Long
    Dim FattoN2 As Integer
    AP.Nuovo
    Dim ArcS As Object
    Set ArcS = New ArchivioSituazioni
    ArcS.ApriFileSituazioni AP.Path + "\Archivio\Archivio.Sit", 2
    Dim FRMSCA As Object
    Set FRMSCA = MDIScacchi.ActiveForm
    Dim Sca As Object
    Set Sca = FRMSCA.Scacchiera
    FattoN = 0: FattoN2 = 1
    On Local Error GoTo fine
    Dim FilePosizione As String
    FilePosizione = AP.Path + "\Archivio\Situazione.Pos"
    Dim FF As Integer
    FF = FreeFile
    If Dir(FilePosizione) <> "" Then
      Open FilePosizione For Binary As FF
        Get #FF, , FattoN
        Get #FF, , FattoN2
      Close #FF
      Dim FLRiprendi As Boolean
      FLRiprendi = True
    End If
    Dim Prs As New Progressione
    Prs.Show 0, Me
    Prs.CMBAnnulla.Visible = True
    Prs.Caption = Msg2
    Prs.Max = QPartite - 1
    Prs.Min = FattoN

    For N = FattoN To QPartite - 1
      AP.CaricaPartita Archivio.Caption, Archivio.SeekPartita(N), MDIScacchi.ActiveForm
      FRMSCA.AggiornamentiAbilitati = False
      Sca.AbilitaVisualizzazionePezzi = False
      If FLRiprendi = True Then
        Sca.MossaAttuale = FattoN2 - 1
        FLRiprendi = False
      End If
      For N2 = Sca.MossaAttuale To 1 Step -1
        DoEvents
        Sca.MossaAttuale = N2
        Call ArcS.AggiungiSituazione(Sca.DisposizioneScacchiera, Sca.Editabilit‡, Sca.BiancoArroccoPossibile, Sca.NeroArroccoPossibile, Sca.XEnPassant, Sca.EventoRisultato, Archivio.SeekPartita(N))
        FattoN = N: FattoN2 = N2
      Next
      Prs.Value = N
      If Prs.Annulla = True Then
        Exit For
      End If
    Next
fine:
    FF = FreeFile
    Open FilePosizione For Binary As FF
      Put #FF, , FattoN
      Put #FF, , FattoN2
    Close #FF

    Unload Prs
    AP.Chiudi
    ArcS.ChiudiFileSituazioni
    ArcS.SalvaOrdine
    Set ArcS = Nothing
  End If
  Unload Archivio
  Unload Me
End Sub

