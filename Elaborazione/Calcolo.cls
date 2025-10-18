VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calcolo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
DefInt A-Z
Option Explicit
Option Base 1
Const MinValLong = -2147483648#
Dim ArcS As New ArchivioSituazioni
Dim PG As New PGN
Type Univoco
  FileRicezione As String
  Posizione As Long
  Mossa As Spostamento
  Valutazione As Long
End Type
Dim Uni(1 To 32) As Univoco
Public Function CaricaIndiceArchivio(File As String)
      ArcS.ApriFileSituazioni File
End Function
Public Function RiferimentoArchivio(Turno As Colore, DisposizioneScacchiera As Variant, XEnPassant As Colonna, BiancoArroccoPossibile As Lato, NeroArroccoPossibile As Lato, Optional FileSituazioni As String) As Variant
    Dim Colore As EditorScacchiera
    If Turno <> Nessun_Colore Then
      If Turno = Bianco Then
        Colore = BiancoMossePossibili
      Else
        Colore = NeroMossePossibili
      End If
      ArcS.ApriFileSituazioni FileSituazioni
      Dim StS As String
      StS = ArcS.CodiceSituazione.Situazione(DisposizioneScacchiera, Colore, BiancoArroccoPossibile, NeroArroccoPossibile, XEnPassant)
      Dim Idx As Long
      If ArcS.EsisteSituazione(StS, Idx) Then
          StS = ArcS.CaricaDaFile(3)
          Dim Vittorie As Integer
          Dim Pareggi As Integer
          Dim Sconfitte As Integer
          Vittorie = Asc(Mid(StS, 1, 1))
          Pareggi = Asc(Mid(StS, 2, 1))
          Sconfitte = Asc(Mid(StS, 3, 1))
          Dim Partite As Integer
          Partite = Vittorie + Pareggi + Sconfitte
          Dim N As Integer
          ReDim SeekPartita(0 To Partite) As Long
          SeekPartita(0) = Partite
          StS = ArcS.CaricaDaFile(4 * Partite)
          For N = 1 To Partite
            SeekPartita(N) = Asc(Mid$(StS, (N - 1) * 4 + 1, 1)) + Asc(Mid$(StS, (N - 1) * 4 + 2, 1)) * CLng(256) + Asc(Mid$(StS, (N - 1) * 4 + 3, 1)) * CLng(65536) + Asc(Mid$(StS, (N - 1) * 4 + 4, 1)) * CLng(16777216)
          Next
          RiferimentoArchivio = SeekPartita()
      Else
        RiferimentoArchivio = Empty
      End If
      ArcS.ChiudiFileSituazioni
    End If
End Function
Public Function MossaElaborata(Univoco As Integer, Optional Output As String, Optional Valutazione As Long = MinValLong, Optional Commento As String) As Spostamento
  'Univoco deve essere un numero che identifica la scacchiera con cui si gioca, la routine in questo modo nominera
  'i file di collocquio con il motore dandogli una estenzione numerica univoca: esempio: absdefg.001 (in questo modo
  'evito i conflitti quando il programma gioca pi˘ partite contemporaneamente
  'Valutazione non Ë un parametro ma bensi restituisce il punteggio della MossaMigliore
  'Apri il file input per ricevere dal motore
  Dim P As Integer, P2 As Integer
  If (Uni(Univoco).Mossa.Provenienza.x <> Uni(Univoco).Mossa.Destinazione.x) Or (Uni(Univoco).Mossa.Provenienza.Y <> Uni(Univoco).Mossa.Destinazione.Y) Then
    'Allora la mossa Ë stata trovata in archivio e Ë anche stata ritenuta valida
    MossaElaborata = Uni(Univoco).Mossa
    Uni(Univoco).Mossa.Provenienza.x = NessunaColonna
    Uni(Univoco).Mossa.Provenienza.Y = NessunaColonna
    Uni(Univoco).Mossa.Destinazione.x = NessunaColonna
    Uni(Univoco).Mossa.Destinazione.Y = NessunaColonna
    Valutazione = Uni(Univoco).Valutazione
    Output = "POINT" + Str(Valutazione)
  Else
    Dim ri As Integer
    ri = FreeFile
    Open Uni(Univoco).FileRicezione For Input As #ri
    Seek ri, Uni(Univoco).Posizione
    Dim Ricezione As String
    Dim bf As String
    Dim p10 As Long
    Do While Not (EOF(ri)) Or bf <> ""
      If bf = "" Then
        Line Input #ri, Ricezione
      Else
        Ricezione = bf
      End If
      p10 = InStr(Ricezione, Chr(10))
      If p10 Then
        bf = Mid(Ricezione, p10 + 1)
        Ricezione = Left(Ricezione, p10 - 1)
      End If
      Uni(Univoco).Posizione = Seek(ri)
      If Ricezione <> "" Then
        'Debug.Print Ricezione
        Commento = Commento + Chr(13) + Chr(10) + Ricezione
        Ricezione = LTrim(Ricezione)
        Output = Output + Ricezione + Chr(13) + Chr(10)
        'Se riceve la mossa
        Dim MossaIntercettata As Boolean
        MossaIntercettata = False
        If Val(Ricezione) <> 0 And (InStr(Ricezione, ".") < InStr(Ricezione, " ")) And InStr(Ricezione, ".") <> 0 Then
          Ricezione = LTrim(RTrim(Ricezione))
          P = InStr(Ricezione, " ")
          Do While P <> 0
            P2 = P
            P = InStr(P + 1, Ricezione, " ")
          Loop
          Ricezione = Mid(Ricezione, P2 + 1)
          MossaIntercettata = True
        ElseIf Left(Ricezione, 5) = "move " Then
          Ricezione = Mid(Ricezione, 6)
          MossaIntercettata = True
        End If
        If MossaIntercettata Then
          MossaElaborata.Provenienza.x = Asc(Mid(Ricezione, 1, 1)) - 97
          MossaElaborata.Provenienza.Y = Val(Mid(Ricezione, 2, 1)) - 1
          MossaElaborata.Destinazione.x = Asc(Mid(Ricezione, 3, 1)) - 97
          MossaElaborata.Destinazione.Y = Val(Mid(Ricezione, 4, 1)) - 1
          If Len(Ricezione) > 4 Then
            'Un pedone Ë andato a promozione
            Select Case LCase(Mid(Ricezione, 5, 1))
              Case "q"
                MossaElaborata.Promosso = Dama
              Case "r"
                MossaElaborata.Promosso = Torre
              Case "b"
                MossaElaborata.Promosso = Alfiere
              Case "n"
                MossaElaborata.Promosso = Cavallo
            End Select
          End If
          'Exit Do
        End If
      End If
    Loop
    Close #ri
  End If
End Function
Public Sub MossaMigliore(Scacchiera As Object, Univoco As Integer, Optional TempoRimanente As Long, Optional TempoDiRiflessione As Integer, Optional MossePerTempo As Integer, Optional TempoBonusPerMossa As Integer, Optional FileSituazioni As String, Optional Motore As String)
  'Univoco deve essere un numero che identifica la scacchiera con cui si gioca, la routine in questo modo nominera
  'i file di collocquio con il motore dandogli una estenzione numerica univoca: esempio: absdefg.001 (in questo modo
  'evito i conflitti quando il programma gioca pi˘ partite contemporaneamente
  'TempoRimanente Ë in Secondi/100
  'TempoDiRiflessione Ë in secondi
  'TempoBonusPerMossa Ë in secondi
  Dim FlgValutazioneAccettabile As Boolean
  Dim Quantit‡Mosse As Integer
  Dim Mosse() As Variant
  Dim Quantit‡Pezzi As Integer
  Quantit‡Mosse = TutteLeMossePossibili(Scacchiera.Editabilit‡, Scacchiera.DisposizioneScacchiera, Mosse, Quantit‡Pezzi, Scacchiera.XEnPassant, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, Scacchiera.PosizioneReBianco, Scacchiera.PosizioneReNero)
  Dim N As Integer
  Dim M As Integer
  Dim Punteggio As Long
  If Quantit‡Mosse <> 0 Then
    'Valutazioine tramite le situazioni analoghe in archivio
    'If True = False Then
      ArcS.ApriFileSituazioni FileSituazioni
      Uni(Univoco).Valutazione = MinValLong
      For N = 1 To Quantit‡Pezzi
        If Mosse(N).Quantit‡Mosse <> 0 Then
          For M = 1 To Mosse(N).Quantit‡Mosse
            Dim ValutazioneEffettuata As Boolean
            Punteggio = ValutazioneInArchivio(Mosse(N).Posizione.x, Mosse(N).Posizione.Y, Mosse(N).Mosse(M).x, Mosse(N).Mosse(M).Y, Scacchiera.Editabilit‡, Scacchiera.DisposizioneScacchiera, Scacchiera.XEnPassant, Scacchiera.BiancoArroccoPossibile, Scacchiera.NeroArroccoPossibile, Scacchiera.PosizioneReBianco, Scacchiera.PosizioneReNero, ValutazioneEffettuata)
            If ValutazioneEffettuata = True Then
              If (Punteggio > Uni(Univoco).Valutazione) And (Punteggio > 0) Then
                FlgValutazioneAccettabile = True
                Uni(Univoco).Valutazione = Punteggio
                Uni(Univoco).Mossa.Provenienza.x = Mosse(N).Posizione.x
                Uni(Univoco).Mossa.Provenienza.Y = Mosse(N).Posizione.Y
                Uni(Univoco).Mossa.Destinazione.x = Mosse(N).Mosse(M).x
                Uni(Univoco).Mossa.Destinazione.Y = Mosse(N).Mosse(M).Y
              End If
            End If
          Next
        End If
      Next
      ArcS.ChiudiFileSituazioni
    'End If
    
    If FlgValutazioneAccettabile = False Then
      Dim Trasmette As String, Riceve As String, Posizione As String, BAT As String, estenzione As String
      estenzione = Right("000" + LTrim(Str(Univoco)), 3)
      Posizione = estenzione + ".fen" 'questo file deve perforza avere l'estenzione ".FEN" cosi il codice univoco lo metto avanti
      Trasmette = "hismove." + estenzione
      Riceve = "mymove." + estenzione
      BAT = "execute.BAT"
      Dim Tr As Integer, Fe As Integer
      Fe = FreeFile
      'Trova la path del Motore di calcolo
      Dim P As Integer, P2 As Integer
      P = InStr(Motore, "\")
      Do While P <> 0
        P2 = P
        P = InStr(P + 1, Motore, "\")
      Loop
      Dim PathMotore As String
      If P2 <> 0 Then
        PathMotore = Left(Motore, P2)
      End If
      'Trova il nome del motore
      Dim NomeMotore As String
      If P2 <> 0 Then
        NomeMotore = UCase(Mid(Motore, P2 + 1))
      Else
        NomeMotore = UCase(Motore)
      End If
      
      'Calcola il FEN
      Dim FEN As String
      FEN = PG.FEN(Scacchiera)
      'il motore cilian piÚ caricare il FEN solo da file
      If InStr(NomeMotore, "CILIAN") Then
        Open PathMotore + Posizione For Output As #Fe
        Print #Fe, FEN
        Close Fe
      End If
      'Crea il file in cui Ë contenuta la digitazione da tastiera che il MOTORE ricever‡ come input all'avvio
      Tr = FreeFile
      Open PathMotore + Trasmette For Output As #Tr
      'ogni motore ha implementato le funzioni in modo diverso !
      If InStr(NomeMotore, "CILIAN") Then
        'Cilinn
        Print #Tr, "wboard"
        Print #Tr, "load " + Posizione
      ElseIf InStr(NomeMotore, "FORTRESS") Then
        'Fortres
        Print #Tr, "fen " + FEN
      ElseIf InStr(NomeMotore, "EXCHESS") Then
        'EXChess
        Print #Tr, "setboard " + FEN
      End If
      If TempoRimanente Then
        Print #Tr, "time" + Str(TempoRimanente)
      End If
      Print #Tr, "post"
      Print #Tr, "level" + Str(MossePerTempo) + Str(TempoDiRiflessione) + Str(TempoBonusPerMossa)
      Print #Tr, "go"
      Print #Tr, "quit"
      Close Tr
      'Lanci in esecuzione il motore
      'Il motore deve essere lanciato indirettamente tramite un file .BAT cosi potra ricevere dei parametri
      'Creo il file .BAT di avvio
      Open PathMotore + BAT For Output As #Tr
        Print #Tr, "@" + NomeMotore + " -nobk <"; Trasmette + ">"; Riceve
        Close #Tr
        'Elimina il vecchio file di Ricezione se esistente
        If Dir(PathMotore + Riceve) <> "" Then
        
        On Local Error GoTo Ritenta
        GoTo ByPass
Ritenta:
        DoEvents
        Resume
ByPass:
        Kill PathMotore + Riceve
        On Local Error GoTo 0
      End If
      'Manda in esecuzione il motore
      Dim OldPath As String
      OldPath = CurDir
      ChDir PathMotore
      Shell PathMotore + BAT, 0
      ChDir OldPath
      'Attendi che il file di input venghi generato
      Do
        DoEvents
      Loop Until Dir(PathMotore + Riceve) <> ""
      Uni(Univoco).FileRicezione = PathMotore + Riceve
      Uni(Univoco).Posizione = 1
      'Dim T As Double
      'T = Timer + 1
      'Do
      '  DoEvents
      'Loop Until Timer > T
      'Kill PathMotore + Riceve
      '===========================================================================================================
      'La parte seguente viene disattivata fino a quando sar‡ pronta una versione con un motore di calcolo proprio
      '===========================================================================================================
      'For N = 1 To Quantit‡Pezzi
      '  If Mosse(N).Quantit‡Mosse <> 0 Then
      '    For M = 1 To Mosse(N).Quantit‡Mosse
      '      Punteggio = ValutaMossa(Mosse(N).Posizione.x, Mosse(N).Posizione.Y, Mosse(N).Mosse(M).x, Mosse(N).Mosse(M).Y, Turno, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero, 0, 4, 0)
      '      If Punteggio > Valutazione Then
      '        Valutazione = Punteggio
      '        MossaMigliore.Provenienza.x = Mosse(N).Posizione.x
      '        MossaMigliore.Provenienza.Y = Mosse(N).Posizione.Y
      '        MossaMigliore.Destinazione.x = Mosse(N).Mosse(M).x
      '        MossaMigliore.Destinazione.Y = Mosse(N).Mosse(M).Y
      '      End If
      '    Next
      '  End If
      'Next
    End If
  End If
End Sub
Public Function ValutazioneInArchivio(X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, Turno As Colore, ByVal DisposizioneScacchiera As Variant, ByVal XEnPassant As Colonna, ByVal BiancoArroccoPossibile As Lato, ByVal NeroArroccoPossibile As Lato, ByVal PosizioneReBianco As Variant, ByVal PosizioneReNero As Variant, ValutazioneEffettuata As Boolean) As Long
  'Questa funzione restituisce una valutazione statistica in base all'archivio in uso
  'In particolare ricerca se sono state giocate pertite in cui si Ë verificata la stessa situazione e in base
  'al risultato di queste viene generata la valutazione
  'Il parametro ValutazioneEffettuata resituisce se Ë stato possibile un tipo di valutazione tramite l'archivio
  'If X1 = Colonna_b Then
  '  If Y1 = Riga_7 Then
  '    If X2 = Colonna_b Then
  '      If Y2 = Riga_4 Then
  '        Stop
  '      End If
  '    End If
  '  End If
  'End If
  EffettuaMossa X1, Y1, X2, Y2, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile
  Dim Colore As EditorScacchiera
  If Turno = Bianco Then
    Colore = NeroMossePossibili
  Else
    Colore = BiancoMossePossibili
  End If
'  Colore = Turno
  Dim StS As String
  StS = ArcS.CodiceSituazione.Situazione(DisposizioneScacchiera, Colore, BiancoArroccoPossibile, NeroArroccoPossibile, XEnPassant)
  Dim Idx As Long
  If ArcS.EsisteSituazione(StS, Idx) Then
    StS = ArcS.CaricaDaFile(3)
    Dim Vittorie As Long
    Dim Pareggi As Long
    Dim Sconfitte As Long
    Sconfitte = Asc(Mid(StS, 1, 1))
    Pareggi = Asc(Mid(StS, 2, 1))
    Vittorie = Asc(Mid(StS, 3, 1))
    ValutazioneInArchivio = (Vittorie * 2) + Pareggi - (Sconfitte * 2)
    ValutazioneEffettuata = True
  End If
End Function
Public Function ValutaMossa(X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, Turno As Colore, ByVal DisposizioneScacchiera As Variant, ByVal XEnPassant As Colonna, ByVal BiancoArroccoPossibile As Lato, ByVal NeroArroccoPossibile As Lato, ByVal PosizioneReBianco As Variant, ByVal PosizioneReNero As Variant, ByVal Recursione As Integer, MaxRecursioni As Integer, Optional ByVal PunteggioProvvisorio As Long, Optional AnalizzaTutto As Boolean) As Long
  Recursione = Recursione + 1
  Dim XR1 As Colonna
  Dim YR1 As Riga
  Dim XR2 As Colonna
  Dim YR2 As Riga
  Dim FiltroRamo As Boolean
  Dim FLG As Boolean
  Dim Quantit‡Mosse As Integer
  Dim Mosse() As Variant
  Dim Quantit‡Pezzi As Integer
  Dim Qualit‡ As Long
  Dim Punti As Long
  Dim PunteggioMinimo As Long
  Dim ProssimoTurno As Colore
  Dim PezzoCatturato As Genere
  '=====================================
  'vi Ë un errore: PezzoCatturato assume il valore 0 con la presa EnPassant (Devo Sistemare)
  PezzoCatturato = Abs(DisposizioneScacchiera(X2, Y2))
  If PezzoCatturato <> Nessun_Genere Then
    Select Case PezzoCatturato
      Case Pedone
        Qualit‡ = ValorePedone
      Case Cavallo
        Qualit‡ = ValoreCavallo
      Case Alfiere
        Qualit‡ = ValoreAlfiere
      Case Torre
        Qualit‡ = ValoreTorre
      Case Dama
        Qualit‡ = ValoreDama
      Case RE
        Qualit‡ = ValoreRe
    End Select
  Else
    Qualit‡ = 0
  End If
'  If X2 = 1 And Y2 = 5 And Recursione = 1 Then
'    Stop
'  End If
  If Recursione < MaxRecursioni Then
    'If TutteLePossibilit‡ = False Then
    '  Dim NM As Boolean
    '  NM = NuovaMinaccia(X1, Y1, X2, Y2, DisposizioneScacchiera)
    'End If
    Dim Continua As Boolean
    Dim NuovoPunteggioProvvisorio
    NuovoPunteggioProvvisorio = Qualit‡ + PunteggioProvvisorio
    If (Recursione Mod 2) = 0 Then
      If PunteggioProvvisorio >= 0 Then
        Continua = True
      End If
    Else
      Continua = True
    End If
    If Continua Then
      Dim AT As Boolean
      If AnalizzaTutto = False Then
        AT = NuovaMinaccia(X1, Y1, X2, Y2, DisposizioneScacchiera)
      End If
      If AT Or (Qualit‡ <> 0) Or AnalizzaTutto Then
        EffettuaMossa X1, Y1, X2, Y2, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile
        If Abs(DisposizioneScacchiera(X2, Y2)) = RE Then
          If DisposizioneScacchiera(X2, Y2) = Re_Bianco Then
            PosizioneReBianco.x = X2
            PosizioneReBianco.Y = Y2
          Else
            PosizioneReNero.x = X2
            PosizioneReNero.Y = Y2
          End If
        End If
        Dim XRe As Colonna
        Dim YRe As Riga
        If Turno = Bianco Then
          XRe = PosizioneReNero.x
          YRe = PosizioneReNero.Y
          ProssimoTurno = Nero
        Else
          XRe = PosizioneReBianco.x
          YRe = PosizioneReBianco.Y
          ProssimoTurno = Bianco
        End If
        Quantit‡Mosse = TutteLeMossePossibili(ProssimoTurno, DisposizioneScacchiera, Mosse, Quantit‡Pezzi, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero)
        If Quantit‡Mosse = 0 Then
          Qualit‡ = ValoreRe
        End If
        Dim N As Integer
        Dim M As Integer
        Dim Punteggio As Long
        If Quantit‡Mosse <> 0 Then
          For N = 1 To Quantit‡Pezzi
            If Mosse(N).Quantit‡Mosse <> 0 Then
              For M = 1 To Mosse(N).Quantit‡Mosse
                XR1 = Mosse(N).Posizione.x: YR1 = Mosse(N).Posizione.Y: XR2 = Mosse(N).Mosse(M).x: YR2 = Mosse(N).Mosse(M).Y
                Punti = Qualit‡ - ValutaMossa(XR1, YR1, XR2, YR2, ProssimoTurno, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero, Recursione, MaxRecursioni, NuovoPunteggioProvvisorio, AT)
                If Punti < PunteggioMinimo Or FLG = False Then
                  FLG = True
                  PunteggioMinimo = Punti
                End If
              Next
            End If
          Next
        End If
      End If
    Else
      'Siamo arrivati al massimo livello di recursione, bisogna effettuare una valutazione
    End If
  End If
  If FLG = True Then
    ValutaMossa = PunteggioMinimo
  Else
    ValutaMossa = Qualit‡
  End If
End Function
Public Function NuovaMinaccia(X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, DisposizioneScacchiera As Variant) As Boolean
  'Questa funzione restituisce True quando lo spostamento di una pedina minaccia una pedina avversaria
  'che prima non era minacciata da questa
  'Questa funzione non fa nessun controllo se vi Ë un pedone che Ë promosso (Pedone Nero su riga 1 o Pedone Bianco su riga 8)
  Dim X3 As Colonna, Y3 As Riga
  Dim pezzo As pezzo
  pezzo = DisposizioneScacchiera(X1, Y1)
  If pezzo > 0 Then
  'Il pezzo Ë Bianco
    If pezzo = Pedone Then
      If Y2 <= Riga_8 Then
        If X2 > Colonna_A Then
          If DisposizioneScacchiera(X2 - 1, Y2 + 1) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        If X2 < Colonna_H Then
          If DisposizioneScacchiera(X2 + 1, Y2 + 1) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    ElseIf pezzo = Cavallo Then
      Y3 = Y2 + 2
      If Y3 <= Riga_8 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2 - 2
      If Y3 >= Riga_1 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 + 2
      If X3 <= Colonna_H Then
        Y3 = Y2 - 1
        If Y3 >= Riga_1 Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        Y3 = Y2 + 1
        If Y3 <= Riga_8 Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 - 2
      If X3 >= Colonna_A Then
        Y3 = Y2 - 1
        If Y3 >= Riga_1 Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        Y3 = Y2 + 1
        If Y3 <= Riga_8 Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    ElseIf pezzo = RE Then
      Y3 = Y2 + 1
      If Y3 <= Riga_8 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2
        If DisposizioneScacchiera(X3, Y3) < 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2 - 1
      If Y3 >= Riga_1 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) < 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2
        If DisposizioneScacchiera(X3, Y3) < 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2
      X3 = X2 - 1
      If X3 >= Colonna_A Then
        If DisposizioneScacchiera(X3, Y3) < 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 + 1
      If X3 <= Colonna_H Then
        If DisposizioneScacchiera(X3, Y3) < 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    Else
      If pezzo And Alfiere Then
        If ((X3 - X2) - (Y3 - Y2)) = 0 Then
          'Controlla la diagonale destra
          'In alto
          Y3 = Y2: X3 = X2
          Do While Y3 < Riga_8 And X3 < Colonna_H
            Y3 = Y3 + 1: X3 = X3 + 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) < 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
          'Controlla la diagonale sinistra
          'In basso
          X3 = X2: Y3 = Y2
          Do While Y3 > Riga_1 And X3 > Colonna_A
            Y3 = Y3 - 1:    X3 = X3 - 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) < 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
        Else
          'Controlla la diagonale destra
          'In basso
          Y3 = Y2: X3 = X2
          Do While Y3 > Riga_1 And X3 < Colonna_H
            Y3 = Y3 - 1: X3 = X3 + 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) < 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
          'Controlla la diagonale sinistra
          'In alto
          Y3 = Y2: X3 = X2
          Do While Y3 < Riga_8 And X3 > Colonna_A
            Y3 = Y3 + 1:    X3 = X3 - 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) < 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
        End If
      End If
      If pezzo And Torre Then
        If Y1 = Y2 Then
          'controlla la verticale sopra
          For Y3 = Y2 + 1 To Riga_8
            If DisposizioneScacchiera(X2, Y3) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X2, Y3) < 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
          'Controlla la verticale sotto
          For Y3 = Y2 - 1 To Riga_1 Step -1
            If DisposizioneScacchiera(X2, Y3) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X2, Y3) < 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
        Else
          'Controlla la orrizzontale a destra
          For X3 = X2 + 1 To Colonna_H
            If DisposizioneScacchiera(X3, Y2) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X3, Y2) < 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
          'Controlla la orrizzontale a sinistra
          For X3 = X2 - 1 To Colonna_A Step -1
            If DisposizioneScacchiera(X3, Y2) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X3, Y2) < 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
        End If
      End If
    End If
  Else
  'Il pezzo Ë Nero
    pezzo = Abs(pezzo)
    If pezzo = Pedone Then
      If Y2 >= Riga_1 Then
        If X2 > Colonna_A Then
          If DisposizioneScacchiera(X2 - 1, Y2 - 1) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        If X2 < Colonna_H Then
          If DisposizioneScacchiera(X2 + 1, Y2 - 1) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    ElseIf pezzo = Cavallo Then
      Y3 = Y2 + 2
      If Y3 <= Riga_8 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2 - 2
      If Y3 >= Riga_1 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 + 2
      If X3 <= Colonna_H Then
        Y3 = Y2 - 1
        If Y3 >= Riga_1 Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        Y3 = Y2 + 1
        If Y3 <= Riga_8 Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 - 2
      If X3 >= Colonna_A Then
        Y3 = Y2 - 1
        If Y3 >= Riga_1 Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
        Y3 = Y2 + 1
        If Y3 <= Riga_8 Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    ElseIf pezzo = RE Then
      Y3 = Y2 + 1
      If Y3 <= Riga_8 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2
        If DisposizioneScacchiera(X3, Y3) > 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2 - 1
      If Y3 >= Riga_1 Then
        X3 = X2 - 1
        If X3 >= Colonna_A Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2 + 1
        If X3 <= Colonna_H Then
          If DisposizioneScacchiera(X3, Y3) > 0 Then
            If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
              NuovaMinaccia = True
              Exit Function
            End If
          End If
        End If
        X3 = X2
        If DisposizioneScacchiera(X3, Y3) > 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      Y3 = Y2
      X3 = X2 - 1
      If X3 >= Colonna_A Then
        If DisposizioneScacchiera(X3, Y3) > 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
      X3 = X2 + 1
      If X3 <= Colonna_H Then
        If DisposizioneScacchiera(X3, Y3) > 0 Then
          If Abs(X3 - X1) > 1 Or Abs(Y3 - Y1) > 1 Then
            NuovaMinaccia = True
            Exit Function
          End If
        End If
      End If
    Else
      If pezzo And Alfiere Then
        If ((X3 - X2) - (Y3 - Y2)) = 0 Then
          'Controlla la diagonale destra
          'In alto
          Y3 = Y2: X3 = X2
          Do While Y3 < Riga_8 And X3 < Colonna_H
            Y3 = Y3 + 1: X3 = X3 + 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) > 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
          'Controlla la diagonale sinistra
          'In basso
          X3 = X2: Y3 = Y2
          Do While Y3 > Riga_1 And X3 > Colonna_A
            Y3 = Y3 - 1:    X3 = X3 - 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) > 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
        Else
          'Controlla la diagonale destra
          'In basso
          Y3 = Y2: X3 = X2
          Do While Y3 > Riga_1 And X3 < Colonna_H
            Y3 = Y3 - 1: X3 = X3 + 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) > 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
          'Controlla la diagonale sinistra
          'In alto
          Y3 = Y2: X3 = X2
          Do While Y3 < Riga_8 And X3 > Colonna_A
            Y3 = Y3 + 1:    X3 = X3 - 1
            If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo Then
                If DisposizioneScacchiera(X3, Y3) > 0 Then
                  NuovaMinaccia = True
                  Exit Function
                End If
              Exit Do
            End If
          Loop
        End If
      End If
      If pezzo And Torre Then
        If Y1 = Y2 Then
          'controlla la verticale sopra
          For Y3 = Y2 + 1 To Riga_8
            If DisposizioneScacchiera(X2, Y3) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X2, Y3) > 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
          'Controlla la verticale sotto
          For Y3 = Y2 - 1 To Riga_1 Step -1
            If DisposizioneScacchiera(X2, Y3) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X2, Y3) > 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
        Else
          'Controlla la orrizzontale a destra
          For X3 = X2 + 1 To Colonna_H
            If DisposizioneScacchiera(X3, Y2) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X3, Y2) > 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
          'Controlla la orrizzontale a sinistra
          For X3 = X2 - 1 To Colonna_A Step -1
            If DisposizioneScacchiera(X3, Y2) <> Nessun_Pezzo Then
              If DisposizioneScacchiera(X3, Y2) > 0 Then
                NuovaMinaccia = True
                Exit Function
              End If
              Exit For
            End If
          Next
        End If
      End If
    End If
  End If
End Function
Public Function InPresa(x As Colonna, Y As Riga, Colore As Colore, DisposizioneScacchiera As Variant) As Genere
  Dim X2 As Colonna
  Dim Y2 As Riga
  If Colore = Bianco Then
    'Pedone
    Y2 = Y - 1
    If Y2 >= Riga_1 Then
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Pedone_Bianco Then
          InPresa = Pedone_Bianco
          Exit Function
        End If
      End If
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Pedone_Bianco Then
          InPresa = Pedone_Bianco
          Exit Function
        End If
      End If
    End If
    'Cavallo
    Y2 = Y + 2
    If Y2 <= Riga_8 Then
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
    End If
    X2 = x + 2
    If X2 <= Colonna_H Then
      Y2 = Y + 1
      If Y2 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
      Y2 = Y - 1
      If Y2 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
    End If
    Y2 = Y - 2
    If Y2 >= Riga_1 Then
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
    End If
    X2 = x - 2
    If X2 >= Colonna_A Then
      Y2 = Y - 1
      If Y2 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
      Y2 = Y + 1
      If Y2 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Bianco Then
          InPresa = Cavallo_Bianco
          Exit Function
        End If
      End If
    End If
    'Re
    X2 = x - 1
    If X2 >= Colonna_A Then
      If DisposizioneScacchiera(X2, Y) = Re_Bianco Then
        InPresa = Re_Bianco
        Exit Function
      End If
      If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y - 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
      End If
      If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y + 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
      End If
    End If
    X2 = x + 1
    If X2 <= Colonna_H Then
      If DisposizioneScacchiera(X2, Y) = Re_Bianco Then
        InPresa = Re_Bianco
        Exit Function
      End If
      If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y - 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
      End If
      If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y + 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
      End If
    End If
    If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(x, Y + 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
    End If
    If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(x, Y - 1) = Re_Bianco Then
          InPresa = Re_Bianco
          Exit Function
        End If
    End If
  Else
  'Allora il colore Ë nero
    'Pedone
    Y2 = Y + 1
    If Y2 <= Riga_8 Then
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Pedone_Nero Then
          InPresa = Pedone_Nero
          Exit Function
        End If
      End If
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Pedone_Nero Then
          InPresa = Pedone_Nero
          Exit Function
        End If
      End If
    End If
    'Cavallo
    Y2 = Y + 2
    If Y2 <= Riga_8 Then
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
    End If
    X2 = x + 2
    If X2 <= Colonna_H Then
      Y2 = Y + 1
      If Y2 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
      Y2 = Y - 1
      If Y2 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
    End If
    Y2 = Y - 2
    If Y2 >= Riga_1 Then
      X2 = x + 1
      If X2 <= Colonna_H Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
      X2 = x - 1
      If X2 >= Colonna_A Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
    End If
    X2 = x - 2
    If X2 >= Colonna_A Then
      Y2 = Y - 1
      If Y2 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
      Y2 = Y + 1
      If Y2 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y2) = Cavallo_Nero Then
          InPresa = Cavallo_Nero
          Exit Function
        End If
      End If
    End If
    'Re
    X2 = x - 1
    If X2 >= Colonna_A Then
      If DisposizioneScacchiera(X2, Y) = Re_Nero Then
        InPresa = Re_Nero
        Exit Function
      End If
      If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y - 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
      End If
      If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y + 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
      End If
    End If
    X2 = x + 1
    If X2 <= Colonna_H Then
      If DisposizioneScacchiera(X2, Y) = Re_Nero Then
        InPresa = Re_Nero
        Exit Function
      End If
      If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(X2, Y - 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
      End If
      If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(X2, Y + 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
      End If
    End If
    If Y + 1 <= Riga_8 Then
        If DisposizioneScacchiera(x, Y + 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
    End If
    If Y - 1 >= Riga_1 Then
        If DisposizioneScacchiera(x, Y - 1) = Re_Nero Then
          InPresa = Re_Nero
          Exit Function
        End If
    End If
  End If
  'Controlla la diagonale destra
  'In alto
  Y2 = Y: X2 = x
  Do While Y2 < Riga_8 And X2 < Colonna_H
    Y2 = Y2 + 1: X2 = X2 + 1
    If DisposizioneScacchiera(X2, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y2) > 0 Then
          If DisposizioneScacchiera(X2, Y2) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y2)) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      End If
      Exit Do
    End If
  Loop
  'In basso
  Y2 = Y: X2 = x
  Do While Y2 > Riga_1 And X2 < Colonna_H
    Y2 = Y2 - 1: X2 = X2 + 1
    If DisposizioneScacchiera(X2, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y2) > 0 Then
          If DisposizioneScacchiera(X2, Y2) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y2)) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      End If
      Exit Do
    End If
  Loop
  'Controlla la diagonale sinistra
  'In basso
  X2 = x: Y2 = Y
  Do While Y2 > Riga_1 And X2 > Colonna_A
    Y2 = Y2 - 1:    X2 = X2 - 1
    If DisposizioneScacchiera(X2, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y2) > 0 Then
          If DisposizioneScacchiera(X2, Y2) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y2)) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      End If
      Exit Do
    End If
  Loop
  'In alto
  Y2 = Y: X2 = x
  Do While Y2 < Riga_8 And X2 > Colonna_A
    Y2 = Y2 + 1:    X2 = X2 - 1
    If DisposizioneScacchiera(X2, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y2) > 0 Then
          If DisposizioneScacchiera(X2, Y2) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y2)) And Alfiere Then
            InPresa = DisposizioneScacchiera(X2, Y2)
            Exit Function
          End If
        End If
      End If
      Exit Do
    End If
  Loop
  'controlla la verticale sopra
  For Y2 = Y + 1 To Riga_8
    If DisposizioneScacchiera(x, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(x, Y2) > 0 Then
          If DisposizioneScacchiera(x, Y2) And Torre Then
            InPresa = DisposizioneScacchiera(x, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(x, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(x, Y2)) And Torre Then
            InPresa = DisposizioneScacchiera(x, Y2)
            Exit Function
          End If
        End If
      End If
      Exit For
    End If
  Next
  'Controlla la verticale sotto
  For Y2 = Y - 1 To Riga_1 Step -1
    If DisposizioneScacchiera(x, Y2) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(x, Y2) > 0 Then
          If DisposizioneScacchiera(x, Y2) And Torre Then
            InPresa = DisposizioneScacchiera(x, Y2)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(x, Y2) < 0 Then
          If Abs(DisposizioneScacchiera(x, Y2)) And Torre Then
            InPresa = DisposizioneScacchiera(x, Y2)
            Exit Function
          End If
        End If
      End If
      Exit For
    End If
  Next
  'Controlla la orrizzontale a destra
  For X2 = x + 1 To Colonna_H
    If DisposizioneScacchiera(X2, Y) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y) > 0 Then
          If DisposizioneScacchiera(X2, Y) And Torre Then
            InPresa = DisposizioneScacchiera(X2, Y)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y)) And Torre Then
            InPresa = DisposizioneScacchiera(X2, Y)
            Exit Function
          End If
        End If
      End If
      Exit For
    End If
  Next
  'Controlla la orrizzontale a sinistra
  For X2 = x - 1 To Colonna_A Step -1
    If DisposizioneScacchiera(X2, Y) <> Nessun_Pezzo Then
      If Colore = Bianco Then
        If DisposizioneScacchiera(X2, Y) > 0 Then
          If DisposizioneScacchiera(X2, Y) And Torre Then
            InPresa = DisposizioneScacchiera(X2, Y)
            Exit Function
          End If
        End If
      Else
        If DisposizioneScacchiera(X2, Y) < 0 Then
          If Abs(DisposizioneScacchiera(X2, Y)) And Torre Then
            InPresa = DisposizioneScacchiera(X2, Y)
            Exit Function
          End If
        End If
      End If
      Exit For
    End If
  Next
End Function
Public Function TutteLeMossePossibili(Colore As Colore, DisposizioneScacchiera As Variant, Mosse As Variant, Quantit‡Pezzi As Integer, Optional XEnPassant As Colonna = NessunaColonna, Optional BiancoArroccoPossibile As Lato, Optional NeroArroccoPossibile As Lato, Optional PosizioneReBianco As Variant, Optional PosizioneReNero As Variant, Optional Scacco As Integer = 2) As Integer
  'XEnPassant indica la possibilit‡ di catturare un pedone en passant nella colonna XEnPassant
  'Il parametro Scacco indica se il re Colore si trova in scacco (True=In Scacco, False=Non in scacco, 2=Non specificato)
  Dim pezzo  As pezzo
  Dim x As Colonna
  Dim Y As Riga
  Dim Quantit‡Mosse As Integer
  Dim MossePezzo As Mosse
  Dim MP() As Variant
  If Scacco = 2 Then
    If Colore = Bianco Then
      Scacco = InPresa(PosizioneReBianco.x, PosizioneReBianco.Y, Nero, DisposizioneScacchiera) <> Nessun_Genere
    Else
      Scacco = InPresa(PosizioneReNero.x, PosizioneReNero.Y, Bianco, DisposizioneScacchiera) <> Nessun_Genere
    End If
  End If
  Quantit‡Pezzi = 0
  If Colore = Bianco Then
    For Y = 0 To 7
      For x = 0 To 7
        If DisposizioneScacchiera(x, Y) > 0 Then
          Quantit‡Pezzi = Quantit‡Pezzi + 1
          ReDim Preserve Mosse(1 To Quantit‡Pezzi)
          pezzo = DisposizioneScacchiera(x, Y)
          MossePezzo.pezzo = pezzo
          MossePezzo.Posizione.x = x
          MossePezzo.Posizione.Y = Y
          Quantit‡Mosse = MossePossibili(x, Y, DisposizioneScacchiera, MP, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero, Scacco)
          TutteLeMossePossibili = TutteLeMossePossibili + Quantit‡Mosse
          MossePezzo.Quantit‡Mosse = Quantit‡Mosse
          If Quantit‡Mosse = 0 Then
            MossePezzo.Mosse = Empty
          Else
            MossePezzo.Mosse = MP
          End If
          Mosse(Quantit‡Pezzi) = MossePezzo
        End If
      Next
    Next
  Else
    For Y = 7 To 0 Step -1
      For x = 0 To 7
        If DisposizioneScacchiera(x, Y) < 0 Then
          Quantit‡Pezzi = Quantit‡Pezzi + 1
          ReDim Preserve Mosse(1 To Quantit‡Pezzi)
          pezzo = DisposizioneScacchiera(x, Y)
          MossePezzo.pezzo = pezzo
          MossePezzo.Posizione.x = x
          MossePezzo.Posizione.Y = Y
          Quantit‡Mosse = MossePossibili(x, Y, DisposizioneScacchiera, MP, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero, Scacco)
          TutteLeMossePossibili = TutteLeMossePossibili + Quantit‡Mosse
          MossePezzo.Quantit‡Mosse = Quantit‡Mosse
          If Quantit‡Mosse = 0 Then
            MossePezzo.Mosse = Empty
          Else
            MossePezzo.Mosse = MP
          End If
          Mosse(Quantit‡Pezzi) = MossePezzo
        End If
      Next
    Next
  End If
End Function
Public Function MossePossibili(x As Colonna, Y As Riga, DisposizioneScacchiera As Variant, Mosse As Variant, Optional XEnPassant As Colonna = NessunaColonna, Optional BiancoArroccoPossibile As Lato, Optional NeroArroccoPossibile As Lato, Optional PosizioneReBianco As Variant, Optional PosizioneReNero As Variant, Optional Scacco As Integer) As Integer
Attribute MossePossibili.VB_Description = "Questa procedura restituisce alla matrice Mosse le case in cui puÚ muoversi il pezzo nella Posizione specificata."
    'X=posizione Riga del pezzo di cui bisogna individuare le mosse
    'Y=posizione Colonna del pezzo di cui bisogna individuare le mosse
    Dim Errore As ErroreMossa
    Dim pezzo As pezzo
    Dim X2 As Colonna
    Dim Y2 As Riga
    Dim Colore As Colore
    pezzo = DisposizioneScacchiera(x, Y)
    If pezzo > 0 Then
      Colore = Bianco
    Else
      Colore = Nero
    End If
    Select Case Abs(pezzo)
      Case Pedone
        Select Case pezzo
          'verifica se il pezzo Ë di colore bianco
          Case Is > 0
            'verifica se Ë possibile muovere il pedone in avanti
            If Y < Riga_8 Then
              Y2 = Y + 1
              If DisposizioneScacchiera(x, Y2) = Nessun_Pezzo Then
                Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                'verifica se Ë possibile muovere il pedone in avanti di due caselle
                If Y = Riga_2 Then
                  If DisposizioneScacchiera(x, 3) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, 3, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  End If
                End If
              End If
              'verifica se Ë possibile che il pedone mangi una pedina a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                'se il pezzo Ë nero allora lo puÚ mangiare
                If DisposizioneScacchiera(X2, Y2) < 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                End If
              End If
              'verifica se Ë possibile che il pedone mangi una pedina a destra
              If x < Colonna_H Then
                X2 = x + 1
                'se il pezzo Ë nero allora lo puÚ mangiare
                If DisposizioneScacchiera(X2, Y2) < 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                End If
              End If
              'Verifica la possibilit‡ di mangiarea un pedone en passant
              If XEnPassant <> NessunaColonna Then
              'verifica se Ë possibile che il pedone mangi en passant una pedina a sinistra
                If x > Colonna_A Then
                  X2 = x - 1
                  If Y = Riga_5 And X2 = XEnPassant Then
                    'se il pezzo Ë un pedone nero allora lo puÚ mangiare
                    If DisposizioneScacchiera(X2, Y) = Pedone_Nero Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    End If
                  End If
                End If
                'verifica se Ë possibile che il pedone mangi en passant una pedina a destra
                If x < Colonna_H Then
                  X2 = x + 1
                  If Y = Riga_5 And X2 = XEnPassant Then
                    'se il pezzo Ë un pedone nero allora lo puÚ mangiare
                    If DisposizioneScacchiera(X2, Y) = Pedone_Nero Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    End If
                  End If
                End If
              End If
            End If
          'il pezzo Ë di colore nero
          Case Else
            'verifica se Ë possibile muovere il pedone in avanti
            If Y > Riga_1 Then
              Y2 = Y - 1
              If DisposizioneScacchiera(x, Y2) = Nessun_Pezzo Then
                Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                'verifica se Ë possibile muovere il pedone in avanti di due caselle
                If Y = Riga_7 Then
                  If DisposizioneScacchiera(x, 4) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, 4, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  End If
                End If
              End If
              'verifica se Ë possibile che il pedone mangi una pedina a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                'se il pezzo Ë bianco allora lo puÚ mangiare
                If DisposizioneScacchiera(X2, Y2) > 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                End If
              End If
              'verifica se Ë possibile che il pedone mangi una pedina a destra
              If x < Colonna_H Then
                X2 = x + 1
                'se il pezzo Ë bianco allora lo puÚ mangiare
                If DisposizioneScacchiera(X2, Y2) > 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                End If
              End If
              'Verifica la possibilit‡ di mangiarea un pedone en passant
              If XEnPassant <> NessunaColonna Then
              'verifica se Ë possibile che il pedone mangi en passant una pedina a sinistra
                If x > Colonna_A Then
                  X2 = x - 1
                  If Y = Riga_4 And X2 = XEnPassant Then
                    'se il pezzo Ë un pedone bianco allora lo puÚ mangiare
                    If DisposizioneScacchiera(X2, Y) = Pedone_Bianco Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    End If
                  End If
                End If
                'verifica se Ë possibile che il pedone mangi en passant una pedina a destra
                If x < Colonna_H Then
                  X2 = x + 1
                  If Y = Riga_4 And X2 = XEnPassant Then
                    'se il pezzo Ë un pedone bianco allora lo puÚ mangiare
                    If DisposizioneScacchiera(X2, Y) = Pedone_Bianco Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    End If
                  End If
                End If
              End If
            End If
        End Select
      Case Cavallo
        Select Case pezzo
          'verifica se il pezzo Ë di colore bianco
          Case Is > 0
            'verifica se il cavallo si puÚ spostare in alto
            If Y <= Riga_6 Then
              Y2 = Y + 2
              'verifica se il cavallo si puÚ spostare in alto a sinistra
              If x <> Colonna_A Then
                X2 = x - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare in alto a destra
              If x <> Colonna_H Then
                X2 = x + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare in basso
            If Y >= Riga_3 Then
              Y2 = Y - 2
              'verifica se il cavallo si puÚ spostare in basso a sinistra
              If x <> Colonna_A Then
                X2 = x - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare in basso a destra
              If x <> Colonna_H Then
                X2 = x + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare a sinistra
            If x >= Colonna_C Then
              X2 = x - 2
              'verifica se il cavallo si puÚ spostare a sinistra in alto
              If Y <> Riga_8 Then
                Y2 = Y + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare a sinistra in basso
              If Y <> Riga_1 Then
                Y2 = Y - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare a destra
            If x <= Colonna_F Then
              X2 = x + 2
              'verifica se il cavallo si puÚ spostare a sinistra in alto
              If Y <> Riga_8 Then
                Y2 = Y + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare a sinistra in basso
              If Y <> Riga_1 Then
                Y2 = Y - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
          'il pezzo Ë di colore nero
          Case Else
            'verifica se il cavallo si puÚ spostare in alto
            Errore = ValoreVuoto
            If Y <= Riga_6 Then
              Y2 = Y + 2
              'verifica se il cavallo si puÚ spostare in alto a sinistra
              If x <> Colonna_A Then
                X2 = x - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare in alto a destra
              If x <> Colonna_H Then
                X2 = x + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare in basso
            If Y >= Riga_3 Then
              Y2 = Y - 2
              'verifica se il cavallo si puÚ spostare in basso a sinistra
              If x <> Colonna_A Then
                X2 = x - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare in basso a destra
              If x <> Colonna_H Then
                X2 = x + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare a sinistra
            If x >= Colonna_C Then
              X2 = x - 2
              'verifica se il cavallo si puÚ spostare a sinistra in alto
              If Y <> Riga_8 Then
                Y2 = Y + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare a sinistra in basso
              If Y <> Riga_1 Then
                Y2 = Y - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
            'verifica se il cavallo si puÚ spostare a destra
            If x <= Colonna_F Then
              X2 = x + 2
              'verifica se il cavallo si puÚ spostare a sinistra in alto
              If Y <> Riga_8 Then
                Y2 = Y + 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
              'verifica se il cavallo si puÚ spostare a sinistra in basso
              If Y <> Riga_1 Then
                Y2 = Y - 1
                'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                  If Errore And 15 Then Exit Function
                End If
              End If
            End If
        End Select
      Case RE
        Select Case pezzo
          'verifica se il Re Ë di colore bianco
          Case Is > 0
            'verifica la possibilit‡ di arroccare
            If x = Colonna_E And Y = Riga_1 Then
              If BiancoArroccoPossibile <> NessunaColonna Then
                If BiancoArroccoPossibile And Sinistra Then
                  If DisposizioneScacchiera(Colonna_D, Riga_1) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_C, Riga_1) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_B, Riga_1) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, Colonna_C, Riga_1, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                End If
                If BiancoArroccoPossibile And Destra Then
                  If DisposizioneScacchiera(Colonna_F, Riga_1) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_G, Riga_1) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, Colonna_G, Riga_1, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                End If
              End If
            End If
            'verifica lo spostamento del re verso il basso
            If Y > Riga_1 Then
              Y2 = Y - 1
              If DisposizioneScacchiera(x, Y2) <= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(x, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Bianco
              End If
              'Verifica lo spostamento del re in basso a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Bianco
                End If
              End If
              'Verifica lo spostamento del re in basso a destra
              If x < Colonna_H Then
                X2 = x + 1
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Bianco
                End If
              End If
            End If
            'verifica lo spostamento del re verso l'alto
            If Y < Riga_8 Then
              Y2 = Y + 1
              If DisposizioneScacchiera(x, Y2) <= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(x, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Bianco
              End If
              'Verifica lo spostamento del re in alto a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Bianco
                End If
              End If
              'Verifica lo spostamento del re in alto a destra
              If x < Colonna_H Then
                X2 = x + 1
                If DisposizioneScacchiera(X2, Y2) <= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Bianco
                End If
              End If
            End If
            'Verifica lo spostamento del re a destra e a sinistra
            'Verifica lo spostamento a destra
            If x > Colonna_A Then
              X2 = x - 1
              If DisposizioneScacchiera(X2, Y) <= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(X2, Y, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Bianco
              End If
            End If
            'Verifica lo spostamento a sinistra
            If x < Colonna_H Then
              X2 = x + 1
              If DisposizioneScacchiera(X2, Y) <= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(X2, Y, Nero, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Bianco
              End If
            End If
          'allora il Re Ë nero
          Case Else
            'verifica la possibilit‡ di arroccare
            If x = Colonna_E And Y = Riga_8 Then
              If NeroArroccoPossibile <> NessunaColonna Then
                If NeroArroccoPossibile And Sinistra Then
                  If DisposizioneScacchiera(Colonna_D, Riga_8) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_C, Riga_8) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_B, Riga_8) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, Colonna_C, Riga_8, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                End If
                If NeroArroccoPossibile And Destra Then
                  If DisposizioneScacchiera(Colonna_F, Riga_8) = Nessun_Pezzo And DisposizioneScacchiera(Colonna_G, Riga_8) = Nessun_Pezzo Then
                    Errore = AggiungiMossa(pezzo, x, Y, Colonna_G, Riga_8, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                End If
              End If
            End If
            'verifica lo spostamento del re verso il basso
            If Y > Riga_1 Then
              Y2 = Y - 1
              If DisposizioneScacchiera(x, Y2) >= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(x, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Nero
              End If
              'Verifica lo spostamento del re in basso a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Nero
                End If
              End If
              'Verifica lo spostamento del re in basso a destra
              If x < Colonna_H Then
                X2 = x + 1
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Nero
                End If
              End If
            End If
            'verifica lo spostamento del re verso l'alto
            If Y < Riga_8 Then
              Y2 = Y + 1
              If DisposizioneScacchiera(x, Y2) >= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(x, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Nero
              End If
              'Verifica lo spostamento del re in alto a sinistra
              If x > Colonna_A Then
                X2 = x - 1
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Nero
                End If
              End If
              'Verifica lo spostamento del re in alto a destra
              If x < Colonna_H Then
                X2 = x + 1
                If DisposizioneScacchiera(X2, Y2) >= 0 Then
                  DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                  If InPresa(X2, Y2, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                  End If
                  DisposizioneScacchiera(x, Y) = Re_Nero
                End If
              End If
            End If
            'Verifica lo spostamento del re a destra e a sinistra
            'Verifica lo spostamento a destra
            If x > Colonna_A Then
              X2 = x - 1
              If DisposizioneScacchiera(X2, Y) >= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(X2, Y, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Nero
              End If
            End If
            'Verifica lo spostamento a sinistra
            If x < Colonna_H Then
              X2 = x + 1
              If DisposizioneScacchiera(X2, Y) >= 0 Then
                DisposizioneScacchiera(x, Y) = Nessun_Pezzo
                If InPresa(X2, Y, Bianco, DisposizioneScacchiera) = Nessun_Genere Then
                  Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, NessunErrore)
                End If
                DisposizioneScacchiera(x, Y) = Re_Nero
              End If
            End If
        End Select
      Case Else
        'alfiere o dama
        If pezzo And Alfiere Then
          Select Case pezzo
            'verifica se il pezzo Ë di colore bianco
            Case Is > 0
              'Prova a spostare l'alfiere in direzione alto destra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 < Colonna_H Then
                  If Y2 < Riga_8 Then
                    X2 = X2 + 1: Y2 = Y2 + 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) <= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione alto sinistra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 > Colonna_A Then
                  If Y2 < Riga_8 Then
                    X2 = X2 - 1: Y2 = Y2 + 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) <= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione basso destra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 < Colonna_H Then
                  If Y2 > Riga_1 Then
                    X2 = X2 + 1: Y2 = Y2 - 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) <= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione basso sinistra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 > Colonna_A Then
                  If Y2 > Riga_1 Then
                    X2 = X2 - 1: Y2 = Y2 - 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) <= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
            'allora il pezzo Ë nero
            Case Else
              'Prova a spostare l'alfiere in direzione alto destra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 < Colonna_H Then
                  If Y2 < Riga_8 Then
                    X2 = X2 + 1: Y2 = Y2 + 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) >= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione alto sinistra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 > Colonna_A Then
                  If Y2 < Riga_8 Then
                    X2 = X2 - 1: Y2 = Y2 + 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) >= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione basso destra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 < Colonna_H Then
                  If Y2 > Riga_1 Then
                    X2 = X2 + 1: Y2 = Y2 - 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) >= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
              'Prova a spostare l'alfiere in direzione basso sinistra
              X2 = x
              Y2 = Y
              Errore = ValoreVuoto
              Do
                If X2 > Colonna_A Then
                  If Y2 > Riga_1 Then
                    X2 = X2 - 1: Y2 = Y2 - 1
                    'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                    If DisposizioneScacchiera(X2, Y2) >= 0 Then
                      Errore = AggiungiMossa(pezzo, x, Y, X2, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                      If Errore And 15 Then Exit Do
                    Else
                      Exit Do
                    End If
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y2) <> 0
          End Select
        End If
        'torre o dama
        If Abs(pezzo) And Torre Then
          Select Case pezzo
            'verifica se il pezzo Ë di colore bianco
            Case Is > 0
              'Prova a spostare la torre in alto
              Y2 = Y
              Errore = ValoreVuoto
              Do
                Y2 = Y2 + 1
                If Y2 <= Riga_8 Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(x, Y2) <= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(x, Y2) <> 0
              'Prova a spostare la torre in basso
              Y2 = Y
              Errore = ValoreVuoto
              Do
                Y2 = Y2 - 1
                If Y2 >= Riga_1 Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(x, Y2) <= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(x, Y2) <> 0
              'Prova a spostare la torre a destra
              X2 = x
              Errore = ValoreVuoto
              Do
                X2 = X2 + 1
                If X2 <= Colonna_H Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(X2, Y) <= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y) <> 0
              'Prova a spostare la torre a sinistra
              X2 = x
              Errore = ValoreVuoto
              Do
                X2 = X2 - 1
                If X2 >= Colonna_A Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(X2, Y) <= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y) <> 0
            'allora il pezzo Ë nero
            Case Else
              'Prova a spostare la torre in alto
              Y2 = Y
              Errore = ValoreVuoto
              Do
                Y2 = Y2 + 1
                If Y2 <= Riga_8 Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(x, Y2) >= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(x, Y2) <> 0
              'Prova a spostare la torre in basso
              Y2 = Y
              Errore = ValoreVuoto
              Do
                Y2 = Y2 - 1
                If Y2 >= Riga_1 Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(x, Y2) >= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, x, Y2, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(x, Y2) <> 0
              'Prova a spostare la torre a destra
              X2 = x
              Errore = ValoreVuoto
              Do
                X2 = X2 + 1
                If X2 <= Colonna_H Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(X2, Y) >= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y) <> 0
              'Prova a spostare la torre a sinistra
              X2 = x
              Errore = ValoreVuoto
              Do
                X2 = X2 - 1
                If X2 >= Colonna_A Then
                  'se nella posizione non vi Ë nessun pezzo o vi Ë un pezzo avversario allora lo spostamento Ë possibile
                  If DisposizioneScacchiera(X2, Y) >= 0 Then
                    Errore = AggiungiMossa(pezzo, x, Y, X2, Y, DisposizioneScacchiera, Mosse, MossePossibili, PosizioneReBianco, PosizioneReNero, Scacco, Colore, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, Errore)
                    If Errore And 15 Then Exit Do
                  Else
                    Exit Do
                  End If
                Else
                  Exit Do
                End If
              Loop Until DisposizioneScacchiera(X2, Y) <> 0
          End Select
        End If
    End Select
End Function
Public Sub EffettuaMossa(X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, DisposizioneScacchiera As Variant, Optional XEnPassant As Colonna = NessunaColonna, Optional BiancoArroccoPossibile As Lato, Optional NeroArroccoPossibile As Lato)
  Dim pezzo As pezzo
  pezzo = DisposizioneScacchiera(X1, Y1)
  If pezzo = Pedone_Bianco Then
    'presa en passant
    If X2 <> X1 Then
      If DisposizioneScacchiera(X2, Y2) = Nessun_Pezzo Then
        DisposizioneScacchiera(X2, Riga_5) = Nessun_Pezzo
      End If
    ElseIf Y2 = Riga_4 Then
      If Y1 = Riga_2 Then
        XEnPassant = X2
      End If
    End If
  ElseIf pezzo = Pedone_Nero Then
    'presa en passant
    If X2 <> X1 Then
      If DisposizioneScacchiera(X2, Y2) = Nessun_Pezzo Then
        DisposizioneScacchiera(X2, Riga_4) = Nessun_Pezzo
      End If
    ElseIf Y2 = Riga_5 Then
      If Y1 = Riga_7 Then
        XEnPassant = X2
      End If
    End If
  ElseIf pezzo = Re_Bianco Then
    If X1 = Colonna_E Then
      If Y1 = Riga_1 Then
        If Y2 = Riga_1 Then
          'arrocco a destra
          If X2 = Colonna_G Then
            DisposizioneScacchiera(Colonna_H, Riga_1) = Nessun_Pezzo
            DisposizioneScacchiera(Colonna_F, Riga_1) = Torre_Bianco
          'arrocco a sinistra
          ElseIf X2 = Colonna_C Then
            DisposizioneScacchiera(Colonna_A, Riga_1) = Nessun_Pezzo
            DisposizioneScacchiera(Colonna_C, Riga_1) = Torre_Bianco
          End If
        End If
      End If
      BiancoArroccoPossibile = 0
    End If
  ElseIf pezzo = Re_Nero Then
    If X1 = Colonna_E Then
      If Y1 = Riga_8 Then
        If Y2 = Riga_8 Then
          'arrocco a destra
          If X2 = Colonna_G Then
            DisposizioneScacchiera(Colonna_H, Riga_8) = Nessun_Pezzo
            DisposizioneScacchiera(Colonna_F, Riga_8) = Torre_Nero
          'arrocco a sinistra
          ElseIf X2 = Colonna_C Then
            DisposizioneScacchiera(Colonna_A, Riga_8) = Nessun_Pezzo
            DisposizioneScacchiera(Colonna_C, Riga_8) = Torre_Nero
          End If
        End If
      End If
      NeroArroccoPossibile = 0
    End If
  ElseIf pezzo = Torre_Bianco Then
    If BiancoArroccoPossibile <> 0 Then
      If Y1 = Riga_1 Then
        If X1 = Colonna_A Then
          BiancoArroccoPossibile = BiancoArroccoPossibile And Not Sinistra
        End If
        If X1 = Colonna_H Then
          BiancoArroccoPossibile = BiancoArroccoPossibile And Not Destra
        End If
      End If
    End If
  ElseIf pezzo = Torre_Nero Then
    If NeroArroccoPossibile <> 0 Then
      If Y1 = Riga_8 Then
        If X1 = Colonna_A Then
          NeroArroccoPossibile = NeroArroccoPossibile And Not Sinistra
        End If
        If X1 = Colonna_H Then
          NeroArroccoPossibile = NeroArroccoPossibile And Not Destra
        End If
      End If
    End If
  End If
  If Abs(pezzo) <> Pedone Then
    XEnPassant = NessunaColonna
  End If
  DisposizioneScacchiera(X2, Y2) = DisposizioneScacchiera(X1, Y1)
  DisposizioneScacchiera(X1, Y1) = Nessun_Pezzo
End Sub
Public Function PosizioneRe(ColoreRe As Colore, DisposizioneScacchiera As Variant) As CoordinateCasa
  Dim x As Colonna
  Dim Y As Riga
  If ColoreRe = Bianco Then
    PosizioneRe.x = Empty
    PosizioneRe.Y = Empty
    For Y = Riga_1 To Riga_8
      For x = Colonna_A To Colonna_H
        If DisposizioneScacchiera(x, Y) = Re_Bianco Then
          PosizioneRe.x = x
          PosizioneRe.Y = Y
          Exit Function
        End If
      Next
    Next
  Else
    PosizioneRe.x = Empty
    PosizioneRe.Y = Empty
    For Y = Riga_8 To Riga_1 Step -1
      For x = Colonna_A To Colonna_H
        If DisposizioneScacchiera(x, Y) = Re_Nero Then
          PosizioneRe.x = x
          PosizioneRe.Y = Y
          Exit Function
        End If
      Next
    Next
  End If
End Function
Public Function InScacco(X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, Turno As Colore, ByVal DisposizioneScacchiera As Variant, ByVal XEnPassant As Colonna, ByVal BiancoArroccoPossibile As Lato, ByVal NeroArroccoPossibile As Lato, PosizioneReBianco As Variant, PosizioneReNero As Variant)
  'Questa funzione restituisce True se dopo aver effettusato la mossa X1,Y1-X2,Y2 il re del colore
  'che ha mosso si trova in scvacco.
  'Se il re si trova in scacco allora la mossa Ë inlegale e non puÚ essere fatta
  EffettuaMossa X1, Y1, X2, Y2, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile
  Dim XPosizioneRe As Colonna
  Dim YPosizioneRe As Riga
  If Turno = Bianco Then
    XPosizioneRe = PosizioneReBianco.x
    YPosizioneRe = PosizioneReBianco.Y
  Else
    XPosizioneRe = PosizioneReNero.x
    YPosizioneRe = PosizioneReNero.Y
  End If
  If X1 = XPosizioneRe Then
    If Y1 = YPosizioneRe Then
      XPosizioneRe = X2
      YPosizioneRe = Y2
    End If
  End If
  If Turno = Bianco Then
    If InPresa(XPosizioneRe, YPosizioneRe, Nero, DisposizioneScacchiera) <> Nessun_Genere Then
      InScacco = True
      Exit Function
    End If
  Else
    If InPresa(XPosizioneRe, YPosizioneRe, Bianco, DisposizioneScacchiera) <> Nessun_Genere Then
      InScacco = True
      Exit Function
    End If
  End If
End Function
Public Function AggiungiMossa(pezzo As pezzo, X1 As Colonna, Y1 As Riga, X2 As Colonna, Y2 As Riga, DisposizioneScacchiera As Variant, Mosse As Variant, Quantit‡Mosse As Integer, PosizioneReBianco As Variant, PosizioneReNero As Variant, Scacco As Integer, Turno As Colore, XEnPassant As Colonna, BiancoArroccoPossibile As Lato, NeroArroccoPossibile As Lato, Optional ErrorePrecedente As ErroreMossa) As ErroreMossa
  'il parametro ErrorePrecedente specifica se i controlli possono essere disabilitati perchË inutili
  'questo parametro viene utilizzato per evitare di ripetere controlli gia effettuati precedentemente
  'in controlli disabilitabili sono quelli innerenti gli scacchi di scoperta
  'nel caso in cui ErrorePrecedente=NessunErrore allora i controlli verranno tralasciati perchË gi‡ effettuati in precedenza
  'mentre se ErrorePrecedente=0 i controlli verranno effettuati perchË significa che non Ë stato effettuato il controllo precedentemente
  
  'Filtra le mosse che mantengono il Re in una situazione di scacco
  If Scacco = True Then
    If True = InScacco(X1, Y1, X2, Y2, Turno, DisposizioneScacchiera, XEnPassant, BiancoArroccoPossibile, NeroArroccoPossibile, PosizioneReBianco, PosizioneReNero) Then
      Exit Function
    End If
  End If
  'filtra le mosse che provocano uno scacco di scoperta
  'controlla se vi Ë uno scacco orizzontale
  Dim X3 As Colonna
  Dim Y3 As Riga
  Dim N As Integer
  If ErrorePrecedente <> NessunErrore Then
    If pezzo > 0 Then
    'si vuole muovere un pezzo bianco
      If pezzo = Re_Bianco Then
      'Verifica se si sta muovendo il Re in una posizione sotto scacco (mossa inlegale)
        If InPresa(X2, Y2, Nero, DisposizioneScacchiera) Then
          Exit Function
        End If
      End If
      If X1 = PosizioneReBianco.x Then
        If X1 <> X2 Then
          If Y1 < PosizioneReBianco.Y Then
          'controlla la verticale sotto il re
            For N = PosizioneReBianco.Y - 1 To Riga_1 Step -1
              If DisposizioneScacchiera(X1, N) <> Nessun_Pezzo And N <> Y1 Then
                If DisposizioneScacchiera(X1, N) = Torre_Nero Or DisposizioneScacchiera(X1, N) = Dama_Nero Then
                  AggiungiMossa = BloccoVerticale
                  Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la verticale sopra il re
            For N = PosizioneReBianco.Y + 1 To Riga_8
              If DisposizioneScacchiera(X1, N) <> Nessun_Pezzo And N <> Y1 Then
                If DisposizioneScacchiera(X1, N) = Torre_Nero Or DisposizioneScacchiera(X1, N) = Dama_Nero Then
                  AggiungiMossa = BloccoVerticale
                  Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf Y1 = PosizioneReBianco.Y Then
        If Y1 <> Y2 Then
          If X1 < PosizioneReBianco.x Then
          'controlla la orizzontale a sinistra del re
            For N = PosizioneReBianco.x - 1 To Colonna_A Step -1
              If DisposizioneScacchiera(N, Y1) <> Nessun_Pezzo And N <> X1 Then
                If DisposizioneScacchiera(N, Y1) = Torre_Nero Or DisposizioneScacchiera(N, Y1) = Dama_Nero Then
                  AggiungiMossa = BloccoOrizzontale
                  Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la orizzontale a destra del re
            For N = PosizioneReBianco.x + 1 To Colonna_H
              If DisposizioneScacchiera(N, Y1) <> Nessun_Pezzo And N <> X1 Then
                If DisposizioneScacchiera(N, Y1) = Torre_Nero Or DisposizioneScacchiera(N, Y1) = Dama_Nero Then
                  AggiungiMossa = BloccoOrizzontale
                  Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf X1 - PosizioneReBianco.x = Y1 - PosizioneReBianco.Y Then
        If X2 - PosizioneReBianco.x <> Y2 - PosizioneReBianco.Y Then
          If Y1 < PosizioneReBianco.Y Then
          'controlla la diagonale sotto il re a 225∞
            For N = 1 To Minore(PosizioneReBianco.x, PosizioneReBianco.Y)
              X3 = PosizioneReBianco.x - N: Y3 = PosizioneReBianco.Y - N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Nero Or DisposizioneScacchiera(X3, Y3) = Dama_Nero Then
                    AggiungiMossa = BloccoDiagonale45g
                    Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la diagonale sopra il re a 45∞
            For N = 1 To 7 - Maggiore(PosizioneReBianco.x, PosizioneReBianco.Y)
              X3 = PosizioneReBianco.x + N: Y3 = PosizioneReBianco.Y + N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Nero Or DisposizioneScacchiera(X3, Y3) = Dama_Nero Then
                    AggiungiMossa = BloccoDiagonale45g
                    Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf X1 - PosizioneReBianco.x = -(Y1 - PosizioneReBianco.Y) Then
        If X2 - PosizioneReBianco.x <> -(Y2 - PosizioneReBianco.Y) Then
          If Y1 < PosizioneReBianco.Y Then
          'controlla la diagonale sotto il re a 135∞
            For N = 1 To Minore(7 - PosizioneReBianco.x, PosizioneReBianco.Y)
              X3 = PosizioneReBianco.x + N: Y3 = PosizioneReBianco.Y - N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Nero Or DisposizioneScacchiera(X3, Y3) = Dama_Nero Then
                    AggiungiMossa = BloccoDiagonale90g
                    Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la diagonale sopra il re a 270∞
            For N = 1 To 7 - Maggiore(7 - PosizioneReBianco.x, PosizioneReBianco.Y)
              X3 = PosizioneReBianco.x - N: Y3 = PosizioneReBianco.Y + N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Nero Or DisposizioneScacchiera(X3, Y3) = Dama_Nero Then
                    AggiungiMossa = BloccoDiagonale90g
                    Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      End If
    Else
    'si vuole muovere un pezzo nero
      If pezzo = Re_Nero Then
      'Verifica se si sta muovendo il Re in una posizione sotto scacco (mossa inlegale)
        If InPresa(X2, Y2, Bianco, DisposizioneScacchiera) Then
          Exit Function
        End If
      End If
      If X1 = PosizioneReNero.x Then
        If X1 <> X2 Then
          If Y1 < PosizioneReNero.Y Then
          'controlla la verticale sotto il re
            For N = PosizioneReNero.Y - 1 To Riga_1 Step -1
              If DisposizioneScacchiera(X1, N) <> Nessun_Pezzo And N <> Y1 Then
                If DisposizioneScacchiera(X1, N) = Torre_Bianco Or DisposizioneScacchiera(X1, N) = Dama_Bianco Then
                  AggiungiMossa = BloccoVerticale
                  Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la verticale sopra il re
            For N = PosizioneReNero.Y + 1 To Riga_8
              If DisposizioneScacchiera(X1, N) <> Nessun_Pezzo And N <> Y1 Then
                If DisposizioneScacchiera(X1, N) = Torre_Bianco Or DisposizioneScacchiera(X1, N) = Dama_Bianco Then
                  AggiungiMossa = BloccoVerticale
                  Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf Y1 = PosizioneReNero.Y Then
        If Y1 <> Y2 Then
          If X1 < PosizioneReNero.x Then
          'controlla la orizzontale a sinistra del re
            For N = PosizioneReNero.x - 1 To Colonna_A Step -1
              If DisposizioneScacchiera(N, Y1) <> Nessun_Pezzo And N <> X1 Then
                If DisposizioneScacchiera(N, Y1) = Torre_Bianco Or DisposizioneScacchiera(N, Y1) = Dama_Bianco Then
                  AggiungiMossa = BloccoOrizzontale
                  Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la orizzontale a destra del re
            For N = PosizioneReNero.x + 1 To Colonna_H
              If DisposizioneScacchiera(N, Y1) <> Nessun_Pezzo And N <> X1 Then
                If DisposizioneScacchiera(N, Y1) = Torre_Bianco Or DisposizioneScacchiera(N, Y1) = Dama_Bianco Then
                  AggiungiMossa = BloccoOrizzontale
                  Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf X1 - PosizioneReNero.x = Y1 - PosizioneReNero.Y Then
        If X2 - PosizioneReNero.x <> Y2 - PosizioneReNero.Y Then
          If Y1 < PosizioneReNero.Y Then
          'controlla la diagonale sotto il re a 225∞
            For N = 1 To Minore(PosizioneReNero.x, PosizioneReNero.Y)
              X3 = PosizioneReNero.x - N: Y3 = PosizioneReNero.Y - N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Bianco Or DisposizioneScacchiera(X3, Y3) = Dama_Bianco Then
                    AggiungiMossa = BloccoDiagonale45g
                    Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la diagonale sopra il re a 45∞
            For N = 1 To 7 - Maggiore(PosizioneReNero.x, PosizioneReNero.Y)
              X3 = PosizioneReNero.x + N: Y3 = PosizioneReNero.Y + N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Bianco Or DisposizioneScacchiera(X3, Y3) = Dama_Bianco Then
                    AggiungiMossa = BloccoDiagonale45g
                    Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      ElseIf X1 - PosizioneReNero.x = -(Y1 - PosizioneReNero.Y) Then
        If X2 - PosizioneReNero.x <> -(Y2 - PosizioneReNero.Y) Then
          If Y1 < PosizioneReNero.Y Then
          'controlla la diagonale sotto il re a 135∞
            For N = 1 To Minore(7 - PosizioneReNero.x, PosizioneReNero.Y)
              X3 = PosizioneReNero.x + N: Y3 = PosizioneReNero.Y - N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Bianco Or DisposizioneScacchiera(X3, Y3) = Dama_Bianco Then
                    AggiungiMossa = BloccoDiagonale90g
                    Exit Function
                End If
                Exit For
              End If
            Next
          Else
          'controlla la diagonale sopra il re a 270∞
            For N = 1 To 7 - Maggiore(7 - PosizioneReNero.x, PosizioneReNero.Y)
              X3 = PosizioneReNero.x - N: Y3 = PosizioneReNero.Y + N
              If DisposizioneScacchiera(X3, Y3) <> Nessun_Pezzo And X3 <> X1 Then
                If DisposizioneScacchiera(X3, Y3) = Alfiere_Bianco Or DisposizioneScacchiera(X3, Y3) = Dama_Bianco Then
                    AggiungiMossa = BloccoDiagonale90g
                    Exit Function
                End If
                Exit For
              End If
            Next
          End If
        End If
      End If
    End If
  End If
  AggiungiMossa = NessunErrore
  Dim Mossa As CoordinateCasa
  Quantit‡Mosse = Quantit‡Mosse + 1
  Mossa.x = X2
  Mossa.Y = Y2
  ReDim Preserve Mosse(1 To Quantit‡Mosse)
  Mosse(Quantit‡Mosse) = Mossa
End Function
Function Minore(n1 As Integer, N2 As Integer) As Integer
  If n1 < N2 Then
    Minore = n1
  Else
    Minore = N2
  End If
End Function
Function Maggiore(n1 As Integer, N2 As Integer) As Integer
  If n1 > N2 Then
    Maggiore = n1
  Else
    Maggiore = N2
  End If
End Function

Private Sub Class_Terminate()
  Set ArcS = Nothing
  Set PG = Nothing
End Sub
