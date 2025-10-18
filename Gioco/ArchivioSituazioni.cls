VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArchivioSituazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CodiceSituazione"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
DefInt A-Z
Option Explicit
Const ByteLong = 4
Const MaxSituazioni = 10000000
Const Msg1 = "File inesistente"
Const GruppiOrdine = 20
Dim OrdineSituazioni() As Long
Dim OrdineSituazioniGruppi() As Long
Dim Quantit‡PerGruppo() As Long
Dim Quantit‡Record As Long
Dim FileSituazioni As String
Dim FilePartite As String
Dim NFS As Integer
Dim NFP As Integer
Private Modalit‡AggiornamentoArchivio As Boolean
Private CS As CodiceSituazione
Public Sub AggiungiSituazione(DisposizioneScacchiera As Variant, Turno As EditorScacchiera, BiancoArroccoPossibile As Lato, NeroArroccoPossibile As Lato, XEnPassant As Colonna, ByVal RisultatoPartita As Risultato, SeekPartita As Long)
  If RisultatoPartita <> RisultatoNonSpecificato Then
    Dim RuotatoSpeculare As Analogia 'Questa variabile Ë un parametro che restituir‡ se la Situazione verr‡ generata a ColoriInvertiti, in talcaso la procedura terr‡ conto che il Risultato Ë capovolto.
    Dim Situazione As String
    Situazione = CodiceSituazione.Situazione(DisposizioneScacchiera, Turno, BiancoArroccoPossibile, NeroArroccoPossibile, XEnPassant, RuotatoSpeculare)
    'La variabile RuotatoSpeculare permette di capovolgere il risultato nel caso la Situazione sia a ColoriInvertiti
    If RuotatoSpeculare And ColoriInvertiti Then
      If RisultatoPartita = VittoriaBianco Then
        RisultatoPartita = VittoriaNero
      ElseIf RisultatoPartita = VittoriaNero Then
        RisultatoPartita = VittoriaBianco
      End If
    End If
    Dim IdxOrdine As Long
    If EsisteSituazione(Situazione, IdxOrdine) Then
      UnisciSituazione IdxOrdine, SeekPartita, RisultatoPartita
    Else
      Dim Posizione As Long
      Posizione = LOF(NFS) + 1
'If (Posizione <= 1) And Quantit‡Record <> 0 Then
'  MsgBox ("Impossibile!")
'End If
      Put #NFS, Posizione, Situazione
      Dim s As String
      If RisultatoPartita = VittoriaBianco Then
        s = Chr$(1) + Chr$(0) + Chr$(0)
      ElseIf RisultatoPartita = Pareggio Then
        s = Chr$(0) + Chr$(1) + Chr$(0)
      ElseIf RisultatoPartita = VittoriaNero Then
        s = Chr$(0) + Chr$(0) + Chr$(1)
      End If
      Put #NFS, , s
      Put #NFS, , SeekPartita
      Dim N As Long
      If Modalit‡AggiornamentoArchivio Then
        'l'inserimento di un record in questo modo Ë pi˘ veloce
        Dim Gruppo As Long
        Dim Index As Long
        Index = IdxOrdine
        Gruppo = IndividuaGruppo(Index)
        For N = Quantit‡PerGruppo(Gruppo) To Index Step -1
          OrdineSituazioniGruppi(Gruppo, N + 1) = OrdineSituazioniGruppi(Gruppo, N)
        Next
        Quantit‡PerGruppo(Gruppo) = Quantit‡PerGruppo(Gruppo) + 1
      Else
        For N = Quantit‡Record To IdxOrdine Step -1
          OrdineSituazioni(N + 1) = OrdineSituazioni(N)
        Next
      End If
      Ordine(IdxOrdine) = Posizione
      Quantit‡Record = Quantit‡Record + 1
    End If
  End If
'  If EsisteSituazione(Situazione) = False Then
'    MsgBox ("Algoritmo errato")
'  End If
End Sub
Public Sub UnisciSituazione(IdxOrdine As Long, SeekPartita As Long, RisultatoPartita As Risultato)
  If RisultatoPartita <> RisultatoNonSpecificato Then
    'Seek NFS, Ordine(IdxOrdine)
    Dim s As String
    s = String(8, " ")
    Get #NFS, Ordine(IdxOrdine), s
    Dim L As Integer
    L = CS.LenPezzi(s) + 1 'si aggiunge 1 perchgË bisogna tener conto del byte indicante l'arrocco e l' en passant
    Dim PosizioneRisultati As Long
    PosizioneRisultati = Seek(NFS) + L
    'Seek NFS, PosizioneRisultati
    Dim R As String
    R = String(3, " ")
    Get #NFS, PosizioneRisultati, R
    Dim C1 As Integer
    Dim C2 As Integer
    Dim C3 As Integer
    C1 = Asc(Mid(R, 1, 1))
    C2 = Asc(Mid(R, 2, 1))
    C3 = Asc(Mid(R, 3, 1))
    Dim NS As Integer
    NS = C1 + C2 + C3
    s = ""
    If RisultatoPartita = VittoriaBianco Then
      If C1 = 255 Then
        Exit Sub
      End If
      s = Chr(C1 + 1) + Chr$(C2) + Chr$(C3)
    ElseIf RisultatoPartita = Pareggio Then
      If C2 = 255 Then
        Exit Sub
      End If
      s = Chr(C1) + Chr$(C2 + 1) + Chr$(C3)
    ElseIf RisultatoPartita = VittoriaNero Then
      If C3 = 255 Then
        Exit Sub
      End If
      s = Chr(C1) + Chr$(C2) + Chr$(C3 + 1)
    End If
    'Seek NFS, PosizioneRisultati
    Put #NFS, PosizioneRisultati, s
    Dim P As Long
    P = PosizioneRisultati + 3 + NS * ByteLong
    If (P + 8) <= LOF(NFS) Then
      'Seek NFS, P
      s = String(8, " ")
      Get #NFS, P, s
      If s <> String(8, Chr(0)) Then
        Dim NextIdx As Long
        NextIdx = IdxOrdineSuccessivo(IdxOrdine)
        If Ordine(NextIdx) <= (P + 4) Then
          SpostaSituazioneInFondo (NextIdx)
        End If
      End If
    End If
    'Seek NFS, P
    Put #NFS, P, SeekPartita
  End If
End Sub
Public Function IdxOrdineSuccessivo(IdxOrdine As Long) As Long
Attribute IdxOrdineSuccessivo.VB_Description = "Restituisce la Index di Ordine del record successivo a quello in posizione Ordine(IdxOrdine)."
  'Restituisce l'Idx del record successivo a IdxOrdine
  'Restituisce 0 se non vi Ë nessun record dopo questo
  If Quantit‡Record <> 0 Then
    Dim N As Long
    Dim IOS As Long
    Dim Posizione As Long
    Posizione = Ordine(IdxOrdine)
    If Modalit‡AggiornamentoArchivio Then
      Dim G As Long
      Dim i As Long
      Dim UltimaPosizione As Long
      For G = 1 To GruppiOrdine
        For i = 1 To Quantit‡PerGruppo(G)
          N = N + 1
          If OrdineSituazioniGruppi(G, i) > Posizione Then
            'In questo modo si evita l'errore generato quando IOS=0 e si vuole leggere la matrice Ordine(IOS)
            If IOS = 0 Then
              IOS = N
              UltimaPosizione = OrdineSituazioniGruppi(G, i)
            ElseIf (OrdineSituazioniGruppi(G, i) < UltimaPosizione) Then
              IOS = N
              UltimaPosizione = OrdineSituazioniGruppi(G, i)
            End If
           End If
        Next
      Next
    Else
      For N = 1 To Quantit‡Record
        If OrdineSituazioni(N) > Posizione Then
          'In questo modo si evita l'errore generato quando IOS=0 e si vuole leggere la matrice Ordine(IOS)
          If IOS = 0 Then
            IOS = N
          ElseIf (OrdineSituazioni(N) < OrdineSituazioni(IOS)) Then
            IOS = N
          End If
        End If
      Next
    End If
    IdxOrdineSuccessivo = IOS
  End If
End Function
Public Sub SpostaSituazioneInFondo(IdxOrdine As Long)
    Dim Rec As String
    Rec = CaricaSituazione(IdxOrdine)
    Dim R As String
    R = String(3, " ")
    Get #NFS, , R
    Rec = Rec + R
    Dim C1 As Integer
    Dim C2 As Integer
    Dim C3 As Integer
    C1 = Asc(Mid(R, 1, 1))
    C2 = Asc(Mid(R, 2, 1))
    C3 = Asc(Mid(R, 3, 1))
    Dim NS As Integer
    NS = C1 + C2 + C3
    Dim s As String
    s = String(NS * ByteLong, " ")
    Get #NFS, , s
    Rec = Rec + s
    'cancella record spostato
    'Seek NFS, Ordine(IdxOrdine)
    s = String(Len(Rec), Chr(0))
    Put #NFS, Ordine(IdxOrdine), s
    'scrivi record in fondo
    Dim Posizione As Long
    Posizione = LOF(NFS) + 1
    Ordine(IdxOrdine) = Posizione
    'Seek NFS, Posizione
    Put #NFS, Posizione, Rec
End Sub
Public Function Risultati(IdxOrdine As Long) As String
Attribute Risultati.VB_Description = "Restituisce una stringa di 3 byte contenente i risultati delle partite indicate lel record IdxOrdine. byte 1 = Vittorie, byte 2 = pareggi, byte 3 =  Sconfitte"
  'restituisce una stringa di 3 byte contenente i risultati delle partite indicate lel record IdxOrdine
  'byte 1 = Vittorie, byte 2 = pareggi, byte 3 = sconfitte
  Dim s As String
  s = CaricaSituazione(IdxOrdine)
  Risultati = CaricaDaFile(3)
End Function
Public Function CaricaDaFile(Qbyte As Long) As String
  CaricaDaFile = String(Qbyte, " ")
  Get #NFS, , CaricaDaFile
End Function
Public Function EsisteSituazione(Situazione As String, Optional IdxOrdine As Long) As Boolean
Attribute EsisteSituazione.VB_Description = "Restituishe se vi Ë gia in archivio una situazione identica, il parametro restituische l' IdxOrdine di questa situazione o l'eventuale Index che assumer‡ in caso di inserimento."
  'Il parametro IdxOrdine restituisce la posizione della situazione (vero o eventuale).
  Dim Rec As String
  Dim Minimo As Long
  Dim Massimo As Long
  Dim Met‡ As Long
  If Quantit‡Record <> 0 Then
    Minimo = 1
    Massimo = Quantit‡Record
    Do
      Met‡ = Intermedio(Minimo, Massimo)
      Rec = CaricaSituazione(Met‡)
      If Minimo = Massimo Then
        If Rec = Situazione Then
          EsisteSituazione = True
          IdxOrdine = Met‡
          Exit Function
        Else
          If Rec < Situazione Then
            EsisteSituazione = False
            IdxOrdine = Met‡ + 1
            Exit Function
          Else '(Rec > Situazione)
            EsisteSituazione = False
            IdxOrdine = Met‡
            Exit Function
          End If
        End If
      Else
        If Rec < Situazione Then
          Minimo = Met‡ + 1
          If Minimo > Massimo Then
            Minimo = Massimo
          End If
        ElseIf Rec > Situazione Then
          Massimo = Met‡ - 1
          If Massimo < Minimo Then
            Minimo = Massimo
          End If
        Else '(Rec = Situazione)
          If Rec = Situazione Then
          EsisteSituazione = True
          IdxOrdine = Met‡
          Exit Function
          End If
        End If
      End If
    Loop
  Else
    IdxOrdine = 1
    EsisteSituazione = False
  End If
End Function
Private Function CaricaSituazione(Optional IdxOrdine As Long, Optional Posizione As Long) As String
    Dim Rec As String
    'Seek NFS, Ordine(IdxOrdine)
    Dim s As String
    s = String(8, " ")
    If Posizione Then
      Get #NFS, Posizione, s
    Else
      Get #NFS, Ordine(IdxOrdine), s
    End If
    Rec = s
    Dim L As Integer
    L = CS.LenPezzi(s) + 1 'si aggiunge 1 perchgË bisogna tener conto del byte indicante l'arrocco e l' en passant
    s = String(L, " ")
    Get #NFS, , s
    CaricaSituazione = Rec + s
End Function
Private Function Intermedio(ByVal Valore1 As Long, ByVal Valore2 As Long) As Long
  If Valore2 < Valore1 Then
    Dim tmp As Long
    tmp = Valore1
    Valore1 = Valore2
    Valore2 = tmp
  End If
  Intermedio = Valore1 + ((Valore2 - Valore1) / 2)
End Function
Public Property Get CodiceSituazione() As CodiceSituazione
    If CS Is Nothing Then
        Set CS = New CodiceSituazione
    End If
    Set CodiceSituazione = CS
End Property


Public Property Set CodiceSituazione(vData As CodiceSituazione)
    Set CS = vData
End Property

Private Sub Class_Initialize()
  FileSituazioni = "Archivio\Archivio.sit"
  FilePartite = "Archivio\Archivio.pgn"
End Sub
Public Sub ApriFilePartite(Optional File As String)
  If File = "" Then
    File = FilePartite
  End If
  NFP = FreeFile
  If Dir(File) <> "" Then
    Open File For Input As NFP
    FilePartite = File
  Else
    MsgBox Msg1 + " (" + File + ")"
  End If
End Sub
Public Sub ChiudiFilePartite()
  Close #NFP
End Sub
Public Sub ApriFileSituazioni(Optional File As String, Optional Modalit‡ As Integer)
  'Valori validi per Modalit‡:
  '0 = I dati verranno solamente letti e mai scritti (Utilizzo Ricerca)
  '1 = I dati verranno distrutti, sar‡ poi possibile ricrearli (Utilizzo Creazione Database)
  '2 = I dati non verranno distrutti ed Ë possibile aggiungerli (Utilizzo Aggiornamento Database)
  Dim OldF As String
  OldF = FileSituazioni
  If File = "" Then
    File = FileSituazioni
  ElseIf InStr(UCase(File), ".SIT") = 0 Then
    File = File + ".sit"
  End If
  FileSituazioni = File
  If Modalit‡ = 1 Then
    If Dir(File) <> "" Then
      Kill File
    End If
    ReDim OrdineSituazioni(1 To MaxSituazioni)
  ElseIf Modalit‡ = 0 Then
    If (Quantit‡Record = 0) Or (UCase(OldF) <> UCase(File)) Then
      CaricaOrdine
    End If
    If Dir(File) = "" Then
      FileSituazioni = ""
      Exit Sub
    End If
  ElseIf Modalit‡ = 2 Then
    CaricaOrdine
    If Quantit‡Record < 1000 * GruppiOrdine Then
      ReDim Preserve OrdineSituazioni(1 To MaxSituazioni)
    Else
      AttivaModalit‡AggiornamentoArchivio
    End If
  End If
  NFS = FreeFile
  Open File For Binary As NFS
End Sub
Public Sub ChiudiFileSituazioni()
  Close #NFS
End Sub
Private Sub Class_Terminate()
    Set CS = Nothing
End Sub
Public Sub SalvaOrdine()
  Dim N As Long
  If Quantit‡Record > 0 Then
    Dim FF As Integer
    Dim File As String
    FF = FreeFile
    File = Left(FileSituazioni, InStr(UCase(FileSituazioni), ".SIT")) + "ord"
    If Dir(File) <> "" Then
      Kill File
    End If
    Open File For Binary As FF
    For N = 1 To Quantit‡Record
      Put #FF, , Ordine(N)
    Next
    Close #FF
  End If
End Sub
Public Sub CaricaOrdine()
    Dim FF As Integer
    Dim File As String
    FF = FreeFile
    File = Left(FileSituazioni, InStr(UCase(FileSituazioni), ".SIT")) + "ord"
    If Dir(File) <> "" Then
      Open File For Binary As FF
      Quantit‡Record = Quantit‡Record + LOF(FF) / 4
      ReDim Preserve OrdineSituazioni(1 To Quantit‡Record)
      Dim N As Long
      Dim LNG As Long
      'For N = 1 To Quantit‡Record
      '  Get #FF, , LNG
      '  Ordine(N) = LNG
      'Next
      
      Get #FF, , OrdineSituazioni()
            
      
      Close #FF
    End If
End Sub
Private Property Get Ordine(Index As Long) As Long
  If Modalit‡AggiornamentoArchivio Then
    'OrdineSituazioniGruppi(Gruppo,index)
    'Quantit‡PerGruppo(Gruppo)
    Dim Gruppo As Long
    Dim Idx As Long
    Idx = Index
    Gruppo = IndividuaGruppo(Idx)
    Ordine = OrdineSituazioniGruppi(Gruppo, Idx)
  Else
    Ordine = OrdineSituazioni(Index)
  End If
End Property
Private Property Let Ordine(Index As Long, ByVal New_Ordine As Long)
  If Modalit‡AggiornamentoArchivio Then
    Dim Idx As Long
    Idx = Index
    Dim Gruppo As Long
    Gruppo = IndividuaGruppo(Idx)
    OrdineSituazioniGruppi(Gruppo, Idx) = New_Ordine
  Else
    OrdineSituazioni(Index) = New_Ordine
  End If
End Property
Private Function IndividuaGruppo(Index As Long) As Long
    Dim gr As Long
    Dim q As Long
    Dim OldQ As Long
    For gr = 1 To GruppiOrdine
      OldQ = q
      q = q + Quantit‡PerGruppo(gr)
      If q >= Index Then
        IndividuaGruppo = gr
        Index = Index - OldQ
        Exit Function
      End If
    Next
End Function
Private Sub AttivaModalit‡AggiornamentoArchivio()
    Dim N As Long
    ReDim Quantit‡PerGruppo(1 To GruppiOrdine)
    ReDim OrdineSituazioniGruppi(1 To GruppiOrdine, 1 To (MaxSituazioni / GruppiOrdine))
    Dim Suddivisione As Long
    Suddivisione = Quantit‡Record \ GruppiOrdine
    Dim gr As Long
    For N = 1 To Quantit‡Record
      gr = 1 + Int(N / Suddivisione)
      If gr > GruppiOrdine Then
        gr = GruppiOrdine
      End If
      Quantit‡PerGruppo(gr) = Quantit‡PerGruppo(gr) + 1
      OrdineSituazioniGruppi(gr, Quantit‡PerGruppo(gr)) = OrdineSituazioni(N)
    Next
    Modalit‡AggiornamentoArchivio = True
    Erase OrdineSituazioni()
End Sub
