VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PGN"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const Msg1 = "Impossibile da interpretare: "
Dim SV As New ScacchieraVirtuale
Dim calcolo As New calcolo

Public Function FEN(Scacchiera As Variant) As String
    Dim x As Colonna
    Dim Y As Riga
    Dim N As Integer
    Dim DS As Variant
    DS = DS
    DS = Scacchiera.DisposizioneScacchiera
    For Y = Riga_8 To Riga_1 Step -1
        N = 0
        For x = Colonna_A To Colonna_H
            If DS(x, Y) <> Nessun_Pezzo And N <> 0 Then
              FEN = FEN + LTrim(Str(N))
              N = 0
            End If
            Select Case DS(x, Y)
                Case Pedone_Bianco
                   FEN = FEN + "P"
                Case Pedone_Nero
                   FEN = FEN + "p"
                Case Cavallo_Bianco
                   FEN = FEN + "N"
                Case Cavallo_Nero
                   FEN = FEN + "n"
                Case Alfiere_Bianco
                   FEN = FEN + "B"
                Case Alfiere_Nero
                   FEN = FEN + "b"
                Case Torre_Bianco
                   FEN = FEN + "R"
                Case Torre_Nero
                   FEN = FEN + "r"
                Case Dama_Bianco
                   FEN = FEN + "Q"
                Case Dama_Nero
                   FEN = FEN + "q"
                Case Re_Bianco
                   FEN = FEN + "K"
                Case Re_Nero
                   FEN = FEN + "k"
                Case Else
                    N = N + 1
            End Select
        Next
        If N <> 0 Then
          FEN = FEN + LTrim(Str(N))
          N = 0
        End If
        If Y > Riga_1 Then
          FEN = FEN + "/"
        End If
    Next
    Select Case Scacchiera.Editabilit‡
        Case Bianco
          FEN = FEN + " w"
        Case Else
          FEN = FEN + " b"
    End Select
    Dim Arrocco As String
    If Scacchiera.NeroArroccoPossibile Or Destra Then
        Arrocco = Arrocco + "K"
    End If
    If Scacchiera.NeroArroccoPossibile Or Sinistra Then
        Arrocco = Arrocco + "Q"
    End If
    If Scacchiera.BiancoArroccoPossibile Or Destra Then
        Arrocco = Arrocco + "k"
    End If
    If Scacchiera.BiancoArroccoPossibile Or Sinistra Then
        Arrocco = Arrocco + "q"
    End If
    If Arrocco <> "" Then
        FEN = FEN + " " + Arrocco
    Else
        FEN = FEN + " -"
    End If
    Dim EnPassant As String
    'verifica se scacchiera.XEnPassant Ë impostato ma non vi Ë nessun pezzo che puÚ effettuare la presa allora scacchiera.XEnPassant va annullato
    Dim FL As Boolean
    If Scacchiera.XEnPassant <> NessunaColonna Then
      If Scacchiera.Editabilit‡ = Nero Then
        If Scacchiera.XEnPassant > Colonna_A Then
          If DS(Scacchiera.XEnPassant - 1, 3) = Pedone_Nero Then
            FL = True
          End If
        End If
        If Scacchiera.XEnPassant < Colonna_H Then
          If DS(Scacchiera.XEnPassant + 1, 3) = Pedone_Nero Then
            FL = True
          End If
        End If
      Else
        If Scacchiera.XEnPassant > Colonna_A Then
          If DS(Scacchiera.XEnPassant - 1, 6) = Pedone_Bianco Then
            FL = True
          End If
        End If
        If Scacchiera.XEnPassant < Colonna_H Then
          If DS(Scacchiera.XEnPassant + 1, 6) = Pedone_Bianco Then
            FL = True
          End If
        End If
      End If
      If FL = False Then
        Scacchiera.XEnPassant = NessunaColonna
      End If
    End If
    If Scacchiera.XEnPassant <> NessunaColonna Then
        Select Case Scacchiera.XEnPassant
            Case Colonna_A
                EnPassant = "a"
            Case Colonna_B
                EnPassant = "b"
            Case Colonna_C
                EnPassant = "c"
            Case Colonna_D
                EnPassant = "d"
            Case Colonna_E
                EnPassant = "e"
            Case Colonna_F
                EnPassant = "f"
            Case Colonna_G
                EnPassant = "g"
            Case Colonna_H
                EnPassant = "h"
        End Select
        If Scacchiera.Editabilit‡ = Bianco Then
            EnPassant = EnPassant + "6"
        Else
            EnPassant = EnPassant + "3"
        End If
    End If
    If EnPassant <> "" Then
        FEN = FEN + " " + EnPassant
    Else
        FEN = FEN + " -"
    End If
    FEN = FEN + " 0" + Str(Scacchiera.MossaAttuale)
End Function
Public Sub CaricaPartita(NomeFile As String, SeekRecord As Long, Scacchiera As Variant, Optional DisabilitaMSG As Boolean = False)
  'Carica la la partita contenuta in NomeFile alla posizione SeekRecord
  Dim FLPos As Boolean
  Dim NumeroFile As Integer
  Dim L As String
  Dim FL As Boolean
  Dim xep As Colonna
  Dim OLDScaAVP As Boolean
  xep = NessunaColonna
  OLDScaAVP = Scacchiera.AbilitaVisualizzazionePezzi
  Scacchiera.AbilitaVisualizzazionePezzi = False
  
  'In caso di Debig Abilitare le righe seguenti:
  'Scacchiera.AbilitaVisualizzazionePezzi = True
  'FRMScacchiera.Aggiorna
  'Scacchiera.Visible = True
  
  Scacchiera.RegistraTempo = False
  NumeroFile = FreeFile
  Open NomeFile For Input As #NumeroFile
  Seek NumeroFile, SeekRecord
  While Seek(NumeroFile) < LOF(NumeroFile)
    Line Input #NumeroFile, L
    Dim s As Variant
    Dim P As Integer
    Dim P2 As Integer
    Dim D As String
    P = InStr(L, Chr(34))
    If (P <> 0) And Left(L, 1) = "[" Then
      D = UCase(Left(L, P - 1))
      P2 = InStr(P + 1, L, Chr(34))
      s = Mid(L, P + 1, P2 - P - 1)
      If s = "?" Then
        s = ""
      End If
      Select Case D
        Case "[EVENT "
          If FL = True Then
            GoTo fine
          End If
          Scacchiera.Evento = s
          FL = True
        Case "[SITE "
          Scacchiera.EventoLuogo = s
        Case "[DATE "
          On Local Error Resume Next
          Scacchiera.EventoData = DateValue(Sostituisci(Right(s, 2) + "\" + Mid(s, 6, 2) + "\" + Left(s, 4), "?", "0"))
          If Err <> 0 Then
            Scacchiera.EventoData = DateValue("01/01/" + Left(s, 4))
            Err = 0
          End If
          On Error GoTo 0
        Case "[ROUND "
          Scacchiera.eventoincontro = CInt(Val(s))
        Case "[WHITE "
          Scacchiera.BiancoGiocatore = s
        Case "[BLACK "
          Scacchiera.NeroGiocatore = s
        Case "[RESULT "
          Select Case s
            Case "1-0"
              Scacchiera.EventoRisultato = 1
            Case "0-1"
              Scacchiera.EventoRisultato = 2
            Case "1/2-1/2"
              Scacchiera.EventoRisultato = 3
            Case Else
              Scacchiera.EventoRisultato = 0
          End Select
        Case "[ECO "
          Scacchiera.EventoECO = s
        Case "[WHITEELO "
          Scacchiera.BiancoElo = s
        Case "[BLACKELO "
          Scacchiera.NeroElo = s
        Case "[WHITEUSCF "
          Scacchiera.BiancoUSCF = s
        Case "[BLACKUSCF "
          Scacchiera.NeroUSCF = s
        Case "[WHITETITLE "
          Select Case UCase(s)
            Case "FM"
              Scacchiera.BiancoTitolo = 3
            Case "IM"
              Scacchiera.BiancoTitolo = 2
            Case "GM"
              Scacchiera.BiancoTitolo = 1
            Case Else
              Scacchiera.BiancoTitolo = 0
          End Select
        Case "[BLACKTITLE "
          Select Case UCase(s)
            Case "FM"
              Scacchiera.NeroTitolo = 3
            Case "IM"
              Scacchiera.NeroTitolo = 2
            Case "GM"
              Scacchiera.NeroTitolo = 1
            Case Else
              Scacchiera.NeroTitolo = 0
          End Select
        Case "[WHITENA "
          Scacchiera.BiancoEM = s
        Case "[BLACKNA "
          Scacchiera.NeroEM = s
        Case "[WHITETYPE "
          Select Case UCase(s)
            Case "HUMAN"
              Scacchiera.BiancoTipoGiocatore = 1
            Case "PROGRAM"
              Scacchiera.BiancoTipoGiocatore = 2
            Case "ROBOT"
              Scacchiera.BiancoTipoGiocatore = 3
            Case "AUTOMA"
              Scacchiera.BiancoTipoGiocatore = 4
            Case "CYBORG"
              Scacchiera.BiancoTipoGiocatore = 5
            Case "DROID"
              Scacchiera.BiancoTipoGiocatore = 6
            Case "MUTANT"
              Scacchiera.BiancoTipoGiocatore = 7
            Case "CLONE"
              Scacchiera.BiancoTipoGiocatore = 8
            Case "ALIEN"
              Scacchiera.BiancoTipoGiocatore = 9
            Case "ANIMAL"
              Scacchiera.BiancoTipoGiocatore = 10
            Case "UNKNOW ENTITY"
              Scacchiera.BiancoTipoGiocatore = 11
            Case Else
              Scacchiera.BiancoTipoGiocatore = 0
          End Select
        Case "[BLACKTYPE "
          Select Case UCase(s)
            Case "HUMAN"
              Scacchiera.NeroTipoGiocatore = 1
            Case "PROGRAM"
              Scacchiera.NeroTipoGiocatore = 2
            Case "ROBOT"
              Scacchiera.NeroTipoGiocatore = 3
            Case "AUTOMA"
              Scacchiera.NeroTipoGiocatore = 4
            Case "CYBORG"
              Scacchiera.NeroTipoGiocatore = 5
            Case "DROID"
              Scacchiera.NeroTipoGiocatore = 6
            Case "MUTANT"
              Scacchiera.NeroTipoGiocatore = 7
            Case "CLONE"
              Scacchiera.NeroTipoGiocatore = 8
            Case "ALIEN"
              Scacchiera.NeroTipoGiocatore = 9
            Case "ANIMAL"
              Scacchiera.NeroTipoGiocatore = 10
            Case "UNKNOW ENTITY"
              Scacchiera.NeroTipoGiocatore = 11
            Case Else
              Scacchiera.NeroTipoGiocatore = 0
          End Select
        Case "[WHITECOUNTRY "
          Scacchiera.BiancoNazionalit‡ = s
        Case "[BLACKCOUNTRY "
          Scacchiera.NeroNazionalit‡ = s
        Case "[SECTION "
          Select Case UCase(s)
            Case "OPEN"
              Scacchiera.EventoSessione = 1
            Case "RESERVE"
              Scacchiera.EventoSessione = 2
            Case Else
              Scacchiera.EventoSessione = 0
          End Select
        Case "[STAGE "
          Select Case UCase(s)
            Case "PRELIMINARY"
              Scacchiera.EventoTurno = 5
            Case "SEMIFINAL"
              Scacchiera.EventoTurno = 2
            Case "FINAL"
              Scacchiera.EventoTurno = 1
            Case "1/4"
              Scacchiera.EventoTurno = 3
            Case "1/8"
              Scacchiera.EventoTurno = 4
            Case Else
              Scacchiera.EventoTurno = 0
          End Select
        Case "[BOARD "
          Scacchiera.EventoNumeroScacchiera = s
        Case "[OPENING "
          Scacchiera.EventoApertura = s
        Case "[VARIATION "
          Scacchiera.EventoSottoApertura = s
        Case "[SUBVARIATION "
          Scacchiera.EventoSottoSottoApertura = s
        Case "[ECO "
          Scacchiera.EventoECO = s
        Case "[NIC "
          Scacchiera.EventoNIC = s
        Case "[TERMINATION "
          Select Case UCase(s)
            Case "NORMAL"
              Scacchiera.EventoTerminato = 1
            Case "UNTERMINATED"
              Scacchiera.EventoTerminato = 2
            Case "ABANDONDED"
              Scacchiera.EventoTerminato = 3
            Case "ABJUDICATION"
              Scacchiera.EventoTerminato = 4
            Case "DEATH"
              Scacchiera.EventoTerminato = 7
            Case "EMERGENCY"
              Scacchiera.EventoTerminato = 6
            Case "RULES INFRACTION"
              Scacchiera.EventoTerminato = 4
            Case "TIMES FORFEIT"
              Scacchiera.EventoTerminato = 5
            Case Else
              Scacchiera.EventoTerminato = 0
          End Select
        Case "[ANNOTATOR "
          Scacchiera.EventoNomeAnnotatore = s
        Case "[MODE "
          Select Case UCase(s)
            Case "OTB"
              Scacchiera.EventoModo = 1
            Case "PM"
              Scacchiera.EventoModo = 2
            Case "EM"
              Scacchiera.EventoModo = 3
            Case "ICS"
              Scacchiera.EventoModo = 4
            Case "TC"
              Scacchiera.EventoModo = 5
            Case Else
              Scacchiera.EventoModo = 0
          End Select
        Case "[EVENTSPONSOR "
          Scacchiera.EventoSponsor = s
        Case "[TIME "
          On Local Error Resume Next
          Scacchiera.EvntoOra = DateValue(Sostituisci(Left(s, 2) + "." + Mid(s, 4, 2) + "." + Right(s, 2), "?", "0"))
          On Error GoTo 0
        Case "[FEN "
          FLPos = True
          Dim x As Colonna
          Dim Y As Riga
          Dim PO As Integer
          Dim DS As Variant
          Dim FLDS As Boolean
          Dim FLTURNO As Boolean
          Dim z As String
          FLTURNO = False
          FLDS = False
          Scacchiera.SvuotaScacchiera True
          DS = Scacchiera.DisposizioneScacchiera
          Y = 7
          x = 0
          For PO = 1 To Len(s)
            z = Mid(s, PO, 1)
            If FLDS = False Then
              Select Case z
                Case "p"
                  DS(x, Y) = Pedone_Nero
                  x = x + 1
                Case "n"
                  DS(x, Y) = Cavallo_Nero
                  x = x + 1
                Case "b"
                  DS(x, Y) = Alfiere_Nero
                  x = x + 1
                Case "r"
                  DS(x, Y) = Torre_Nero
                  x = x + 1
                Case "q"
                  DS(x, Y) = Dama_Nero
                  x = x + 1
                Case "k"
                  DS(x, Y) = Re_Nero
                  x = x + 1
                Case "P"
                  DS(x, Y) = Pedone_Bianco
                  x = x + 1
                Case "N"
                  DS(x, Y) = Cavallo_Bianco
                  x = x + 1
                Case "B"
                  DS(x, Y) = Alfiere_Bianco
                  x = x + 1
                Case "R"
                  DS(x, Y) = Torre_Bianco
                  x = x + 1
                Case "Q"
                  DS(x, Y) = Dama_Bianco
                  x = x + 1
                Case "K"
                  DS(x, Y) = Re_Bianco
                  x = x + 1
                Case "1"
                  DS(x, Y) = Nessun_Pezzo
                  x = x + 1
                Case "2"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo
                  x = x + 2
                Case "3"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo
                  x = x + 3
                Case "4"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo: DS(x + 3, Y) = Nessun_Pezzo
                  x = x + 4
                Case "5"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo: DS(x + 3, Y) = Nessun_Pezzo: DS(x + 4, Y) = Nessun_Pezzo
                  x = x + 5
                Case "6"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo: DS(x + 3, Y) = Nessun_Pezzo: DS(x + 4, Y) = Nessun_Pezzo: DS(x + 5, Y) = Nessun_Pezzo
                  x = x + 6
                Case "7"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo: DS(x + 3, Y) = Nessun_Pezzo: DS(x + 4, Y) = Nessun_Pezzo: DS(x + 5, Y) = Nessun_Pezzo: DS(x + 6, Y) = Nessun_Pezzo
                  x = x + 7
                Case "8"
                  DS(x, Y) = Nessun_Pezzo: DS(x + 1, Y) = Nessun_Pezzo: DS(x + 2, Y) = Nessun_Pezzo: DS(x + 3, Y) = Nessun_Pezzo: DS(x + 4, Y) = Nessun_Pezzo: DS(x + 5, Y) = Nessun_Pezzo: DS(x + 6, Y) = Nessun_Pezzo:: DS(x + 7, Y) = Nessun_Pezzo
                  x = x + 8
              End Select
              If x = 8 Then
                x = 0
                Y = Y - 1
                If Y = -1 Then
                  FLDS = True
                End If
              End If
            Else
              If FLTURNO = False Then
                Select Case z
                  Case "B", "b"
                    FLTURNO = True
                    Scacchiera.Editabilit‡ = NeroMossePossibili
                  Case "W", "w"
                    Scacchiera.Editabilit‡ = BiancoMossePossibili
                    FLTURNO = True
                End Select
              Else
                Select Case z
                  Case "K"
                    Scacchiera.BiancoArroccoPossibile = Scacchiera.BiancoArroccoPossibile Or Destra
                  Case "k"
                    Scacchiera.NeroArroccoPossibile = Scacchiera.NeroArroccoPossibile Or Destra
                  Case "Q"
                    Scacchiera.BiancoArroccoPossibile = Scacchiera.BiancoArroccoPossibile Or Sinistra
                  Case "q"
                    Scacchiera.NeroArroccoPossibile = Scacchiera.NeroArroccoPossibile Or Sinistra
                  Case "a"
                    xep = 0
                  Case "b"
                    xep = 1
                  Case "c"
                    xep = 2
                  Case "d"
                    xep = 3
                  Case "e"
                    xep = 4
                  Case "f"
                    xep = 5
                  Case "g"
                    xep = 6
                  Case "h"
                    xep = 7
                End Select
              End If
            End If
          Next
        Scacchiera.DisposizioneScacchiera = DS
      End Select
    Else
      If FLPos = False Then
        Scacchiera.PosizioneIniziale True
        Scacchiera.Editabilit‡ = BiancoMossePossibili
        FLPos = True
      End If
      'decodifica mosse
      Dim Livello As Integer: Rem Livello viene utilizzata per contare quante parentesi sono state aperte, esempio "(1(2(3(4COMMENTO))))" in questo caso il livello va da 1 a 4
      If L <> "" Then
        Dim C As String * 1
        s = ""
        Dim a As Byte
        Dim oa As Byte
        Dim fla As Byte
        For P = 1 To Len(L)
          C = Mid(L, P, 1)
          oa = a
          a = Asc(C)
          s = s + C
          'carattere ".."
          If a = 46 And oa = 46 Then
            If fla = 1 And Livello = 0 Then
              AggiungiCommento CStr(s), Scacchiera
              s = ""
              fla = 0
            Else
              s = ""
            End If
          'carattere "."
          ElseIf a = 46 And fla = 0 Then
            s = ""
          'carattere "{" e "("
          ElseIf a = 123 Or a = 40 Then
            If (Len(s) > 1) And (fla = 0) Then
              Scacchiera.EffettuaMossa ConvertiMossa(CStr(s), Scacchiera)
            End If
            s = C
            fla = 1
            Livello = Livello + 1
          'carattere "}" e ")"
          ElseIf a = 125 Or a = 41 Then
            'aggiungi commento
            Livello = Livello - 1
            If Livello = 0 Then
              If fla = 1 Then
                AggiungiCommento CStr(s), Scacchiera
                s = ""
                fla = 0
              End If
            End If
          'ATTENZIONE: Attivando questo si hanno errori nel caricare i commenti per esempio il file: KAMSKY.PGN Record: 392
          'nel caso si stiino aggiungendo pi˘ righe di commento
          'questa condizione verifica la fine delle aggiunte delle line di commento multiple
          'ElseIf a = 32 And oa = 46 And (P = 3 Or P = 4) And fla <> 0 Then
          '  s = ""
          '  fla = 0
          'carattere ";"
          ElseIf a = 59 Then
            fla = 1
            s = ""
          'carattere "$"
          ElseIf a = 36 Then
            fla = 2: Rem Quando FLA Ë uguale a due verr‡ registrato un commento che terminer‡ con lo spazio
            s = ""
          'Fine riga
          ElseIf (P = Len(L)) And fla = 1 Then
            AggiungiCommento CStr(s), Scacchiera
            s = ""
          ElseIf (a = 32 Or P = Len(L)) And fla <> 1 Then
            If s = "1-0" Or s = "0-1" Or s = "1/2-1/2" Then
              If Scacchiera.EventoRisultato = 0 Then
                Select Case s
                   Case "1-0"
                     Scacchiera.EventoRisultato = 1
                   Case "0-1"
                     Scacchiera.EventoRisultato = 2
                   Case "1/2-1/2"
                     Scacchiera.EventoRisultato = 3
                   Case Else
                     Scacchiera.EventoRisultato = 0
                 End Select
              End If
            'Viene registrato un commento che termina con il carattere di spazio, questi commenti hanno FLA impostato a 2
            ElseIf fla = 2 Then
                If Livello = 0 Then
                  AggiungiCommento CStr(s), Scacchiera
                  fla = 0
                Else
                  fla = 1
                End If
            Else
                Scacchiera.EffettuaMossa ConvertiMossa(CStr(s), Scacchiera, DisabilitaMSG)
            End If
            s = ""
          End If
        Next
      End If
    End If
  Wend
fine:
  Close #NumeroFile
  Scacchiera.XEnPassant = xep
  Scacchiera.AbilitaVisualizzazionePezzi = OLDScaAVP
End Sub
Function ConvertiMossa(s As String, Sca As Variant, Optional DisabilitaMSG As Boolean = False) As Spostamento
  Dim N As Integer
  Dim C As String * 1
  Dim Apc As Byte
  Dim pezzo As pezzo
  Dim PartenzaX As Colonna
  Dim PartenzaY As Riga
  Dim DestinazioneX As Colonna
  Dim DestinazioneY As Riga
  Dim FLPromozione As Boolean
  Dim CasaPartenza As CoordinateCasa
  Dim CasaDestinazione As CoordinateCasa
  PartenzaX = -1
  PartenzaY = -1
  DestinazioneX = -1
  DestinazioneY = -1
  If Sca.Editabilit‡ = BiancoMossePossibili Then
    pezzo = Pedone_Bianco
  Else
    pezzo = Pedone_Nero
  End If
  If s <> "" Then
    For N = 1 To Len(s)
      C = Mid(s, N, 1)
      If N < Len(s) Then
        Dim N2 As Integer
        For N2 = N + 1 To Len(s)
          Apc = Asc(Mid(s, N2, 1))
          If (Apc >= 97 And Apc <= 104) Or (Apc >= 49 And Apc <= 56) Then
            Exit For
          Else
            Apc = 0
          End If
        Next
      Else
        Apc = 0
      End If
      Select Case C
        Case "B"
          'Alfiere
          If FLPromozione = True Then
            ConvertiMossa.Provenienza = CasaPartenza
            ConvertiMossa.Destinazione = CasaDestinazione
            ConvertiMossa.Promosso = Alfiere
          Else
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              pezzo = Alfiere_Bianco
            Else
              pezzo = Alfiere_Nero
            End If
          End If
        Case "N"
          'Cavallo
          If FLPromozione = True Then
            ConvertiMossa.Provenienza = CasaPartenza
            ConvertiMossa.Destinazione = CasaDestinazione
            ConvertiMossa.Promosso = Cavallo
          Else
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              pezzo = Cavallo_Bianco
            Else
              pezzo = Cavallo_Nero
            End If
          End If
        Case "R"
          'Torre
          If FLPromozione = True Then
            ConvertiMossa.Provenienza = CasaPartenza
            ConvertiMossa.Destinazione = CasaDestinazione
            ConvertiMossa.Promosso = Torre
          Else
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              pezzo = Torre_Bianco
            Else
              pezzo = Torre_Nero
            End If
          End If
        Case "Q"
          'Dama
          If FLPromozione = True Then
            ConvertiMossa.Provenienza = CasaPartenza
            ConvertiMossa.Destinazione = CasaDestinazione
            ConvertiMossa.Promosso = Dama
          Else
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              pezzo = Dama_Bianco
            Else
              pezzo = Dama_Nero
            End If
          End If
        Case "K"
          'Re
          If Sca.Editabilit‡ = BiancoMossePossibili Then
            pezzo = Re_Bianco
          Else
            pezzo = Re_Nero
          End If
        Case "O"
          'arrocco
          If Mid(s, N, 5) = "O-O-O" Then
            'Arrocco Lungo
            N = N + 4
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              ConvertiMossa.Provenienza.x = Colonna_E
              ConvertiMossa.Provenienza.Y = Riga_1
              ConvertiMossa.Destinazione.x = Colonna_C
              ConvertiMossa.Destinazione.Y = Riga_1
            Else
              ConvertiMossa.Provenienza.x = Colonna_E
              ConvertiMossa.Provenienza.Y = Riga_8
              ConvertiMossa.Destinazione.x = Colonna_C
              ConvertiMossa.Destinazione.Y = Riga_8
            End If
          Else
            'Arrocco Corto
            If Sca.Editabilit‡ = BiancoMossePossibili Then
              ConvertiMossa.Provenienza.x = Colonna_E
              ConvertiMossa.Provenienza.Y = Riga_1
              ConvertiMossa.Destinazione.x = Colonna_G
              ConvertiMossa.Destinazione.Y = Riga_1
            Else
              ConvertiMossa.Provenienza.x = Colonna_E
              ConvertiMossa.Provenienza.Y = Riga_8
              ConvertiMossa.Destinazione.x = Colonna_G
              ConvertiMossa.Destinazione.Y = Riga_8
            End If
            N = N + 2
          End If
        Case "a"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 0
          Else
            DestinazioneX = 0
          End If
        Case "b"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 1
          Else
            DestinazioneX = 1
          End If
        Case "c"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 2
          Else
            DestinazioneX = 2
          End If
        Case "d"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 3
          Else
            DestinazioneX = 3
          End If
        Case "e"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 4
          Else
            DestinazioneX = 4
          End If
        Case "f"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 5
          Else
            DestinazioneX = 5
          End If
        Case "g"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 6
          Else
            DestinazioneX = 6
          End If
        Case "h"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaX = 7
          Else
            DestinazioneX = 7
          End If
        Case "1"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 0
          Else
            DestinazioneY = 0
          End If
        Case "2"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 1
          Else
            DestinazioneY = 1
          End If
        Case "3"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 2
          Else
            DestinazioneY = 2
          End If
        Case "4"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 3
          Else
            DestinazioneY = 3
          End If
        Case "5"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 4
          Else
            DestinazioneY = 4
          End If
        Case "6"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 5
          Else
            DestinazioneY = 5
          End If
        Case "7"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 6
          Else
            DestinazioneY = 6
          End If
        Case "8"
          If Apc >= 97 And Apc <= 104 Then
            PartenzaY = 7
          Else
            DestinazioneY = 7
          End If
      End Select
      If DestinazioneY <> -1 Then
        Dim q As Integer
        Dim Turno As Colore
        Dim Quantit‡Pezzi As Integer
        Dim Mosse() As Variant
        Quantit‡Pezzi = 0
        If Sca.Editabilit‡ = BiancoMossePossibili Then
          Turno = Bianco
        Else
          Turno = Nero
        End If
        Dim Y As Riga
        Dim x As Colonna
        Dim flOK As Boolean
        Dim DisposizioneScacchiera As Variant
        DisposizioneScacchiera = Sca.DisposizioneScacchiera
        flOK = False
        For Y = 0 To 7
          If PartenzaY = -1 Or PartenzaY = Y Then
            For x = 0 To 7
              If PartenzaX = -1 Or PartenzaX = x Then
                If DisposizioneScacchiera(x, Y) = pezzo Then
                  q = calcolo.MossePossibili(x, Y, Sca.DisposizioneScacchiera, Mosse, Sca.XEnPassant, Sca.BiancoArroccoPossibile, Sca.NeroArroccoPossibile, Sca.PosizioneReBianco, Sca.PosizioneReNero)
                  Dim NPZ As Integer
                  If q <> 0 Then
                    For NPZ = 1 To q
                      If Mosse(NPZ).x = DestinazioneX And Mosse(NPZ).Y = DestinazioneY Then
                        PartenzaX = x
                        PartenzaY = Y
                        flOK = True
                      End If
                    Next
                  End If
                End If
              End If
            Next
          End If
        Next
        'esegui la mossa
        If flOK = False Then
          If DisabilitaMSG = False Then
            MsgBox Msg1 + " (" + s + ")", vbInformation
          Else
            DisabilitaMSG = 2
          End If
        Else
          If Abs(pezzo) = Pedone And ((DestinazioneY = Riga_1) Or (DestinazioneY = Riga_8)) Then
            'Il pedone va a Promozione
            CasaPartenza.x = PartenzaX
            CasaPartenza.Y = PartenzaY
            CasaDestinazione.x = DestinazioneX
            CasaDestinazione.Y = DestinazioneY
            FLPromozione = True
          Else
            ConvertiMossa.Provenienza.x = PartenzaX
            ConvertiMossa.Provenienza.Y = PartenzaY
            ConvertiMossa.Destinazione.x = DestinazioneX
            ConvertiMossa.Destinazione.Y = DestinazioneY
          End If
        End If
        PartenzaX = -1
        PartenzaY = -1
        DestinazioneX = -1
        DestinazioneY = -1
      End If
    Next
  End If
End Function
Public Sub CopiaPartita(NomeFile As String, SeekRecord As Long, FileDestinazione As String)
  Dim NF As Integer
  Dim NF2 As Integer
  NF2 = FreeFile
  Open NomeFile For Input As NF2
  NF = FreeFile
  Open FileDestinazione For Append As NF
      Seek NF2, SeekRecord
      Dim L As String
      Dim FL As Integer
      Do While EOF(NF2) = False
        Line Input #NF2, L
        If UCase(Left(L, 7)) = "[EVENT " Then
          FL = FL + 1
        End If
        If FL = 2 Then
          Exit Do
        End If
        Print #NF, L
      Loop
  Close NF
  Close NF2
End Sub
Public Function Partita(File As Variant, Optional Posizione As Long, Optional PosizioneSuccessivo As Long = -1) As String
'Questa funzione restituisce una stringa che descrive il record PGN specificato dalla Posizione
'il parametro Optional Posizione se a 0 imposta la lettura del record dalla posizione attuale di Seek(IndexFile)
'mentre se Ë diverso da zero imposta la lettura dalla Posizione Seek specificata, inoltre Posizione restituisce
'la posizione del recod letto
'PosizioneSuccessivo restituisce la posizione del pecord seguente a quello letto, se non vi sono record successivi
'assumer‡ il valore 0
'EndFile restituisce True se si Ë letto l'ultimo record
'FILE (as Variant) puÚ assumere il nome del File di archivio PGN oppure il numero del file in caso sia gi‡ aperto
    Dim NumeroFile As Integer
    Dim FLType As Boolean
    If TypeName(File) = "String" Then
      NumeroFile = FreeFile
      Open File For Input As #NumeroFile
      FLType = True
    Else
      NumeroFile = File
    End If
    Dim NonRilevaSuccessivo As Boolean
    If PosizioneSuccessivo = -1 Then
      NonRilevaSuccessivo = True
    End If
    Dim SK As Long
    Dim L As String
    If Posizione <> 0 Then
      Seek #NumeroFile, Posizione
    End If
    While EOF(NumeroFile) = False
      SK = Seek(NumeroFile)
      Line Input #NumeroFile, L
      Dim FL As Boolean
      Dim s As String
      Dim P As Integer
      Dim P2 As Integer
      Dim D As String
      P = InStr(L, Chr(34))
      If P <> 0 Then
        D = UCase(Left(L, P - 1))
        P2 = InStr(P + 1, L, Chr(34))
        If P2 <> 0 Then
          s = Mid(L, P + 1, P2 - P - 1)
        End If
        Dim EV As String
        Dim SI As String
        Dim DA As String
        Dim RO As String
        Dim WH As String
        Dim RE As String
        Dim BL As String
        Dim EC As String
        Select Case D
          Case "[EVENT "
            If FL = False Then
              FL = True
              Posizione = SK
              EV = s
            Else
              PosizioneSuccessivo = SK
              GoTo FineRecord
            End If
          Case "[SITE "
            SI = s
          Case "[DATE "
            DA = Right(s, 2) + "\" + Mid(s, 6, 2) + "\" + Left(s, 4)
          Case "[ROUND "
            RO = s
          Case "[WHITE "
            WH = s
          Case "[BLACK "
            BL = s
          Case "[RESULT "
            RE = s
          Case "[ECO "
            EC = s
          Case Else
            If NonRilevaSuccessivo Then
              GoTo FineRecord
            End If
        End Select
      Else
        If NonRilevaSuccessivo Then
          GoTo FineRecord
        End If
      End If
  Wend
PosizioneSuccessivo = 0
FineRecord:
    Partita = EV + Chr(9) + SI + Chr$(9) + DA + Chr(9) + RO + Chr(9) + WH + Chr(9) + BL + Chr(9) + RE + Chr(9) + EC
  If FLType Then
    Close NumeroFile
  End If
End Function
Private Function Sostituisci(Stringa As String, carattere As String, sostituto As String)
  Dim NS As String
  Dim N As Integer
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
End Function

Sub AggiungiCommento(s As String, Sca As Variant)
  s = LTrim(s)
  If Sca.Quantit‡MosseTraccia = 0 Then
    If Sca.Evento <> "" Then
      Sca.Evento = Sca.Evento + Chr(13) + Chr(10)
    End If
    Sca.Evento = Sca.Evento + s
  Else
    Dim Traccia As Variant
    Traccia = Sca.Traccia
    If Traccia(Sca.Quantit‡MosseTraccia).Nota <> "" Then
      Traccia(Sca.Quantit‡MosseTraccia).Nota = Traccia(Sca.Quantit‡MosseTraccia).Nota + Chr$(13) + Chr$(10)
    End If
    Traccia(Sca.Quantit‡MosseTraccia).Nota = Traccia(Sca.Quantit‡MosseTraccia).Nota + s
    Sca.Traccia = Traccia
  End If
End Sub
Private Sub Class_Terminate()
  Set SV = Nothing
End Sub

