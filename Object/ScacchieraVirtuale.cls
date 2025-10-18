VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScacchieraVirtuale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
DefInt A-Z
Const InizioAsseX = 0
Const FineAsseX = 7
Const InizioAsseY = 0
Const FineAsseY = 7
Public Enum Colonna
    NessunaColonna = InizioAsseX - 1
    Colonna_A = InizioAsseX
    Colonna_B = Colonna_A + 1
    Colonna_C = Colonna_B + 1
    Colonna_D = Colonna_C + 1
    Colonna_E = Colonna_D + 1
    Colonna_F = Colonna_E + 1
    Colonna_G = Colonna_F + 1
    Colonna_H = Colonna_G + 1
End Enum
Public Enum Riga
    NessunaRiga = InizioAsseY - 1
    Riga_1 = InizioAsseY
    Riga_2 = Riga_1 + 1
    Riga_3 = Riga_2 + 1
    Riga_4 = Riga_3 + 1
    Riga_5 = Riga_4 + 1
    Riga_6 = Riga_5 + 1
    Riga_7 = Riga_6 + 1
    Riga_8 = Riga_7 + 1
End Enum
Public Enum IndexCasa
    NessunaCasa = -1
    A1 = 0
    B1 = 1
    C1 = 2
    D1 = 3
    E1 = 4
    F1 = 5
    G1 = 6
    H1 = 7
    A2 = 8
    B2 = 9
    C2 = 10
    D2 = 11
    E2 = 12
    F2 = 13
    G2 = 14
    H2 = 15
    A3 = 16
    B3 = 17
    C3 = 18
    D3 = 19
    E3 = 20
    F3 = 21
    G3 = 22
    H3 = 23
    A4 = 24
    B4 = 25
    C4 = 26
    D4 = 27
    E4 = 28
    F4 = 29
    G4 = 30
    H4 = 31
    A5 = 32
    B5 = 33
    C5 = 34
    D5 = 35
    E5 = 36
    F5 = 37
    G5 = 38
    H5 = 39
    A6 = 40
    B6 = 41
    C6 = 42
    D6 = 43
    E6 = 44
    F6 = 45
    G6 = 46
    H6 = 47
    A7 = 48
    B7 = 49
    C7 = 50
    D7 = 51
    E7 = 52
    F7 = 53
    G7 = 54
    H7 = 55
    A8 = 56
    B8 = 57
    C8 = 58
    D8 = 59
    E8 = 60
    F8 = 61
    G8 = 62
    H8 = 63
End Enum
Public Enum Lato
  Nessun_Lato = 0
  Sinistra = 1
  Destra = 2
  Entrambi = 3
End Enum
Public Enum EditorScacchiera
  NonEditabile = 0
  Editabile = 2
  BiancoMossePossibili = 1
  NeroMossePossibili = -1
End Enum
Public Enum OrientamentoScacchiera
    BiancoNero = 0
    NeroBianco = 1
    SinistraDestra = 2
    DestraSinistra = 3
End Enum
Public Enum TitoloFide
    NessuTitolo = 0
    GM = 1
    IM = 2
    FM = 3
End Enum
Public Enum TurnoDiGioco
  NessunTurno = 0
  Finale = 1
  Semifinale = 2
  QuartiDiFinale = 3
  OttaviDiFinale = 4
  Qualificazione = 5
End Enum
Public Enum TipoGiocatore
  TipoNonSpecificato = 0
  Umano = 1
  Computer = 2
  Robot = 3
  Automa = 4
  Cibernetico = 5
  Droide = 6
  Mutante = 7
  Clone = 8
  Alieno = 9
  Animale = 10
  Entit‡Sconosciuta = 11
End Enum
Public Enum Sessione
  NessunaSessione = 0
  Aperta = 1
  Riservata = 2
End Enum
Public Enum TipoGara
  TipoGaraNonSpecificato = 0
  Normale = 1
  Lampo = 2
  Semilampo = 3
End Enum
Public Enum TermineEvento
  TermineNonSpecificato = 0
  Normalmente = 1
  NonFinito = 2
  Abbandono = 3
  Aggiudicato = 4
  FineTempo = 5
  Emmergenza = 6
  Concordato = 7
End Enum
Public Enum ModoDiGioco
  ModoNonSpecificato = 0
  ConScacchiera = 1
  TramiteLettera = 2
  TramiteEMAIL = 3
  TramiteInternetServerDiScacchi = 4
  TramiteMezzoDiComunicazione = 5
  AllaCieca = 6
  ScacchieraNonStandard = 7
  Ologramma = 8
End Enum
Public Enum Risultato
  RisultatoNonSpecificato = 0
  VittoriaBianco = 1
  VittoriaNero = 2
  Pareggio = 3
End Enum
'Tipi definiti dall'utente
Type ImpostazioneScacchiera
    pezzo(InizioAsseX To FineAsseX, InizioAsseY To FineAsseY) As pezzo
End Type
Public Type CoordinateCasa
    x As Colonna
    Y As Riga
End Type
Type Spostamento
  Provenienza As CoordinateCasa
  Destinazione As CoordinateCasa
  Promosso As Genere
End Type
Type CollocazionePezzo
  pezzo As pezzo
  Posizione As CoordinateCasa
End Type
Type DisposizionePezzi
    pezzo(Nero To Bianco, 1 To 64) As CollocazionePezzo
End Type
Type Mossa
  Provenienza As CollocazionePezzo
  Destinazione As CollocazionePezzo
  Promozione As Genere
  NeroArroccoPossibile As Lato
  BiancoArroccoPossibile As Lato
  XEnPassant As Colonna
  Editabilit‡ As EditorScacchiera
  Nota As String
  Collegamento As Variant
  Tempo As Date
End Type
Public Function IndividuaXCasa(ByVal Index As IndexCasa) As Colonna
  If Index = NessunaCasa Then
    IndividuaXCasa = NessunaColonna
  Else
    IndividuaXCasa = (Index Mod 8)
  End If
End Function
Public Function IndividuaYCasa(ByVal Index As IndexCasa) As Riga
  If Index = NessunaCasa Then
    IndividuaYCasa = NessunaRiga
  Else
    IndividuaYCasa = (Index - Index Mod 8) / 8
  End If
End Function
Public Function IndividuaCoordinateCasa(ByVal Index As IndexCasa) As CoordinateCasa
  If Index = NessunaCasa Then
    IndividuaCoordinateCasa.x = NessunaColonna
    IndividuaCoordinateCasa.Y = NessunaRiga
  Else
    IndividuaCoordinateCasa.x = (Index Mod 8)
    IndividuaCoordinateCasa.Y = (Index - IndividuaCoordinateCasa.x) / 8
  End If
End Function
Public Function IndividuaIndexCasa(ByVal x As Colonna, ByVal Y As Riga) As IndexCasa
  If x = NessunaColonna Or Y = NessunaRiga Then
    IndividuaIndexCasa = NessunaCasa
  Else
    IndividuaIndexCasa = x + Y * 8
  End If
End Function
Public Function AnnotazioneAlgebrica(x As Colonna, Y As Riga) As String
  AnnotazioneAlgebrica = LTrim(Chr(x + 65)) + LTrim(Str(Y))
End Function

