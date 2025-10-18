VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodiceSituazione"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
DefInt A-X
Option Explicit
Const Msg1 = "Operazione errata: I/Il Re possono/puÚ ancora arroccare, quindi non si possono fare operazioni speculari"
Public Enum Analogia
    ColoriInvertiti = 1
    Speculare = 2
    InvertitoSpeculare = 3
    NessunaAnalogia = 0
End Enum
Public Function Situazione(DisposizioineScacchiera As Variant, ColoreCheMuove As EditorScacchiera, BiancoArroccoPossibile As Lato, NeroArroccoPossibile As Lato, XEnPassant As Colonna, Optional RuotatoSpeculare As Analogia) As String
  'restituisce una stringa di 8 + LenPezzi("********") + 1 byte
  Situazione = Posizione(DisposizioineScacchiera, ColoreCheMuove, BiancoArroccoPossibile, NeroArroccoPossibile, RuotatoSpeculare)
  Situazione = Situazione + Pezzi(DisposizioineScacchiera, RuotatoSpeculare)
  Situazione = Situazione + ArroccoXEnPassant(BiancoArroccoPossibile, NeroArroccoPossibile, XEnPassant, RuotatoSpeculare)
End Function
Public Function ArroccoXEnPassant(ByVal BiancoArroccoPossibile As Lato, ByVal NeroArroccoPossibile As Lato, ByVal XEnPassant As Colonna, Optional RuotatoSpeculare As Analogia) As String
Attribute ArroccoXEnPassant.VB_Description = "Restituisce un byte che raffigura: bit 0,1,2,3 = X En Passant (0 = posizione A, 7 = posizione H, 8 = en passant inesistente) ; bit 4 e 5 = Arrocco bianco; 6 e 7 = Arrocco nero."
  'restituisce un stringa di lunghezza fissa 1
  If RuotatoSpeculare And Speculare Then
    If (BiancoArroccoPossibile <> Nessun_Lato) Or (NeroArroccoPossibile <> Nessun_Lato) Then
      MsgBox Msg1
    End If
    If XEnPassant = NessunaColonna Then
      XEnPassant = 7 - XEnPassant
    End If
  End If
  If RuotatoSpeculare And ColoriInvertiti Then
    Dim tmp As Lato
    tmp = BiancoArroccoPossibile
    BiancoArroccoPossibile = NeroArroccoPossibile
    NeroArroccoPossibile = tmp
  End If
  Dim B As Byte
  If XEnPassant = NessunaColonna Then
    B = 8
  Else
    B = XEnPassant
  End If
  If BiancoArroccoPossibile And 1 Then
    B = B Or 16
  End If
  If BiancoArroccoPossibile And 2 Then
    B = B Or 32
  End If
  If NeroArroccoPossibile And 1 Then
    B = B Or 64
  End If
  If NeroArroccoPossibile And 2 Then
    B = B Or 128
  End If
  ArroccoXEnPassant = Chr(B)
End Function
Public Function Pezzi(DisposizioineScacchiera As Variant, Optional RuotatoSpeculare As Analogia) As String
Attribute Pezzi.VB_Description = "Stringa che indica i pezzi nella scacchiera, ogni pezzo occupa 4 bit (1 byte = 2 pezzi)."
  'restituisce una stringa di lunghezza variabile calcolabile con:
  'L = Int((1+Quantit‡Pezzi(Posizione$))/2)      - Posizione$ Ë sempre una stringa di 8 byte restituita dalla funzione Posizione
  Dim x As Colonna
  Dim Y As Riga
  Dim B As Byte
  Dim q As Byte
  Dim figura As pezzo
  Dim N As Byte
  For Y = Riga_1 To Riga_8
    For x = Colonna_A To Colonna_H
      figura = DS(x, Y, DisposizioineScacchiera, RuotatoSpeculare)
      If figura <> Nessun_Pezzo Then
        If figura = Pedone_Bianco Then
          N = 1
        ElseIf figura = Cavallo_Bianco Then
          N = 2
        ElseIf figura = Alfiere_Bianco Then
          N = 3
        ElseIf figura = Torre_Bianco Then
          N = 4
        ElseIf figura = Dama_Bianco Then
          N = 5
        ElseIf figura = Re_Bianco Then
          N = 6
        ElseIf figura = Pedone_Nero Then
          N = 7
        ElseIf figura = Cavallo_Nero Then
          N = 8
        ElseIf figura = Alfiere_Nero Then
          N = 9
        ElseIf figura = Torre_Nero Then
          N = 10
        ElseIf figura = Dama_Nero Then
          N = 11
        ElseIf figura = Re_Nero Then
          N = 12
        End If
        q = q + 1
        'dispari
        If q And 1 Then
          B = N
        'pari
        Else
          B = B Or (N * 16)
          Pezzi = Pezzi + Chr(B)
          B = 0
        End If
      End If
    Next
  Next
  If q And 1 Then
    Pezzi = Pezzi + Chr(B)
  End If
End Function
Public Function Posizione(DisposizioneScacchiera As Variant, ColoreCheMuove As EditorScacchiera, Optional BiancoArroccoPossibile As Lato, Optional NeroArroccoPossibile As Lato, Optional RuotatoSpeculare As Analogia) As String
  'restuisce una stringa di lunghezza fissa 8 byte
  Dim x As Colonna
  Dim Y As Riga
  Dim B As Byte
  If ColoreCheMuove <> BiancoMossePossibili Then
    RuotatoSpeculare = RuotatoSpeculare Xor ColoriInvertiti
  End If
  If NeroArroccoPossibile = Nessun_Lato Then
    If BiancoArroccoPossibile = Nessun_Lato Then
      Dim XRe As Colonna
      XRe = NessunaColonna
      For Y = Riga_1 To Riga_8
        For x = Colonna_A To Colonna_H
          If DS(x, Y, DisposizioneScacchiera, RuotatoSpeculare) = Re_Bianco Then
            XRe = x
            If XRe <= Colonna_D Then
              RuotatoSpeculare = RuotatoSpeculare Xor Speculare
            End If
            Exit For
          End If
        Next
        If XRe <> NessunaColonna Then
          Exit For
        End If
      Next
    End If
  End If
  For Y = Riga_1 To Riga_8
    B = 0
    For x = Colonna_A To Colonna_H
      If DS(x, Y, DisposizioneScacchiera, RuotatoSpeculare) <> Nessun_Pezzo Then
        B = B Or (2 ^ x)
      End If
    Next
    Posizione = Posizione + Chr(B)
  Next
End Function
Public Function DS(ByVal x, ByVal Y, DisposizioneScacchiera As Variant, Optional RuotatoSpeculare As Analogia)
Attribute DS.VB_Description = "Individua un pezzo nella scacchiera sia normalmente che a colori invertiti o in posizione speculare."
  If RuotatoSpeculare And Speculare Then
    x = 7 - x
  End If
  If RuotatoSpeculare And ColoriInvertiti Then
    DS = -DisposizioneScacchiera(x, 7 - Y)
  Else
    DS = DisposizioneScacchiera(x, Y)
  End If
End Function
Public Function Quantit‡Pezzi(Stringa As String) As Byte
  If Len(Stringa) = 8 Then
    Dim N As Integer
    Dim B As Byte
    For N = 1 To 8
      B = Asc(Mid(Stringa, N, 1))
      Quantit‡Pezzi = Quantit‡Pezzi + QPezziRiga(B)
    Next
  Else
    MsgBox "Len <> 8"
  End If
End Function
Private Function QPezziRiga(R As Byte) As Byte
  'restituisce quanti pezzi vi sono sulla riga rappresentata dal byte R
  If R And 1 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 2 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 4 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 8 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 16 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 32 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 64 Then
    QPezziRiga = QPezziRiga + 1
  End If
  If R And 128 Then
    QPezziRiga = QPezziRiga + 1
  End If
End Function
Public Function LenPezzi(Posizione As String) As Integer
  LenPezzi = Int((1 + Quantit‡Pezzi(Posizione)) / 2)
End Function

