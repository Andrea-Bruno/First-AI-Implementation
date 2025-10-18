VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PezziScacchiera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
DefInt A-Z
Public Enum Colore
    Nessun_Colore = 0
    Bianco = 1
    Nero = -1
End Enum
Public Enum pezzo
    Nessun_Pezzo = 0
    Pedone_Bianco = 1
    Cavallo_Bianco = 2
    Alfiere_Bianco = 4
    Torre_Bianco = 8
    Dama_Bianco = 12
    Re_Bianco = 16
    Pedone_Nero = -1
    Cavallo_Nero = -2
    Alfiere_Nero = -4
    Torre_Nero = -8
    Dama_Nero = -12
    Re_Nero = -16
End Enum
Public Enum Genere
  Nessun_Genere = 0
  Pedone = 1
  Cavallo = 2
  Alfiere = 4
  Torre = 8
  Dama = 12
  RE = 16
End Enum
Public Enum TipoDiSelezione
  Nessuna = 0
  Bordo = 1
  Evidenziato = 2
  EvidenziatoEBordo = 3
End Enum
Public Function ColorePezzo(pz As pezzo) As Colore
Attribute ColorePezzo.VB_Description = "Restituisce il colore di un pezzo del gioco degli scacchi"
  If pz > 0 Then
    ColorePezzo = Bianco
  ElseIf pz < 0 Then
    ColorePezzo = Nero
  Else
    ColorePezzo = Nessun_Colore
  End If
End Function
Public Function GenerePezzo(pz As pezzo) As Genere
  GenerePezzo = Abs(pz)
End Function
