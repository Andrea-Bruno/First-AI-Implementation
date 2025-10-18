VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClasseMosse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
DefInt A-Z
Option Explicit
Enum ErroreMossa
  ValoreVuoto = 0
  BloccoOrizzontale = 1
  BloccoVerticale = 2
  BloccoDiagonale45g = 4
  BloccoDiagonale90g = 8
  NessunErrore = 16
End Enum
Type Mosse
  pezzo As pezzo
  Posizione As CoordinateCasa
  Quantit‡Mosse As Integer
  Mosse As Variant
End Type
