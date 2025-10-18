VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Nota 
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati inerenti l'incintro"
      Height          =   3975
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   5775
      Begin VB.TextBox TxtExtra 
         Height          =   285
         Left            =   3960
         TabIndex        =   19
         ToolTipText     =   "Campo libero per informazioni aggiuntive"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox CmbModo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Indica lo strumento che ha permesso di svolgere la partita"
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox TxtAnnotatore 
         Height          =   285
         Left            =   3960
         TabIndex        =   17
         ToolTipText     =   "Cognome e nome dell'annotatore"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.ComboBox CmbTerminato 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Indica in che modo ha avuto fine l'incontro"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox TxtNic 
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         ToolTipText     =   "Codice di riferimento al database delle aperture"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox TxtEco 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         ToolTipText     =   "Indica il codice di apertura riportato nella ""Encyclopedia of chess openings"""
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox TxtSottoSottoApertura 
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         ToolTipText     =   "Indica una variante della SottoApertura"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtSottoApertura 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Indica la sottoapertura adottata nella partita"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox TxtApertura 
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         ToolTipText     =   "Indica il tipo di apertura adottata in questa partita"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox CmbRisultato 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Indica quale giocatore ha ottenuto la vittoria o se vi è stato un pareggio"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TxtNumeroScacchiera 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         ToolTipText     =   "Numero nella scacchiera nel caso in cui si sta giocando in un torneo con scacchiere numerate"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox TxtIncontro 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "Indica il numero di partita nel caso in cui vi è una numerazione delle partite"
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox CmbTurno 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Turno di gara nel caso che si tratti di un torneo o campionato"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox CmbSessione 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Sessione dell'evento (in particolare descrive la possibilita di accedervi)"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TxtSponsor 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Finanziatore della gara"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.ComboBox CmbTipoGara 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Sessione dell'evento (in particolare descrive la possibilita di accedervi)"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtOra 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   "Ora in cui ha avuto inizio l'incontro"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtData 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Data in cui si è disputato l'incontro"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtLuogo 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Posto in cui si è svolta la partita"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Extra:"
         Height          =   255
         Index           =   18
         Left            =   3360
         TabIndex        =   54
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Modo:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   53
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Annotatore:"
         Height          =   255
         Index           =   16
         Left            =   2880
         TabIndex        =   52
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Conclusione:"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   51
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "NIC:"
         Height          =   255
         Index           =   14
         Left            =   3120
         TabIndex        =   50
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ECO:"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   49
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sottovariante:"
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   48
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Variante:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   47
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Apertura:"
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   46
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Risultato:"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   45
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero della scacchiera:"
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   44
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Incontro:"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   43
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Turno:"
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sessione:"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   41
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sponsor:"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   39
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ora:"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   38
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Luogo:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati inerenti il giocatore che ha il colore bianco"
      Height          =   3975
      Index           =   0
      Left            =   120
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtUscf 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   22
         ToolTipText     =   "Indice USCF"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   20
         ToolTipText     =   "Cognome e nome del giocatore"
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox TxtElo 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   21
         ToolTipText     =   "Punteggio ELO fide"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox CmbTitolo 
         Height          =   315
         Index           =   0
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Titolo in qualita di giocatore"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   24
         ToolTipText     =   "Iindirizzo E-MAIL o Network del giocatore"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Index           =   0
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Titolo in qualita di giocatore"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox TxtNazione 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   26
         ToolTipText     =   "Nazionalità del giocatore"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Elo:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Uscf:"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   68
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Titolo:"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   67
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Mail:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nazione:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati inerenti il giocatore che ha il colore nero"
      Height          =   3975
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtNome 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   27
         ToolTipText     =   "Cognome e nome del giocatore"
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox TxtElo 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   28
         ToolTipText     =   "Punteggio ELO fide"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtUscf 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   29
         ToolTipText     =   "Indice USCF"
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox CmbTitolo 
         Height          =   315
         Index           =   1
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Titolo in qualita di giocatore"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   31
         ToolTipText     =   "Iindirizzo E-MAIL o Network del giocatore"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Index           =   1
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Titolo in qualita di giocatore"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox TxtNazione 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   33
         ToolTipText     =   "Nazionalità del giocatore"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   62
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Elo:"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   61
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Uscf:"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   60
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Titolo:"
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   59
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Mail:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   58
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   57
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nazione:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   56
         Top             =   2040
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4455
      Left            =   0
      TabIndex        =   34
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Incontro"
            Object.ToolTipText     =   "Impostazione dei dati rigusrdanti l'incontro"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bianco"
            Object.ToolTipText     =   "Impostazione dei dati inerenti il giocatore bianco"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nero"
            Object.ToolTipText     =   "Impostazione dei dati inerenti il giocatore bianco"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NS = "Non specificato"
Const TipoGara1 = "Normale"
Const TipoGara2 = "Lampo"
Const TipoGara3 = "Semilampo"
Const Sessione1 = "Aperta"
Const Sessione2 = "Riservata"
Const Turno1 = "Finale"
Const Turno2 = "Semifinale"
Const Turno3 = "Quarti di finale"
Const Turno4 = "Ottavi di finale"
Const Turno5 = "Qualificazione"
Const Risultato1 = "Vittoria del bianco"
Const Risultato2 = "Vittoria del nero"
Const Risultato3 = "Pareggio"
Const Termine1 = "Normalmente"
Const Termine2 = "In sospeso"
Const Termine3 = "Abbandono"
Const Termine4 = "Aggiudicato"
Const Termine5 = "Tempo scaduto"
Const Termine6 = "Emergenza"
Const Termine7 = "Concordato"
Const Modo1 = "Con scacchiera standard"
Const Modo2 = "Tramite lettera"
Const Modo3 = "Tramite E-Mail"
Const Modo4 = "Tramite Internet server"
Const Modo5 = "Mezzo di comunicazione"
Const Modo6 = "Alla cieca"
Const Modo7 = "Scacchiera non standard"
Const Modo8 = "Ologramma"
Const Titolo1 = "FM"
Const Titolo2 = "IM"
Const Titolo3 = "GM"
Const Tipo1 = "Umano"
Const Tipo2 = "Computer"
Const Tipo3 = "Robot"
Const Tipo4 = "Automa"
Const Tipo5 = "Cibernetico"
Const Tipo6 = "Droide"
Const Tipo7 = "Mutante"
Const Tipo8 = "Clone"
Const Tipo9 = "Alieno"
Const Tipo10 = "Animale"
Const Tipo11 = "Entità Sconosciuta"
Dim AP As Variant
Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
End Sub
Private Sub Form_Load()
  'Centra il form
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  CmbTipoGara.AddItem NS
  CmbTipoGara.AddItem TipoGara1
  CmbTipoGara.AddItem TipoGara2
  CmbTipoGara.AddItem TipoGara3
  
  CmbSessione.AddItem NS
  CmbSessione.AddItem Sessione1
  CmbSessione.AddItem Sessione2

  CmbTurno.AddItem NS
  CmbTurno.AddItem Turno1
  CmbTurno.AddItem Turno2
  CmbTurno.AddItem Turno3
  CmbTurno.AddItem Turno4
  CmbTurno.AddItem Turno5
  
  CmbRisultato.AddItem NS
  CmbRisultato.AddItem Risultato1
  CmbRisultato.AddItem Risultato2
  CmbRisultato.AddItem Risultato3
  
  CmbTerminato.AddItem NS
  CmbTerminato.AddItem Termine1
  CmbTerminato.AddItem Termine2
  CmbTerminato.AddItem Termine3
  CmbTerminato.AddItem Termine4
  CmbTerminato.AddItem Termine5
  CmbTerminato.AddItem Termine6
  CmbTerminato.AddItem Termine7
  
  CmbModo.AddItem NS
  CmbModo.AddItem Modo1
  CmbModo.AddItem Modo2
  CmbModo.AddItem Modo3
  CmbModo.AddItem Modo4
  CmbModo.AddItem Modo5
  CmbModo.AddItem Modo6
  CmbModo.AddItem Modo7
  CmbModo.AddItem Modo8
  
  CmbTitolo(0).AddItem NS
  CmbTitolo(0).AddItem Titolo1
  CmbTitolo(0).AddItem Titolo2
  CmbTitolo(0).AddItem Titolo3
  
  CmbTitolo(1).AddItem NS
  CmbTitolo(1).AddItem Titolo1
  CmbTitolo(1).AddItem Titolo2
  CmbTitolo(1).AddItem Titolo3
  
  CmbTipo(0).AddItem NS
  CmbTipo(0).AddItem Tipo1
  CmbTipo(0).AddItem Tipo2
  CmbTipo(0).AddItem Tipo3
  CmbTipo(0).AddItem Tipo4
  CmbTipo(0).AddItem Tipo5
  CmbTipo(0).AddItem Tipo6
  CmbTipo(0).AddItem Tipo7
  CmbTipo(0).AddItem Tipo8
  CmbTipo(0).AddItem Tipo9
  CmbTipo(0).AddItem Tipo10
  CmbTipo(0).AddItem Tipo11
  
  CmbTipo(1).AddItem NS
  CmbTipo(1).AddItem Tipo1
  CmbTipo(1).AddItem Tipo2
  CmbTipo(1).AddItem Tipo3
  CmbTipo(1).AddItem Tipo4
  CmbTipo(1).AddItem Tipo5
  CmbTipo(1).AddItem Tipo6
  CmbTipo(1).AddItem Tipo7
  CmbTipo(1).AddItem Tipo8
  CmbTipo(1).AddItem Tipo9
  CmbTipo(1).AddItem Tipo10
  CmbTipo(1).AddItem Tipo11
End Sub

Private Sub CmbModo_Click()
  AP.ImpostaModo CmbModo.ListIndex
End Sub

Private Sub CmbRisultato_Click()
  AP.ImpostaRisultato CmbRisultato.ListIndex
End Sub

Private Sub CmbSessione_Click()
  AP.ImpostaSessione CmbSessione.ListIndex
End Sub
Private Sub CmbTipoGara_Click()
  AP.ImpostaTipoGara CmbTipoGara.ListIndex
End Sub

Private Sub CmbTerminato_Click()
  AP.ImpostaTerminato CmbTerminato.ListIndex
End Sub

Private Sub CmbTipo_Click(Index As Integer)
  AP.ImpostaTipo CmbTipo(Index).ListIndex, Index
End Sub

Private Sub CmbTitolo_Click(Index As Integer)
  AP.ImpostaTitolo CmbTitolo(Index).ListIndex, Index
End Sub

Private Sub CmbTurno_Click()
  AP.ImpostaTurno CmbTurno.ListIndex
End Sub

Private Sub Form_Resize()
  Dim h As Integer
  Dim base As Integer
  base = tbsOptions.Height
  'base = 5565
  h = ScaleHeight - base
  If h < 1 Then
    h = 1
  End If
  Text.Move 0, base, ScaleWidth, h
End Sub

Private Sub Form_Terminate()
  Set AP = Nothing
End Sub


Private Sub Text_Change()
  AP.ImpostaEvento Text.Text
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'Gestisce la combinazione di tasti CTRL+TAB per lo spostamento alla scheda successiva
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'È l'ultima scheda ed è pertanto necessario tornare alla scheda 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'Incrementa il conteggio delle schede
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub
Private Sub tbsOptions_Click()
  'Visualizza e attiva i controlli della scheda selezionata nascondendo e disattivando tutti gli altri
  Select Case tbsOptions.SelectedItem.Index
    Case 1
      Frame1.Visible = True
      Frame2(0).Visible = False
      Frame2(1).Visible = False
    Case 2
      Frame1.Visible = False
      Frame2(0).Visible = True
      Frame2(1).Visible = False
    Case 3
      Frame1.Visible = False
      Frame2(0).Visible = False
      Frame2(1).Visible = True
  End Select
End Sub

Private Sub TxtAnnotatore_Change()
  AP.ImpostaAnnotatore TxtAnnotatore.Text
End Sub

Private Sub TxtApertura_Change()
  AP.ImpostaApertura TxtApertura.Text
End Sub

Private Sub TxtData_Change()
  On Local Error Resume Next
  AP.ImpostaData DateValue(TxtData.Text)
End Sub

Private Sub TxtEco_Change()
  AP.ImpostaEco TxtEco.Text
End Sub

Private Sub TxtElo_Change(Index As Integer)
  On Local Error Resume Next
  AP.ImpostaElo TxtElo(Index).Text, Index
End Sub

Private Sub TxtEmail_Change(Index As Integer)
  AP.ImpostaEMail TxtEmail(Index).Text, Index
End Sub
Private Sub TxtExtra_Change()
  AP.ImpostaExtra TxtExtra.Text
End Sub
Private Sub TxtIncontro_Change()
  On Local Error Resume Next
  AP.ImpostaIncontro TxtIncontro.Text
End Sub

Private Sub TxtLuogo_Change()
  AP.ImpostaLuogo TxtLuogo.Text
End Sub

Private Sub TxtNazione_Change(Index As Integer)
  AP.ImpostaNazione TxtNazione(Index).Text, Index
End Sub

Private Sub TxtNic_Change()
  AP.ImpostaNic TxtNic.Text
End Sub

Private Sub TxtNome_Change(Index As Integer)
  AP.ImpostaNome TxtNome(Index).Text, Index
End Sub

Private Sub TxtNumeroScacchiera_Change()
  On Local Error Resume Next
  AP.ImpostaNumeroScacchiera TxtNumeroScacchiera.Text
End Sub

Private Sub TxtOra_Change()
  On Local Error Resume Next
  AP.ImpostaData DateValue(TxtOra.Text)
End Sub

Private Sub TxtSottoApertura_Change()
  AP.ImpostaSottoApertura TxtSottoApertura.Text
End Sub

Private Sub TxtSottoSottoApertura_Change()
  AP.ImpostaSottoSottoApertura TxtSottoSottoApertura.Text
End Sub

Private Sub TxtSponsor_Change()
  AP.ImpostaSponsor TxtSponsor.Text
End Sub

Private Sub TxtUscf_Change(Index As Integer)
  AP.ImpostaUscf TxtUscf(Index).Text, Index
End Sub

