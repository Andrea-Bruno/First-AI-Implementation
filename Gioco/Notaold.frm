VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Nota 
   ClientHeight    =   6360
   ClientLeft      =   105
   ClientTop       =   2295
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtOra 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn:ss AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   48
      Tag             =   "1"
      ToolTipText     =   "Ora dell'incontro"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtLuogo 
      Height          =   375
      Left            =   1080
      TabIndex        =   47
      Tag             =   "1"
      ToolTipText     =   "Posto in cui si è svolta la partita"
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox TxtData 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   46
      Tag             =   "1"
      ToolTipText     =   "Data dell'incontro"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox TxtSponsor 
      Height          =   375
      Left            =   1080
      TabIndex        =   45
      Tag             =   "1"
      ToolTipText     =   "Finanziatore della gara"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.ComboBox CmbSessione 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Tag             =   "1"
      ToolTipText     =   "Sessione dell'evento (in particolare descrive la possibilita di accedervi)"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox CmbTurno 
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Tag             =   "1"
      ToolTipText     =   "Turno di gara nel caso che si tratti di un torneo o campionato"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox TxtNumeroScacchiera 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0000;(0000)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   42
      Tag             =   "1"
      ToolTipText     =   "Numero nella scacchiera nel caso in cui si sta giocando in un torneo con scacchiere numerate"
      Top             =   2640
      Width           =   495
   End
   Begin VB.ComboBox CmbRisultato 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Tag             =   "1"
      ToolTipText     =   "Indica quale giocatore ha ottenuto la vittoria o se vi è stato un pareggio"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox TxtApertura 
      Height          =   375
      Left            =   3960
      TabIndex        =   40
      Tag             =   "1"
      ToolTipText     =   "Indica il tipo di apertura adottata in questa partita"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TxtSottoApertura 
      Height          =   375
      Left            =   1080
      TabIndex        =   39
      Tag             =   "1"
      ToolTipText     =   "Indica la sottoapertura adottata nella partita"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox TxtSottoSottoApertura 
      Height          =   375
      Left            =   3960
      TabIndex        =   38
      Tag             =   "1"
      ToolTipText     =   "Indica una variante della SottoApertura"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox TxtEco 
      Height          =   375
      Left            =   1080
      TabIndex        =   37
      Tag             =   "1"
      ToolTipText     =   "Indica il codice di apertura riportato nella ""Encyclopedia of chess openings"""
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox TxtNic 
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Tag             =   "1"
      ToolTipText     =   "Codice di riferimento al database delle aperture"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox CmbTerminato 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Tag             =   "1"
      ToolTipText     =   "Indica in che modo ha avuto fine l'incontro"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox TxtAnnotatore 
      Height          =   375
      Left            =   3960
      TabIndex        =   34
      Tag             =   "1"
      ToolTipText     =   "Cognome e nome dell'annotatore"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox TxtExtra 
      Height          =   375
      Left            =   3960
      TabIndex        =   33
      Tag             =   "1"
      ToolTipText     =   "Campo libero per informazioni aggiuntive"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ComboBox CmbModo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Tag             =   "1"
      ToolTipText     =   "Indica lo strumento che ha permesso di svolgere la partita"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.ComboBox CmbTipoGara 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Tag             =   "1"
      ToolTipText     =   "Sessione dell'evento (in particolare descrive la possibilita di accedervi)"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TxtIncontro 
      Height          =   375
      Left            =   1080
      TabIndex        =   30
      Tag             =   "1"
      ToolTipText     =   "Indica il numero di partita nel caso in cui vi è una numerazione delle partite"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox TxtNome 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      Tag             =   "3"
      ToolTipText     =   "Cognome e nome del giocatore"
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox TxtElo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "00000;(00000)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Tag             =   "3"
      ToolTipText     =   "Fide ELO"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtUscf 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   20
      Tag             =   "3"
      ToolTipText     =   "Indice USCF"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CmbTitolo 
      Height          =   315
      Index           =   0
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "3"
      ToolTipText     =   "Titolo in qualita di giocatore"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtEmail 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   18
      Tag             =   "3"
      ToolTipText     =   "Iindirizzo E-MAIL o Network del giocatore"
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   17
      Tag             =   "3"
      ToolTipText     =   "Titolo in qualita di giocatore"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtNazione 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Tag             =   "3"
      ToolTipText     =   "Nazionalità del giocatore"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtNazione 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Tag             =   "2"
      ToolTipText     =   "Nazionalità del giocatore"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Tag             =   "2"
      ToolTipText     =   "Titolo in qualita di giocatore"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TxtEmail 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Iindirizzo E-MAIL o Network del giocatore"
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ComboBox CmbTitolo 
      Height          =   315
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Titolo in qualita di giocatore"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtUscf 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Indice USCF"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtElo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "00000;(00000)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Fide ELO"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtNome 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Cognome e nome del giocatore"
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox Text 
      Height          =   855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5520
      Width           =   6135
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1138
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Incontro"
            Key             =   "Incontro"
            Object.ToolTipText     =   "Imposta le caratteristiche della partita di gioco"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bianco"
            Key             =   "GiocatoreBianco"
            Object.ToolTipText     =   "Imposta le caratteristiche del giocatoire che ha il bianco"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nero"
            Key             =   "GiocatoreNero"
            Object.ToolTipText     =   "Imposta le caratteristiche del giocatoire che ha il nero"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Luogo:"
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Tag             =   "1"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Data:"
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Tag             =   "1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ora:"
      Height          =   255
      Left            =   2160
      TabIndex        =   65
      Tag             =   "1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Sponsor:"
      Height          =   255
      Left            =   240
      TabIndex        =   64
      Tag             =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Sessione:"
      Height          =   255
      Left            =   240
      TabIndex        =   63
      Tag             =   "1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Turno:"
      Height          =   255
      Left            =   3120
      TabIndex        =   62
      Tag             =   "1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Incontro:"
      Height          =   255
      Left            =   240
      TabIndex        =   61
      Tag             =   "1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero della scacchiera:"
      Height          =   255
      Left            =   1800
      TabIndex        =   60
      Tag             =   "1"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Risultato:"
      Height          =   255
      Left            =   240
      TabIndex        =   59
      Tag             =   "1"
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Apertura:"
      Height          =   255
      Left            =   3120
      TabIndex        =   58
      Tag             =   "1"
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Variante:"
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Tag             =   "1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Sottovariante:"
      Height          =   255
      Left            =   2880
      TabIndex        =   56
      Tag             =   "1"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "ECO:"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Tag             =   "1"
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "NIC:"
      Height          =   255
      Left            =   3120
      TabIndex        =   54
      Tag             =   "1"
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Conclusione:"
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Tag             =   "1"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Annotatore:"
      Height          =   255
      Left            =   2880
      TabIndex        =   52
      Tag             =   "1"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Modo:"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Tag             =   "1"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Extra:"
      Height          =   255
      Left            =   3000
      TabIndex        =   50
      Tag             =   "1"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Tag             =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Tag             =   "3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Elo:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Tag             =   "3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Uscf:"
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Tag             =   "3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Titolo:"
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Tag             =   "3"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Tag             =   "3"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Tag             =   "3"
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "Nazione:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Tag             =   "3"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "Nazione:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Tag             =   "2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Tag             =   "2"
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Mail:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Tag             =   "2"
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Titolo:"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Tag             =   "2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Caption         =   "Uscf:"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Tag             =   "2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "Elo:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Tag             =   "2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Tag             =   "2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   68
      Top             =   840
      Width           =   6135
   End
End
Attribute VB_Name = "Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
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
  'base=tbsOptions.Height
  base = 5565
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
  
  Dim obj As Object
  For Each obj In Me.Controls
    If obj.Tag <> "" Then
      If Val(obj.Tag) <> tbsOptions.SelectedItem.Index Then
        obj.Visible = False
      Else
        obj.Visible = True
      End If
    End If
  Next
  Select Case tbsOptions.SelectedItem.Index
    Case 1
      Label34.Caption = ""
    Case 2
      Label34.Caption = "Dati del giocatore nero"
    Case 3
      Label34.Caption = "Dati del giocatore bianco"
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
