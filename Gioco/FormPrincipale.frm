VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{706DFD08-82E6-11D1-AD79-F03E07C10000}#96.0#0"; "ControlloScacchiera.ocx"
Begin VB.MDIForm MDIScacchi 
   BackColor       =   &H8000000C&
   Caption         =   "Scacchi"
   ClientHeight    =   7860
   ClientLeft      =   4710
   ClientTop       =   3630
   ClientWidth     =   9390
   Icon            =   "FormPrincipale.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList FigureScacchi 
      Left            =   720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   1905
      BandCount       =   5
      VariantHeight   =   0   'False
      _CBWidth        =   9390
      _CBHeight       =   1080
      _Version        =   "6.0.8169"
      Child1          =   "Frame1"
      MinHeight1      =   495
      Width1          =   2520
      NewRow1         =   0   'False
      Child2          =   "Frame2"
      MinHeight2      =   495
      Width2          =   2115
      NewRow2         =   0   'False
      Child3          =   "Frame4"
      MinHeight3      =   495
      Width3          =   1095
      NewRow3         =   0   'False
      Child4          =   "Frame5"
      MinWidth4       =   2640
      MinHeight4      =   495
      NewRow4         =   0   'False
      Child5          =   "Frame3"
      MinHeight5      =   495
      Width5          =   4440
      NewRow5         =   -1  'True
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5985
         TabIndex        =   21
         Top             =   30
         Width           =   3315
         Begin VB.CommandButton Successiva 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   480
            Picture         =   "FormPrincipale.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Vai alla mossa successiva"
            Top             =   0
            Width           =   495
         End
         Begin VB.ComboBox Editabilit‡ 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Precedente 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   0
            Picture         =   "FormPrincipale.frx":0663
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Vai alla mossa precedente"
            Top             =   0
            Width           =   495
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   2760
            Shape           =   3  'Circle
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4860
         TabIndex        =   16
         Top             =   30
         Width           =   900
         Begin VB.CommandButton Commento 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   480
            Picture         =   "FormPrincipale.frx":09B9
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Commenta la posizione attuale"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Analisi 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   0
            Picture         =   "FormPrincipale.frx":0D4F
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Analizza la posizione attuale"
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   165
         TabIndex        =   13
         Top             =   555
         Width           =   9135
         Begin VB.CommandButton SvuotaScacchiera 
            Height          =   495
            Left            =   480
            Picture         =   "FormPrincipale.frx":10DA
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Toglie tutti i pezzi dalla scacchiera"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Edita 
            Height          =   495
            Left            =   0
            Picture         =   "FormPrincipale.frx":1463
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Disponi i pezzi nella scacchiera"
            Top             =   0
            Width           =   495
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   0
            Left            =   960
            TabIndex        =   30
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   1
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   1
            Left            =   1440
            TabIndex        =   31
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   2
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   2
            Left            =   1920
            TabIndex        =   32
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   4
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   3
            Left            =   2400
            TabIndex        =   33
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   8
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   4
            Left            =   2880
            TabIndex        =   34
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   12
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   5
            Left            =   3360
            TabIndex        =   35
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   16
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   6
            Left            =   3840
            TabIndex        =   36
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -1
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   7
            Left            =   4320
            TabIndex        =   37
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -2
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   8
            Left            =   4800
            TabIndex        =   38
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -4
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   9
            Left            =   5280
            TabIndex        =   39
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -8
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   10
            Left            =   5760
            TabIndex        =   40
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -12
         End
         Begin ControlloScacchiera.Casa Casa 
            Height          =   495
            Index           =   11
            Left            =   6240
            TabIndex        =   41
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            ScaleHeight     =   435
            ScaleMode       =   0
            ScaleWidth      =   435
            Figura          =   -16
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2715
         TabIndex        =   5
         Top             =   30
         Width           =   1920
         Begin VB.CommandButton Avanti 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   2400
            Picture         =   "FormPrincipale.frx":17DA
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Ripristina ultima operazione"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Indietro 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1920
            Picture         =   "FormPrincipale.frx":1B33
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Elimina ultima operazione"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Duplica 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1440
            Picture         =   "FormPrincipale.frx":1E89
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Duplica la scacchiera selezionata"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Incolla 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   960
            Picture         =   "FormPrincipale.frx":221F
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Incolla scacchiera dalla memoria"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Copia 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   480
            Picture         =   "FormPrincipale.frx":25C7
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Copia Scacchiera in memoria"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Taglia 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   0
            Picture         =   "FormPrincipale.frx":294E
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Taglia scacchiera selezionata"
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   2325
         Begin VB.CommandButton Blocca 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1920
            Picture         =   "FormPrincipale.frx":2CC2
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Blocca la scacchiera."
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Stampa 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1440
            Picture         =   "FormPrincipale.frx":304C
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Stampa la situazione di gioco attuale"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Salva 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   960
            Picture         =   "FormPrincipale.frx":33DB
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Salva scacchiera"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Apri 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   480
            Picture         =   "FormPrincipale.frx":3761
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Carica Scacchiera"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Nuovo 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   0
            Picture         =   "FormPrincipale.frx":3AE7
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Nuova scacchiera"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Aiuto 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   2400
            Picture         =   "FormPrincipale.frx":3E62
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Mostra la dovumentazione"
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9390
      TabIndex        =   2
      Top             =   1080
      Width           =   9390
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9390
      TabIndex        =   3
      Top             =   1080
      Width           =   9390
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9390
      TabIndex        =   4
      Top             =   1080
      Width           =   9390
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9390
      TabIndex        =   8
      Top             =   1080
      Width           =   9390
   End
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9390
      TabIndex        =   20
      Top             =   1080
      Width           =   9390
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuNuovo 
         Caption         =   "&Nuovo gioco"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuApri 
         Caption         =   "&Apri gioco"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSalva 
         Caption         =   "&Salva gioco"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSalvaConNome 
         Caption         =   "&Salva gioco con no&me"
      End
      Begin VB.Menu mnuStampa 
         Caption         =   "S&tampa"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "&Esci"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuModifica 
      Caption         =   "&Modifica"
      Begin VB.Menu mnuAvanti 
         Caption         =   "&Avanti"
      End
      Begin VB.Menu mnuIndietro 
         Caption         =   "&Indietro"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuTaglia 
         Caption         =   "&Taglia"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopia 
         Caption         =   "&Copia"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuIncolla 
         Caption         =   "&Incolla"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDuplica 
         Caption         =   "&Duplica"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuScacchiera 
      Caption         =   "&Scacchiera"
      Begin VB.Menu mnuMuovi 
         Caption         =   "&Muovi"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuGiocatore 
         Caption         =   "&Giocatore"
         Begin VB.Menu MnuGiocatoreBianco 
            Caption         =   "&Bianco"
            Begin VB.Menu mnuGBComputer 
               Caption         =   "&Computer"
            End
            Begin VB.Menu MnuGBUmano 
               Caption         =   "&Umano"
            End
         End
         Begin VB.Menu MnuGiocatoreNero 
            Caption         =   "&Nero"
            Begin VB.Menu MnuGNComputer 
               Caption         =   "&Computer"
            End
            Begin VB.Menu MnuGNUmano 
               Caption         =   "&Umano"
            End
         End
      End
      Begin VB.Menu mnuOrientamento 
         Caption         =   "&Orientamento"
         Begin VB.Menu mnuOrientamentoBianco 
            Caption         =   "&Bianco"
            Begin VB.Menu mnuBianco 
               Caption         =   "In &basso"
               Index           =   0
            End
            Begin VB.Menu mnuBianco 
               Caption         =   "In &alto"
               Index           =   1
            End
            Begin VB.Menu mnuBianco 
               Caption         =   "A &sinistra"
               Index           =   2
            End
            Begin VB.Menu mnuBianco 
               Caption         =   "A &destra"
               Index           =   3
            End
         End
         Begin VB.Menu mnuOrientamentoNero 
            Caption         =   "&Nero"
            Begin VB.Menu mnuNero 
               Caption         =   "In &basso"
               Index           =   0
            End
            Begin VB.Menu mnuNero 
               Caption         =   "In &alto"
               Index           =   1
            End
            Begin VB.Menu mnuNero 
               Caption         =   "A &sinistra"
               Index           =   2
            End
            Begin VB.Menu mnuNero 
               Caption         =   "A &destra"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuNota 
         Caption         =   "&Nota"
      End
      Begin VB.Menu MnuElencoMosse 
         Caption         =   "&Elenco mosse"
      End
      Begin VB.Menu mnuCambiaSet 
         Caption         =   "Cambia il set delle  &figure"
      End
      Begin VB.Menu mnuCoordinate 
         Caption         =   "&Coordinate"
      End
      Begin VB.Menu mnuTipoSelezione 
         Caption         =   "&Tipo selezione casa"
         Begin VB.Menu mnuBordo 
            Caption         =   "&Spessore Bordo"
         End
         Begin VB.Menu mnuEvidenzia 
            Caption         =   "&Evidenzia superficie"
         End
      End
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "&Strumenti"
      Begin VB.Menu mnuAnalisi 
         Caption         =   "&Analisi"
      End
      Begin VB.Menu mnuOrologio 
         Caption         =   "&Orologio"
      End
      Begin VB.Menu MnuPensiero 
         Caption         =   "&Pensiero"
      End
   End
   Begin VB.Menu mnuProgetto 
      Caption         =   "&Progetto"
      Begin VB.Menu mnuEdita 
         Caption         =   "Edita scacchiera"
         Begin VB.Menu mnuPosizioneIniziale 
            Caption         =   "&Imposta i pezzi nella posizione iniziale"
         End
         Begin VB.Menu mnuSvuotaScacchiera 
            Caption         =   "&Svuota la scacchiera"
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &pedone bianco"
            Index           =   0
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &cavallo bianco"
            Index           =   1
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &alfiere bianco"
            Index           =   2
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &torre bianca"
            Index           =   3
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &dama bianca"
            Index           =   4
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &re bianco"
            Index           =   5
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &pedone nero"
            Index           =   6
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &cavallo nero"
            Index           =   7
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &alfiere nero"
            Index           =   8
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &torre nera"
            Index           =   9
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &dama nera"
            Index           =   10
         End
         Begin VB.Menu mnuAggiungi 
            Caption         =   "Aggiungi &Re Nero"
            Index           =   11
         End
      End
      Begin VB.Menu MnuGioca 
         Caption         =   "Gioca"
         Begin VB.Menu mnuMossaBianco 
            Caption         =   "Mossa al bianco"
         End
         Begin VB.Menu mnuMossaNero 
            Caption         =   "Mossa al nero"
         End
      End
   End
   Begin VB.Menu mnuFinestra 
      Caption         =   "Fi&nestra"
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu mnuAffianca 
         Caption         =   "&Affianca"
      End
      Begin VB.Menu mnuSovrapponi 
         Caption         =   "&Sovrapponi"
      End
      Begin VB.Menu mnuDisponi 
         Caption         =   "&Disponi Icone"
      End
   End
   Begin VB.Menu mnuAiuto 
      Caption         =   "?"
      Begin VB.Menu mnuVersione 
         Caption         =   "&Versione"
      End
      Begin VB.Menu mnuInformazioni 
         Caption         =   "&Informazioni sul programma"
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "&Extra"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIScacchi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Dim AP As Variant


Private Sub MDIForm_Initialize()
  Set AP = ModuloGioco.AP
End Sub
Private Sub Analisi_Click()
    AP.MostraAnalisi Me
End Sub

Private Sub Apri_Click()
  AP.Apri
End Sub


Private Sub Blocca_Click()
  AP.Blocca
End Sub

Private Sub Casa_Click(Index As Integer)
  AP.InserisciPezzo Casa(Index).figura
End Sub

Private Sub Duplica_Click()
  AP.Duplica
End Sub

Private Sub Edita_Click()
  AP.Edita
End Sub

Private Sub Editabilit‡_Click()
    Dim E As EditorScacchiera
    Select Case Editabilit‡.ListIndex
      Case 0
        E = NonEditabile
      Case 1
        E = Editabile
      Case 2
        E = BiancoMossePossibili
      Case 3
        E = NeroMossePossibili
    End Select
    AP.Editabilit‡Scacchiera E
End Sub

Private Sub Incolla_Click()
  AP.Incolla
End Sub


Private Sub MDIForm_Load()
  Top = 0: Left = 0: Width = Screen.Width: Height = Screen.Height
  Editabilit‡.AddItem "Blocco"
  Editabilit‡.AddItem "Editazione"
  Editabilit‡.AddItem "Mossa al bianco"
  Editabilit‡.AddItem "Mossa al nero"
  Editabilit‡.ListIndex = 0
  AP.AggiornaMDIScacchi
  Dim PthPz As String
  PthPz = Casa(0).Path + "\pictures\figure\Default"
  AP.CaricaFigure PthPz
  frmAbout.Show 1, Me
  'Elaborazione.CaricaIndiceArchivio AP.Path + "\Archivio\Archivio.Sit"
  ChDir AP.Path + "\archivio"
End Sub

Private Sub MDIForm_Terminate()
  Set AP = Nothing
  ModuloGioco.Terminate
End Sub

Private Sub mnuAggiungi_Click(Index As Integer)
    AP.InserisciPezzo Casa(Index).figura
End Sub

Private Sub mnuAnalisi_Click()
  AP.MostraAnalisi Me
End Sub

Private Sub mnuApri_Click()
  AP.Apri
End Sub

Private Sub mnuAvanti_Click()
  AP.Avanti
End Sub

Private Sub mnuBianco_Click(Index As Integer)
  AP.CambiaOrientamento (Index)
End Sub

Private Sub mnuCambiaSet_Click()
  AP.CaricaFigure
End Sub

Private Sub MnuElencoMosse_Click()
  AP.MostraElencoMosse Me
End Sub

Private Sub mnuEvidenzia_Click()
 Dim Valore As TipoDiSelezione
 If mnuEvidenzia.Checked = False Then
    Valore = 2
  End If
  If mnuBordo.Checked = True Then
    Valore = Valore Or 1
  End If
  AP.TipoDiSelezioneScacchiera Valore
End Sub
Private Sub mnuBordo_Click()
  Dim Valore As TipoDiSelezione
  If mnuBordo.Checked = False Then
    Valore = 1
  End If
  If mnuEvidenzia.Checked = True Then
    Valore = Valore Or 2
  End If
  AP.TipoDiSelezioneScacchiera Valore
End Sub

Private Sub mnuCoordinate_Click()
  AP.CoordinateScacchiera Not (mnuCoordinate.Checked)
End Sub

Private Sub mnuCopia_Click()
  AP.Copia
End Sub

Private Sub mnuDisponi_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuDuplica_Click()
  AP.Duplica
End Sub

Private Sub mnuEsci_Click()
  Unload Me
End Sub



Private Sub mnuExtra_Click()
  Extra.Show 0, Me
End Sub

Private Sub mnuGBComputer_Click()
  AP.ImpostaGiocatore Bianco, Computer
End Sub

Private Sub MnuGBUmano_Click()
  AP.ImpostaGiocatore Bianco, Umano
End Sub

Private Sub MnuGNComputer_Click()
  AP.ImpostaGiocatore Nero, Computer
End Sub

Private Sub MnuGNUmano_Click()
  AP.ImpostaGiocatore Nero, Umano
End Sub

Private Sub mnuIncolla_Click()
  AP.Incolla
End Sub

Private Sub mnuIndietro_Click()
  AP.Indietro
End Sub

Private Sub mnuInformazioni_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuMossaBianco_Click()
  AP.Turno Bianco
End Sub

Private Sub mnuMossaNero_Click()
  AP.Turno Nero
End Sub

Private Sub mnuMuovi_Click()
  AP.Muovi
End Sub

Private Sub mnuNero_Click(Index As Integer)
  Dim O As Integer
  If Index = 0 Then O = 1
  If Index = 1 Then O = 0
  If Index = 2 Then O = 3
  If Index = 3 Then O = 2
  AP.CambiaOrientamento (O)
End Sub

Private Sub mnuNota_Click()
  AP.MostraNota Me
End Sub

Private Sub mnuNuovo_Click()
  AP.Nuovo True
End Sub
Private Sub mnuAffianca_Click()
  Me.Arrange vbTileHorizontal
End Sub


Private Sub mnuOrologio_Click()
  AP.MostraOrologio Me
End Sub

Private Sub MnuPensiero_Click()
  AP.MostraPensiero Me
End Sub

Private Sub mnuPosizioneIniziale_Click()
  AP.ScacchieraPosizioneIniziale
End Sub

Private Sub mnuSalva_Click()
  AP.Salva
End Sub

Private Sub mnuSalvaConNome_Click()
  AP.SalvaConNome
End Sub

Private Sub mnuSovrapponi_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuStampa_Click()
  AP.Stampa
End Sub

Private Sub mnuSvuotaScacchiera_Click()
  AP.SvuotaScacchiera
End Sub

Private Sub mnuTaglia_Click()
  AP.Taglia
End Sub

Private Sub mnuVersione_Click()
  frmSplash.Show 0, Me
End Sub

Private Sub Nuovo_Click()
  AP.Nuovo True
End Sub

Private Sub Precedente_Click()
  AP.Indietro
End Sub

Private Sub Salva_Click()
  AP.Salva
End Sub

Private Sub Stampa_Click()
  AP.Stampa
End Sub

Private Sub Successiva_Click()
  AP.Avanti
End Sub

Private Sub SvuotaScacchiera_Click()
  AP.SvuotaScacchiera
End Sub

Private Sub Taglia_Click()
  AP.Taglia
End Sub
Private Sub Copia_Click()
  AP.Copia
End Sub
