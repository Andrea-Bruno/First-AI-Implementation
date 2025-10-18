VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Archivio 
   Caption         =   "Archivio"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Doppio click sulla partita per caricarla"
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Avvia ricerca"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Inserite il testo della ricerca"
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Trova:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Archivio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SeekRecord() As Long
Dim AP As Variant
Public ListCount As Long
Const Msg1 = "Terminata ricerca nell'archivio con esito negarivo!"
Const Msg2 = "Caricamento indice"
Const Msg3 = "Archivio enorme! (Non lo si può visualizzare tutto)"
Const FormattazioneGrid = "<Evento                       |<Luogo                                 |>           Data|>Round|<Bianco                      |<Nero                        |>Risultato|>     ECO"
Option Explicit


Private Sub Form_Load()
  MSFlexGrid.FormatString = FormattazioneGrid
End Sub

Private Sub Form_Terminate()
  Set AP = Nothing
End Sub
Private Sub Form_Initialize()
  Set AP = ModuloGioco.AP
End Sub
Public Function SeekPartita(Index)
  SeekPartita = SeekRecord(Index)
End Function
Private Sub Command1_Click()
  Dim s() As String
  Dim R As String
  Dim P As Integer
  Dim OP As Integer
  Dim NS As Integer
  Dim C As Integer
  R = UCase(Trim(Text.Text))
  If Len(R) <> 0 And MSFlexGrid.Rows > 1 Then
    R = R + " "
    Do
      OP = P
      P = InStr(P + 1, R, " ")
      NS = NS + 1
      ReDim Preserve s(NS)
      s(NS) = Mid(R, OP + 1, P - OP - 1)
    Loop Until P = Len(R)
    Dim N As Long
    Dim CS As String
    Dim N2 As Integer
    Dim Fl As Boolean
    N = MSFlexGrid.Row + 1
    Do
      If N >= MSFlexGrid.Rows Or N < 1 Then
        N = 1
      End If
      CS = ""
      'Crea una stringa contenente tutte le colonne
      For C = 0 To MSFlexGrid.Cols - 1
        CS = CS + MSFlexGrid.TextMatrix(N, C)
      Next
      CS = UCase(CS)
      Fl = True
      For N2 = 1 To NS
        If InStr(CS, s(N2)) = False Then
          Fl = False
          Exit For
        End If
      Next
      If Fl = True Then
        MSFlexGrid.Row = N
        MSFlexGrid.TopRow = N
        MSFlexGrid.SetFocus
        Exit Sub
      End If
      If N = MSFlexGrid.Row Then
        MsgBox Msg1, vbInformation
        Exit Sub
      End If
      N = N + 1
    Loop
  End If
End Sub
Private Sub Form_Resize()
  Dim h As Integer
  h = ScaleHeight - 600
  If h < 1 Then
    h = 1
  End If
  MSFlexGrid.Move 0, 600, ScaleWidth, h
End Sub
Public Sub LeggiFileArchivio(File As String, Optional NonCreareLista As Boolean)
  Dim NumeroFile As Integer
  Dim SK As Long
  If Dir(File) <> "" Then
    Caption = File
    NumeroFile = FreeFile
    Dim Prs As New Progressione
    Prs.Show 0, Me
    Prs.Caption = Msg2
    Open File For Input As #NumeroFile
      Dim VList As String
      Dim PosizioneSuccessivo As Long
      ListCount = 0
      Prs.Min = 1
      Prs.Max = LOF(NumeroFile)
      SK = 0
      'MSFlexGrid.ListIndex = 2
      Do
        VList = Elaborazione.Partita(File, SK, PosizioneSuccessivo)
        If ListCount Mod 1000 = 0 Then
          Prs.Value = SK
        End If
        ReDim Preserve SeekRecord(ListCount)
        SeekRecord(ListCount) = SK
        ListCount = ListCount + 1
        If NonCreareLista = False Then
          MSFlexGrid.AddItem VList, ListCount
          If ListCount >= 43000 Then
            MsgBox Msg3, vbInformation
            Exit Do
          End If
        End If
        SK = PosizioneSuccessivo
      Loop Until PosizioneSuccessivo = 0
    Close #NumeroFile
    Unload Prs
  End If
End Sub
Private Sub MSFlexGrid_DblClick()
  Carica
End Sub

Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 32 Then
    Carica
  End If
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Command1_Click
  End If
End Sub
Public Sub Carica()
  If MSFlexGrid.Row < MSFlexGrid.Rows - 1 Then
    AP.CaricaPartita Caption, SeekRecord(MSFlexGrid.Row - 1)
  End If
End Sub
