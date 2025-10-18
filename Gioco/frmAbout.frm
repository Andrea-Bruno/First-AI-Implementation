VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su MiaApplicazione"
   ClientHeight    =   4230
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   8340
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   2919.621
   ScaleMode       =   0  'User
   ScaleWidth      =   7831.691
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   "Sequencer"
      FileName        =   ""
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5520
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Appearance      =   0  'Flat
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   6960
      TabIndex        =   1
      Top             =   3840
      Width           =   1245
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":7ED56
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Titolo applicazione"
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   3165
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Versione"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   3165
   End
   Begin VB.Label lblDisclaimer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Avviso: Tutti i diritti sono riservati"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   4590
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opzioni di protezione per la chiave del registro di configurazione
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Chiavi principali del registro di configurazione
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Stringa Unicode che termina con un carattere Null
Const REG_DWORD = 4                      ' Numero a 32 bit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Informazioni su " & App.Title
    lblVersion.Caption = "Versione " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    If Dir(App.Path + "\Alexisso.mid") <> "" Then
      MMControl1.FileName = App.Path + "\Alexisso.mid"
      MMControl1.Notify = False
      MMControl1.Wait = True
      MMControl1.Shareable = False
      MMControl1.DeviceType = "Sequencer"
      MMControl1.Command = "Open"
      MMControl1.Command = "Play"
    End If
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Tenta di recuperare dal registro di configurazione il percorso e il nome
    ' del programma che consente di visualizzare le informazioni sul sistema
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Tenta di recuperare dal registro di configurazione solo il percorso
    ' del programma che consente di visualizzare le informazioni sul sistema
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Convalida l'esistenza di una versione a 32 bit del file conosciuta
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Errore. Il file non è stato trovato.
        Else
            GoTo SysInfoErr
        End If
    ' Errore. La chiave del registro di configurazione non è stata trovata.
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Le informazioni sul sistema non sono attualmente disponibili.", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contatore per il ciclo
    Dim rc As Long                                          ' Codice restituito
    Dim hKey As Long                                        ' Handle a una chiave del registro di configurazione aperta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo di dati di una chiave del registro di configurazione
    Dim tmpVal As String                                    ' Posizione di memorizzazione temporanea del valore di una chiave del registro di configurazione
    Dim KeyValSize As Long                                  ' Dimensioni della variabile della chiave del registro di configurazione
    '------------------------------------------------------------
    ' Apre una chiave del registro di configurazione in una chiave principale {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Apre la chiave del registro di configurazione
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestione degli errori
    
    tmpVal = String$(1024, 0)                             ' Assegna spazio alla variabile
    KeyValSize = 1024                                       ' Specifica le dimensioni della variabile
    
    '------------------------------------------------------------
    ' Recupera il valore della chiave del registro di configurazione
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Recupera/crea il valore della chiave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestione degli errori
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' In Win95 viene aggiunta una stringa che termina con un carattere Null
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' È stato trovato un carattere Null, che viene estratto dalla stringa
    Else                                                    ' In WinNT non viene aggiunto un carattere Null al termine della stringa
        tmpVal = Left(tmpVal, KeyValSize)                   ' Non è stato trovato nessun carattere Null, pertanto estrae solo la stringa
    End If
    '------------------------------------------------------------
    ' Determina il tipo del valore della chiave per la conversione
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Cerca i tipi di dati
    Case REG_SZ                                             ' Tipo di dati String per la chiave del registro di configurazione
        KeyVal = tmpVal                                     ' Copia il valore String
    Case REG_DWORD                                          ' Tipo di dati Double Word per la chiave del registro di configurazione
        For i = Len(tmpVal) To 1 Step -1                    ' Converte ogni bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Crea il valore carattere per carattere
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Converte Double Word in String
    End Select
    
    GetKeyValue = True                                      ' Restituisce un valore che indica che l'operazione è riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione
    Exit Function                                           ' Esce dalla routine
    
GetKeyError:      ' Reimposta i dati se viene generato un errore
    KeyVal = ""                                             ' Imposta su una stringa vuota il valore restituito
    GetKeyValue = False                                     ' Restituisce un valore che indica che l'operazione non è riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione
End Function

Private Sub Form_Unload(Cancel As Integer)
  MMControl1.Command = "Close"
End Sub
