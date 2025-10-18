VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Progressione 
   Caption         =   "Operazione in corso"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMBAnnulla 
      Caption         =   "&Annulla"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Attendere la fine del processo !"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo residuo stimato:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Progressione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim T As Variant
  Public Annulla As Boolean

Private Sub CMBAnnulla_Click()
  Annulla = True
End Sub

Private Sub Form_Initialize()
  T = Now
End Sub
Public Property Get Min() As Variant
    Min = ProgressBar.Min
End Property
Public Property Let Min(ByVal New_Min As Variant)
    ProgressBar.Min = New_Min
End Property
Public Property Get Max() As Variant
    Max = ProgressBar.Max
End Property
Public Property Let Max(ByVal New_Max As Variant)
  If New_Max > 1 Then
    ProgressBar.Max = New_Max
  End If
End Property
Public Property Get Value() As Variant
    Value = ProgressBar.Value
End Property
Public Property Let Value(ByVal New_Value As Variant)
    ProgressBar.Value = New_Value
    Dim Med As Variant
    Dim Rimanenti As Long
    Med = (Now - T) / (1 + New_Value - ProgressBar.Min)
    Rimanenti = ProgressBar.Max - New_Value
    Dim TStimato As Variant
    TStimato = Med * Rimanenti
    Label.Caption = Format(TStimato, "hh.mm.ss")
    Label3.Caption = LTrim(Str(New_Value)) + "/" + LTrim(Str(ProgressBar.Max))
    DoEvents
End Property

