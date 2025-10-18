VERSION 5.00
Begin VB.UserControl Casa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3990
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2895
      Left            =   600
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2625
   End
End
Attribute VB_Name = "Casa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'DefInt A-Z
Option Explicit
'Valori predefiniti proprietà:
Const m_def_ColoreCasa = Bianco
Const m_def_Figura = 0
Const m_def_ColoreFigura = Bianco
Const ColoreBianco As Long = &H8000000F
Const ColoreNero As Long = &H80000010
Const ColoreSelezionato As Long = &H8000000D
Const m_def_Selezionato As Boolean = False
Const m_def_TipoDiSelezione = Bordo
Const m_def_AggiornaFigura = True
Const Msg1 = "Directory Pictures mancante in "
'Variabili proprietà:
Dim m_def_PathFigure As String
Dim m_TipoDiSelezione As TipoDiSelezione
Dim m_PathFigure As String
Dim m_ColoreCasa As Colore
Dim m_Selezionato As Boolean
Dim m_AggiornaFigura As Boolean
'Dim m_PathFigure As DataSource
Dim m_Figura As pezzo
Dim m_ColoreFigura As Colore
'Dichiarazioni di eventi:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Event Change()
Event Resize()
Attribute Resize.VB_Description = "Viene generato non appena un form viene visualizzato o quando le dimensioni di un oggetto vengono modificate."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Viene generato quando la proprietà Visible del controllo viene impostata su False."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Viene generato quando la proprietà Visible del controllo viene impostata su True."
Private Sub Image1_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
  m_def_PathFigure = App.Path & "\pictures\figure\Default"
End Sub

Private Sub UserControl_Resize()
  Image1.Move 0, 0, ScaleWidth, ScaleHeight
  RaiseEvent Resize
End Sub
Private Sub UserControl_Hide()
  RaiseEvent Hide
End Sub
Private Sub UserControl_Show()
  RaiseEvent Show
End Sub

Public Property Get ColoreFigura() As Colore
Attribute ColoreFigura.VB_Description = "Colore del pezzo posto sulla casella (nero o bianco)"
  ColoreFigura = m_ColoreFigura
End Property

Public Property Let ColoreFigura(ByVal New_ColoreFigura As Colore)
  m_ColoreFigura = New_ColoreFigura
  PropertyChanged "ColoreFigura"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'Inizializza le proprietà di UserControl
Private Sub UserControl_InitProperties()
  m_ColoreFigura = m_def_ColoreFigura
  m_Figura = m_def_Figura
  m_AggiornaFigura = m_def_AggiornaFigura
  m_ColoreCasa = m_def_ColoreCasa
  m_PathFigure = m_def_PathFigure
  m_TipoDiSelezione = m_def_TipoDiSelezione
  m_Selezionato = m_def_Selezionato
End Sub

'Carica i valori della proprietà dalla memoria
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3600)
  UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
  UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
  UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
  UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
  m_AggiornaFigura = PropBag.ReadProperty("AggiornaFigura", m_def_AggiornaFigura)
  m_ColoreFigura = PropBag.ReadProperty("ColoreFigura", m_def_ColoreFigura)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_Figura = PropBag.ReadProperty("Figura", m_def_Figura)
  UserControl.BackColor = m_ColoreCasa
  m_ColoreCasa = PropBag.ReadProperty("ColoreCasa", m_def_ColoreCasa)
  Select Case m_ColoreCasa
    Case Bianco
      UserControl.BackColor = ColoreBianco
    Case Nero
      UserControl.BackColor = ColoreNero
  End Select
  m_PathFigure = PropBag.ReadProperty("PathFigure", m_def_PathFigure)
  VisualizzaFigura (m_Figura)
  m_ColoreCasa = PropBag.ReadProperty("ColoreCasa", m_def_ColoreCasa)
  m_TipoDiSelezione = PropBag.ReadProperty("TipoDiSelezione ", m_def_TipoDiSelezione)
  m_Selezionato = PropBag.ReadProperty("Selezionato", m_def_Selezionato)
  ImpostaSelezione (m_Selezionato)
End Sub

'Scrive i valori della proprietà in memoria
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
  Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
  Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
  Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
  Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
  Call PropBag.WriteProperty("ColoreFigura", m_ColoreFigura, m_def_ColoreFigura)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("AggiornaFigura", m_AggiornaFigura, m_def_AggiornaFigura)
  Call PropBag.WriteProperty("Figura", m_Figura, m_def_Figura)
  Call PropBag.WriteProperty("ColoreCasa", m_ColoreCasa, m_def_ColoreCasa)
  Call PropBag.WriteProperty("PathFigure", m_PathFigure, m_def_PathFigure)
  Call PropBag.WriteProperty("TipoDiSelezione", m_TipoDiSelezione, m_def_TipoDiSelezione)
  Call PropBag.WriteProperty("Selezionato", m_Selezionato, m_def_Selezionato)
End Sub
'Public Sub WriteProperties()
'  UserControl_WriteProperties
'End Sub
'Public Sub ReadProperties()
'
'End Sub
Public Property Get AggiornaFigura() As Boolean
    AggiornaFigura = m_AggiornaFigura
End Property
Public Property Let AggiornaFigura(ByVal New_AggiornaFigura As Boolean)
    m_AggiornaFigura = New_AggiornaFigura
    PropertyChanged "AggiornaFigura"
    If m_AggiornaFigura = True Then
      VisualizzaFigura (m_Figura)
    End If
End Property

Public Property Get Figura() As pezzo
Attribute Figura.VB_Description = "Figura del gioco degli scacchi (re, dama,alfiere,cavallo,torre,pedone,nessuna)"
    Figura = m_Figura
End Property

Public Property Let Figura(ByVal new_Figura As pezzo)
  If new_Figura <> m_Figura Then
    m_Figura = new_Figura
    VisualizzaFigura (new_Figura)
    PropertyChanged "Figura"
    RaiseEvent Change
    'Image1.Refresh
  End If
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un oggetto."
  UserControl.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=23,0,0,0
Public Property Get ColoreCasa() As Colore
Attribute ColoreCasa.VB_Description = "Colore della casella in cui vi è posizionato il pezzo (nera o bianca)"
  ColoreCasa = m_ColoreCasa
End Property

Public Property Let ColoreCasa(ByVal New_ColoreCasa As Colore)
  UserControl.BackColor = CorrispondenzaColore(m_ColoreCasa)
  ImpostaSelezione (m_Selezionato)
  m_ColoreCasa = New_ColoreCasa
  PropertyChanged "ColoreCasa"
End Property
Public Property Get TipoDiSelezione() As TipoDiSelezione
  TipoDiSelezione = m_TipoDiSelezione
End Property

Public Property Let TipoDiSelezione(ByVal New_TipoDiSelezione As TipoDiSelezione)
  m_TipoDiSelezione = New_TipoDiSelezione
  PropertyChanged "TipoDiSelezione"
  ImpostaSelezione (m_Selezionato)
End Property
Public Property Get Selezionato() As Boolean
  Selezionato = m_Selezionato
End Property
Public Property Let Selezionato(ByVal New_Selezionato As Boolean)
  If New_Selezionato <> m_Selezionato Then
    m_Selezionato = New_Selezionato
    PropertyChanged "Selezionato"
    ImpostaSelezione (m_Selezionato)
  End If
End Property
Public Property Get PathFigure() As String
  PathFigure = m_PathFigure
End Property

Public Property Let PathFigure(ByVal New_PathFigure As String)
  m_PathFigure = New_PathFigure
  PropertyChanged "PathFigure"
End Property

Private Sub VisualizzaFigura(Figura As pezzo)
  If m_AggiornaFigura = True Then
    On Local Error GoTo daFile
    If Figura = Nessun_Pezzo Then
        Image1.Picture = LoadPicture()
    ElseIf Figura = Pedone_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(1).Picture
    ElseIf Figura = Cavallo_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(2).Picture
    ElseIf Figura = Alfiere_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(3).Picture
    ElseIf Figura = Torre_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(4).Picture
    ElseIf Figura = Dama_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(5).Picture
    ElseIf Figura = Re_Bianco Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(6).Picture
    ElseIf Figura = Pedone_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(7).Picture
    ElseIf Figura = Cavallo_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(8).Picture
    ElseIf Figura = Alfiere_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(9).Picture
    ElseIf Figura = Torre_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(10).Picture
    ElseIf Figura = Dama_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(11).Picture
    ElseIf Figura = Re_Nero Then
        Image1.Picture = Extender.Container.FigureScacchi.ListImages(12).Picture
    End If
Exit Sub

daFile:
    VisualizzaFiguraDaFile Figura
  End If
End Sub
Private Sub VisualizzaFiguraDaFile(Figura As pezzo)
    If Figura = Nessun_Pezzo Then
        Image1.Picture = LoadPicture()
    ElseIf Figura = Pedone_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.pedone.gif")
    ElseIf Figura = Cavallo_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.cavallo.gif")
    ElseIf Figura = Alfiere_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.alfiere.gif")
    ElseIf Figura = Torre_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.torre.gif")
    ElseIf Figura = Dama_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.dama.gif")
    ElseIf Figura = Re_Bianco Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "bianco.re.gif")
    ElseIf Figura = Pedone_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.pedone.gif")
    ElseIf Figura = Cavallo_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.cavallo.gif")
    ElseIf Figura = Alfiere_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.alfiere.gif")
    ElseIf Figura = Torre_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.torre.gif")
    ElseIf Figura = Dama_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.dama.gif")
    ElseIf Figura = Re_Nero Then
        Image1.Picture = LoadPicture(m_PathFigure & "\" & "nero.re.gif")
    End If
End Sub
Sub ImpostaSelezione(Valore As Boolean)
  If m_AggiornaFigura = True Then
    If Valore = True Then
      'Dim Spessore As Integer
      'Spessore = UserControl.Width
      'If UserControl.Height > UserControl.Width Then
      '  Spessore = UserControl.Height
      'End If
      'Spessore = Spessore / 4
      'Spessore = 10
      'DrawWidth = 1
      'UserControl.CurrentX = 0
      'UserControl.CurrentY = 0
      'UserControl.Line Step(0, 0)-(UserControl.Width, UserControl.Height), 1, BF
      NessunaEvidenziazione
      If m_TipoDiSelezione And Bordo Then
        Image1.BorderStyle = 1
      End If
      If m_TipoDiSelezione And Evidenziato Then
        UserControl.BackColor = ColoreSelezionato
      End If
    Else
      NessunaEvidenziazione
    End If
  End If
End Sub
Function CorrispondenzaColore(Colore As Colore) As Long
  If Colore = Nero Then
    CorrispondenzaColore = ColoreNero
  Else
    CorrispondenzaColore = ColoreBianco
  End If
End Function
Private Function NessunaEvidenziazione()
    Image1.BorderStyle = 0
    UserControl.BackColor = CorrispondenzaColore(m_ColoreCasa)
End Function
Public Function Path() As String
  Path = App.Path
End Function
