VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
DefInt A-Z
Sub AggiungiCommento(s As String, Sca As Object)
  s = LTrim(s)
  If Sca.Quantit‡MosseTraccia = 0 Then
    If Sca.Evento <> "" Then
      Sca.Evento = Sca.Evento + Chr(13) + Chr(10)
    End If
    Sca.Evento = Sca.Evento + s
  Else
    Dim Traccia As Variant
    Traccia = Sca.Traccia
    If Traccia(Sca.Quantit‡MosseTraccia).Nota <> "" Then
      Traccia(Sca.Quantit‡MosseTraccia).Nota = Traccia(Sca.Quantit‡MosseTraccia).Nota + Chr$(13) + Chr$(10)
    End If
    Traccia(Sca.Quantit‡MosseTraccia).Nota = Traccia(Sca.Quantit‡MosseTraccia).Nota + s
    Sca.Traccia = Traccia
  End If
End Sub
Private Function Sostituisci(Stringa As String, carattere As String, sostituto As String)
  Dim NS As String
  Dim N As Integer
  If Len(Stringa) <> 0 Then
    For N = 1 To Len(Stringa)
      Dim C As String
      C = Mid(Stringa, N, 1)
      If C = carattere Then
        C = sostituto
      End If
      NS = NS + C
    Next
  End If
  Stringa = NS
End Function
Public Sub Muovi()
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Muovi
  End If
End Sub
Public Sub ImpostaGiocatore(Colore As Colore, Tipo As TipoGiocatore)
  If EsisteGiocoSelezionato(True) = True Then
    ImpostaTipo Tipo, CInt(Colore)
    AggiornaCheckTipoGiocatore
  End If
End Sub
Public Sub ImpostaTurno(Valore As TurnoDiGioco)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoTurno = Valore
  End If
End Sub
Public Sub ImpostaTitolo(Valore As TitoloFide, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoTitolo = Valore
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroTitolo = Valore
    End If
  End If
End Sub
Public Sub ImpostaTipo(Valore As TipoGiocatore, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoTipoGiocatore = Valore
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroTipoGiocatore = Valore
    End If
  End If
End Sub
Public Sub ImpostaTerminato(Valore As TermineEvento)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoTerminato = Valore
  End If
End Sub
Public Sub ImpostaSessione(Valore As Sessione)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoSessione = Valore
  End If
End Sub
Public Sub ImpostaTipoGara(Valore As TipoGara)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoTipoGara = Valore
  End If
End Sub
Public Sub ImpostaRisultato(Valore As Risultato)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoRisultato = Valore
  End If
End Sub
Public Sub ImpostaModo(Valore As ModoDiGioco)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoModo = Valore
  End If
End Sub
Public Sub ImpostaEvento(Testo As String)
  'aggiunge una nota alla finestra di gioco selezionata
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.Evento = Testo
  End If
End Sub
Public Sub ImpostaEco(Valore As Variant)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoECO = Valore
  End If
End Sub
Public Sub ImpostaElo(Numero As Integer, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoElo = Numero
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroElo = Numero
    End If
  End If
End Sub
Public Sub ImpostaEMail(Testo As String, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoEM = Testo
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroEM = Testo
    End If
  End If
End Sub
Public Sub ImpostaExtra(Valore As Variant)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoExtra = Valore
  End If
End Sub
Public Sub ImpostaIncontro(Numero As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.eventoincontro = Numero
  End If
End Sub
Public Sub ImpostaLuogo(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoLuogo = Testo
  End If
End Sub
Public Sub ImpostaNazione(Testo As String, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoNazionalit‡ = Testo
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroNazionalit‡ = Testo
    End If
  End If
End Sub

Public Sub ImpostaNic(Valore As Variant)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoNIC = Valore
  End If
End Sub
Public Sub ImpostaNome(Testo As String, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoGiocatore = Testo
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroGiocatore = Testo
    End If
  End If
End Sub
Public Sub ImpostaNumeroScacchiera(Numero As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoNumeroScacchiera = Numero
  End If
End Sub
Public Sub ImpostaOra(Valore As Variant)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoOra = Valore
  End If
End Sub
Public Sub ImpostaSottoApertura(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoSottoApertura = Testo
  End If
End Sub
Public Sub ImpostaSottoSottoApertura(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoSottoSottoApertura = Testo
  End If
End Sub
Public Sub ImpostaSponsor(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoSponsor = Testo
  End If
End Sub
Public Sub ImpostaUscf(Valore As Variant, Colore As Integer)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = 1 Then
      MDIScacchi.ActiveForm.Scacchiera.BiancoUSCF = Valore
    Else
      MDIScacchi.ActiveForm.Scacchiera.NeroUSCF = Valore
    End If
  End If
End Sub
Public Sub ImpostaData(data As Date)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoData = data
  End If
End Sub
Public Sub ImpostaApertura(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoApertura = Testo
  End If
End Sub
Public Sub ImpostaAnnotatore(Testo As String)
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.EventoNomeAnnotatore = Testo
  End If
End Sub
Public Property Get Caption() As String
  Caption = MDIScacchi.Caption
End Property
Public Property Let Caption(vNewValue As String)
  MDIScacchi.Caption = vNewValue
End Property
Public Property Get Name() As String
  Name = App.Title
End Property
Public Property Get Path() As String
  Path = App.Path
  If InStr(UCase(Path), "SCACCHI\GIOCO") Then
    Path = Path + "\.."
  End If
End Property
Public Property Get Version() As String
  Version = App.Major & "." & App.Minor
End Property
Public Sub fine()
  Unload MDIScacchi
End Sub
Public Sub Editabilit‡Scacchiera(Editabilit‡ As EditorScacchiera)
  If EsisteGiocoSelezionato Then
    If (MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = NonEditabile Or MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = Editabile) Then
      If (Editabilit‡ = BiancoMossePossibili Or Editabilit‡ = NeroMossePossibili) Then
        MDIScacchi.ActiveForm.Scacchiera.EliminaTraccia
      End If
    End If
    MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = Editabilit‡
    AggiornaEditabilit‡
  End If
End Sub
Public Sub Avanti()
  If MDIScacchi.ActiveForm.Scacchiera.MossaAttuale < MDIScacchi.ActiveForm.Scacchiera.Quantit‡MosseTraccia Then
    MDIScacchi.ActiveForm.Scacchiera.MossaAttuale = MDIScacchi.ActiveForm.Scacchiera.MossaAttuale + 1
    AggiornaFinestre 34
  End If
End Sub
Public Sub Indietro()
  If MDIScacchi.ActiveForm.Scacchiera.MossaAttuale > 0 Then
    MDIScacchi.ActiveForm.Scacchiera.MossaAttuale = MDIScacchi.ActiveForm.Scacchiera.MossaAttuale - 1
    AggiornaFinestre 34
  End If
End Sub
Public Sub VaiAlla(NumeroMossa As Integer)
  If MDIScacchi.ActiveForm.Scacchiera.MossaAttuale > 0 Then
    MDIScacchi.ActiveForm.Scacchiera.MossaAttuale = NumeroMossa
    AggiornaFinestre 34
  End If
End Sub
Public Sub CambiaOrientamento(Disposizione As OrientamentoScacchiera)
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.Scacchiera.Orientamento = Disposizione
  End If
End Sub
Public Sub SvuotaScacchiera(Optional NonChiediConferma As Boolean)
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.Scacchiera.SvuotaScacchiera NonChiediConferma
    Edita
  End If
End Sub
Public Sub ScacchieraPosizioneIniziale()
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.Scacchiera.PosizioneIniziale
  End If
End Sub
Public Sub CoordinateScacchiera(Valore As Boolean)
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.Scacchiera.VisualizzaCoordinate = Valore
    AggiornaMDIScacchi
  End If
End Sub
Public Sub TipoDiSelezioneScacchiera(Valore As TipoDiSelezione)
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.Scacchiera.TipoDiSelezione = Valore
    AggiornaMDIScacchi
  End If
End Sub
Public Sub AvviaMDI()
  MDIScacchi.Show
End Sub
Public Sub Nuovo(Optional PosizioneIniziale As Boolean, Optional FormGioco As Variant)
  Set FormGioco = New Gioco
  FormGioco.Colloca
  If PosizioneIniziale = True Then
    Editabilit‡Scacchiera BiancoMossePossibili
    FormGioco.Scacchiera.PosizioneIniziale
  End If
  Dim nome As String
  nome = NuovoNome
  FormGioco.Caption = nome
  FormGioco.Tag = LTrim(Str(NuovoNumeroUnivoco))
  FormGioco.NomeFile = nome
  AggiornaFinestre
  ImpostaGiocatore Bianco, Umano
  ImpostaGiocatore Nero, Umano
End Sub
'Restituisce un nome da assegnare a una nuova finestra di Gioco che si intende creare
Private Function NuovoNome() As String
  Dim N As Integer
  Dim Fl As Integer
  Do
    N = N + 1
    NuovoNome = "Situazione n∞ " + CStr(N)
    Fl = False
    Dim f As Variant
    If EsisteGiocoSelezionato(True) = True Then
      For Each f In Forms
         If f.Caption = NuovoNome Then
           Fl = True
           Exit For
         End If
       Next
    End If
    If Fl = False Then
      'il nuovo nome non utilizzato Ë stato trovato
      Exit Function
    End If
  Loop
End Function
'Restituisce un numero univoco da assegnare a una nuova finestra di Gioco che si intende creare
Private Function NuovoNumeroUnivoco() As Integer
  Dim N As Integer
  Dim Fl As Integer
  Do
    N = N + 1
    NuovoNumeroUnivoco = N
    Fl = False
    Dim f As Variant
    If EsisteGiocoSelezionato(True) = True Then
      For Each f In Forms
         If Val(f.Tag) = NuovoNumeroUnivoco Then
           Fl = True
           Exit For
         End If
       Next
    End If
    If Fl = False Then
      'il NuovoNumeroUnivoco nome non utilizzato Ë stato trovato
      Exit Function
    End If
  Loop
End Function
Public Sub MostraAnalisi(Optional formproprietario As Object)
  If EsisteGiocoSelezionato(True) = True Then
    Albero.Show 0, formproprietario
    AggiornaAnalisi
  End If
End Sub
Public Sub MostraOrologio(Optional formproprietario As Object)
  If EsisteGiocoSelezionato(True) = True Then
    Orologio.Show 0, formproprietario
    AggiornaOrologio
  End If
End Sub
Public Sub MostraPensiero(Optional formproprietario As Object)
  If EsisteGiocoSelezionato(True) = True Then
    Pensiero.Show 0, formproprietario
    AggiornaPensiero
  End If
End Sub
Public Sub MostraNota(Optional formproprietario As Object)
  If EsisteGiocoSelezionato(True) = True Then
    Nota.Show 0, formproprietario
    AggiornaNota
  End If
End Sub
Public Sub MostraElencoMosse(Optional formproprietario As Object)
  If EsisteGiocoSelezionato(True) = True Then
    ElencoMosse.Show 0, formproprietario
    AggiornaElencoMosse
  End If
End Sub
Public Sub AggiornaNota()
  If Nota.Visible = True Then
    Nota.Caption = "Nota di " + MDIScacchi.ActiveForm.Caption
    Nota.Text.Text = MDIScacchi.ActiveForm.Scacchiera.Evento
    Nota.TxtAnnotatore = MDIScacchi.ActiveForm.Scacchiera.EventoNomeAnnotatore
    Nota.TxtApertura = MDIScacchi.ActiveForm.Scacchiera.EventoApertura
    Nota.TxtData = CDate(MDIScacchi.ActiveForm.Scacchiera.EventoData)
    Nota.TxtEco = MDIScacchi.ActiveForm.Scacchiera.EventoECO
    Nota.TxtElo(1) = MDIScacchi.ActiveForm.Scacchiera.BiancoElo
    Nota.TxtElo(0) = MDIScacchi.ActiveForm.Scacchiera.NeroElo
    Nota.TxtEmail(1) = MDIScacchi.ActiveForm.Scacchiera.BiancoEM
    Nota.TxtEmail(0) = MDIScacchi.ActiveForm.Scacchiera.NeroEM
    Nota.TxtExtra = MDIScacchi.ActiveForm.Scacchiera.EventoExtra
    Nota.TxtIncontro = MDIScacchi.ActiveForm.Scacchiera.eventoincontro
    Nota.TxtLuogo = MDIScacchi.ActiveForm.Scacchiera.EventoLuogo
    Nota.TxtNazione(1) = MDIScacchi.ActiveForm.Scacchiera.BiancoNazionalit‡
    Nota.TxtNazione(0) = MDIScacchi.ActiveForm.Scacchiera.NeroNazionalit‡
    Nota.TxtNic = MDIScacchi.ActiveForm.Scacchiera.EventoNIC
    Nota.TxtNome(1) = MDIScacchi.ActiveForm.Scacchiera.BiancoGiocatore
    Nota.TxtNome(0) = MDIScacchi.ActiveForm.Scacchiera.NeroGiocatore
    Nota.TxtNumeroScacchiera = MDIScacchi.ActiveForm.Scacchiera.EventoNumeroScacchiera
    Nota.TxtOra = Format(MDIScacchi.ActiveForm.Scacchiera.EventoOra, "HH:MM:SS")
    Nota.TxtSottoApertura = MDIScacchi.ActiveForm.Scacchiera.EventoSottoApertura
    Nota.TxtSottoSottoApertura = MDIScacchi.ActiveForm.Scacchiera.EventoSottoSottoApertura
    Nota.TxtSponsor = MDIScacchi.ActiveForm.Scacchiera.EventoSponsor
    Nota.TxtUscf(1) = MDIScacchi.ActiveForm.Scacchiera.BiancoUSCF
    Nota.TxtUscf(0) = MDIScacchi.ActiveForm.Scacchiera.NeroUSCF
    Nota.CmbModo.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoModo
    Nota.CmbRisultato.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoRisultato
    Nota.CmbTipoGara.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoTipoGara
    Nota.CmbSessione.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoSessione
    Nota.CmbTerminato.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoTerminato
    Nota.CmbTipo(1).ListIndex = MDIScacchi.ActiveForm.Scacchiera.BiancoTipoGiocatore
    Nota.CmbTipo(0).ListIndex = MDIScacchi.ActiveForm.Scacchiera.NeroTipoGiocatore
    Nota.CmbTitolo(1).ListIndex = MDIScacchi.ActiveForm.Scacchiera.BiancoTitolo
    Nota.CmbTitolo(0).ListIndex = MDIScacchi.ActiveForm.Scacchiera.NeroTitolo
    Nota.CmbTurno.ListIndex = MDIScacchi.ActiveForm.Scacchiera.EventoTurno
  End If
End Sub
Public Sub AggiornaAnalisi()
  If Albero.Visible = True Then
    Albero.CreaAlberoDellaPosizione MDIScacchi.ActiveForm.Scacchiera
  End If
End Sub
Public Sub AggiornaOrologio()
  If Orologio.Visible = True Then
    Orologio.Aggiorna MDIScacchi.ActiveForm
  End If
End Sub
Public Sub AggiornaPensiero()
  If Pensiero.Visible = True Then
    Pensiero.Aggiorna MDIScacchi.ActiveForm
  End If
End Sub
Public Sub AggiornaElencoMosse()
  If ElencoMosse.Visible = True Then
    ElencoMosse.Aggiorna
  End If
End Sub
Public Sub AggiornaFinestre(Optional Escludi As Long)
  If EsisteGiocoSelezionato Then
    If (Escludi And 1) = 0 Then
      AggiornaMDIScacchi
    End If
    If (Escludi And 2) = 0 Then
      AggiornaNota
    End If
    If (Escludi And 4) = 0 Then
      AggiornaAnalisi
    End If
    If (Escludi And 8) = 0 Then
      AggiornaOrologio
    End If
    If (Escludi And 16) = 0 Then
      AggiornaElencoMosse
    End If
    If (Escludi And 32) = 0 Then
      AggiornaCheckTipoGiocatore
    End If
    If (Escludi And 64) = 0 Then
      AggiornaPensiero
    End If
  Else
    AggiornaMDIScacchi
    Unload Albero
    Unload Nota
    Unload ElencoMosse
    Unload Orologio
    Unload Pensiero
    AggiornaCheckTipoGiocatore
  End If
End Sub
Public Sub AggiornaEditabilit‡()
  If EsisteGiocoSelezionato Then
    Select Case MDIScacchi.ActiveForm.Scacchiera.Editabilit‡
      Case NonEditabile
        MDIScacchi.mnuMuovi.Visible = False
        MDIScacchi.Editabilit‡.ListIndex = 0
      Case Editabile
        MDIScacchi.mnuMuovi.Visible = False
        MDIScacchi.Editabilit‡.ListIndex = 1
      Case BiancoMossePossibili
        MDIScacchi.mnuMuovi.Visible = True
        MDIScacchi.Editabilit‡.ListIndex = 2
      Case NeroMossePossibili
        MDIScacchi.mnuMuovi.Visible = True
        MDIScacchi.Editabilit‡.ListIndex = 3
    End Select
    MDIScacchi.mnuMossaBianco.Checked = MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = BiancoMossePossibili
    MDIScacchi.mnuMossaNero.Checked = MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = NeroMossePossibili
    If MDIScacchi.ActiveForm.Scacchiera.MossaAttuale < MDIScacchi.ActiveForm.Scacchiera.Quantit‡MosseTraccia Then
      MDIScacchi.Successiva.Enabled = True
    Else
      MDIScacchi.Successiva.Enabled = False
    End If
    If MDIScacchi.ActiveForm.Scacchiera.MossaAttuale > 0 Then
      MDIScacchi.Precedente.Enabled = True
    Else
      MDIScacchi.Precedente.Enabled = False
    End If
  Else
    MDIScacchi.Editabilit‡.ListIndex = 0
    MDIScacchi.mnuMossaBianco.Checked = False
    MDIScacchi.mnuMossaNero.Checked = False
  End If
End Sub
Public Sub AggiornaMDIScacchi()
  If EsisteGiocoSelezionato Then
    Dim N As Integer
    For N = 0 To 3
      MDIScacchi.mnuBianco(N).Checked = False
      MDIScacchi.mnuNero(N).Checked = False
    Next
    Select Case MDIScacchi.ActiveForm.Scacchiera.Orientamento
      Case BiancoNero
        MDIScacchi.mnuBianco(0).Checked = True
        MDIScacchi.mnuNero(1).Checked = True
      Case NeroBianco
        MDIScacchi.mnuBianco(1).Checked = True
        MDIScacchi.mnuNero(0).Checked = True
      Case SinistraDestra
        MDIScacchi.mnuBianco(2).Checked = True
        MDIScacchi.mnuNero(3).Checked = True
      Case DestraSinistra
        MDIScacchi.mnuBianco(3).Checked = True
        MDIScacchi.mnuNero(2).Checked = True
    End Select
    MDIScacchi.mnuCoordinate.Checked = MDIScacchi.ActiveForm.Scacchiera.VisualizzaCoordinate
    MDIScacchi.mnuBordo.Checked = (MDIScacchi.ActiveForm.Scacchiera.TipoDiSelezione And 1) <> 0
    MDIScacchi.mnuEvidenzia.Checked = (MDIScacchi.ActiveForm.Scacchiera.TipoDiSelezione And 2) <> 0
    Dim s As String
    If MDIScacchi.ActiveForm.NomeFile <> "" Then
      s = " "
    Else
      s = ""
    End If
    MDIScacchi.mnuSalvaConNome.Caption = "Salva " + s + MDIScacchi.ActiveForm.Caption + " con no&me"
  End If
  'visuolizza o nasconde le voci di menu
  Dim Valore As Boolean
  If EsisteGiocoSelezionato Then
    Valore = True
  Else
    Valore = False
  End If
  MDIScacchi.mnuSalva.Visible = Valore
  MDIScacchi.mnuSalvaConNome.Visible = Valore
  MDIScacchi.mnuStampa.Visible = Valore
  MDIScacchi.mnuModifica.Visible = Valore
  MDIScacchi.mnuScacchiera.Visible = Valore
  MDIScacchi.mnuStrumenti.Visible = Valore
  MDIScacchi.mnuProgetto.Visible = Valore
  MDIScacchi.mnuFinestra.Visible = Valore
  Dim C As Control
  For Each C In MDIScacchi.Controls
    If (TypeOf C Is CommandButton) Or (TypeOf C Is ComboBox) Then
      If C.Name <> "Apri" And C.Name <> "Nuovo" Then
        C.Enabled = Valore
      End If
    End If
  Next
  For N = 0 To 11
    MDIScacchi.Casa(N).Visible = Valore
  Next
  AggiornaEditabilit‡
End Sub
Public Sub AggiornaCheckTipoGiocatore()
  'Aggiorna i check che indicano il tipo di giocatore (Umano o Computer)
  If EsisteGiocoSelezionato(False) Then
    If MDIScacchi.ActiveForm.Scacchiera.BiancoTipoGiocatore = Umano Then
      MDIScacchi.MnuGBUmano.Checked = True
      MDIScacchi.mnuGBComputer.Checked = False
    Else
      MDIScacchi.MnuGBUmano.Checked = False
      MDIScacchi.mnuGBComputer.Checked = True
    End If
    If MDIScacchi.ActiveForm.Scacchiera.NeroTipoGiocatore = Umano Then
      MDIScacchi.MnuGNUmano.Checked = True
      MDIScacchi.MnuGNComputer.Checked = False
    Else
      MDIScacchi.MnuGNUmano.Checked = False
      MDIScacchi.MnuGNComputer.Checked = True
    End If
  End If
End Sub
Public Sub Salva()
  If EsisteGiocoSelezionato(True) Then
    Dim Value As Integer
    Value = SaveToFile(MDIScacchi.ActiveForm)
  End If
End Sub
Public Sub SalvaConNome()
  If EsisteGiocoSelezionato(True) Then
    Dim Value As Integer
    Value = SaveToFile(MDIScacchi.ActiveForm, True)
  End If
End Sub
Public Sub Apri(Optional File As String)
  Dim Value As Integer
  Value = ReadToFile(, File)
  AggiornaFinestre
End Sub
Public Sub Chiudi()
  If EsisteGiocoSelezionato(False) Then
    Unload MDIScacchi.ActiveForm
  End If
End Sub
Private Function ReadToFile(Optional FinestraGioco As Object, Optional File As String) As Boolean
  On Local Error Resume Next
  If File = Empty Then
    MDIScacchi.CommonDialog.CancelError = True
    MDIScacchi.CommonDialog.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNLongNames + cdlOFNFileMustExist
    MDIScacchi.CommonDialog.Filter = "File di posizione di gioco scacchi (*.Sca)|*.Sca|Archivio Portable Games (*.Pg?)|*.Pg?"
    MDIScacchi.CommonDialog.ShowOpen
    File = MDIScacchi.CommonDialog.FileName
  End If
  If Err = 0 Then
    If UCase(Left(Right(File, 4), 3)) = ".PG" Then
      Archivio.LeggiFileArchivio (File)
      Archivio.Show 0, MDIScacchi
    Else
      If TypeName(FinestraGioco) <> "Gioco" Then
        Nuovo False, FinestraGioco
      End If
      FinestraGioco.Scacchiera.ReadToFile (File)
      Dim P As Integer
      P = InStr(File, ".")
      If P <> 0 Then
        File = Left(File, P - 1)
      End If
      FinestraGioco.NomeFile = File
      FinestraGioco.Caption = NomeFile2Caption(File)
      ReadToFile = True
    End If
  Else
    ReadToFile = False
  End If
End Function
Private Function SaveToFile(FinestraGioco As Object, Optional ConNome As Boolean) As Boolean
  On Local Error Resume Next
  If ConNome = True Or MDIScacchi.ActiveForm.NomeFile = "" Then
    On Local Error Resume Next
    MDIScacchi.CommonDialog.CancelError = True
    MDIScacchi.CommonDialog.Flags = cdlOFNOverwritePrompt + cdlOFNHelpButton + cdlOFNPathMustExist
    MDIScacchi.CommonDialog.Filter = "File di posizione di gioco scacchi (*.Sca)|*.Sca"
    MDIScacchi.CommonDialog.FileName = MDIScacchi.ActiveForm.NomeFile
    MDIScacchi.CommonDialog.ShowSave
    Dim File As String
    File = MDIScacchi.CommonDialog.FileName
  Else
    File = FinestraGioco.NomeFile + ".Sca"
  End If
  If Err = 0 Then
    SaveToFile = True
    FinestraGioco.Scacchiera.SaveToFile (File)
  Else
    SaveToFile = False
  End If
End Function
Public Sub Stampa()
  If EsisteGiocoSelezionato(True) Then
    MDIScacchi.ActiveForm.PrintForm
  End If
End Sub
Public Function EsisteGiocoSelezionato(Optional Avvisa As Boolean) As Integer
  On Local Error Resume Next
  If MDIScacchi.ActiveForm.Name <> "" Then
    If Err.Number <> 0 Then
      Dim respose As Variant
      If Avvisa = True Then
        respose = MsgBox("Non Ë stata specificata nessuna situazione di gioco, per fare ciÚ occore aprirne una esistente o creare una nuova!", vbRetryCancel)
      End If
      If respose = 4 Then
        MsgBox "Il problema verr‡ risolto automaticamente creando un nuovo gioco in cui i pezzi sono disposti nella situazione iniziale!", vbInformation
        Nuovo
        EsisteGiocoSelezionato = True
        Exit Function
      Else
        EsisteGiocoSelezionato = False
        Exit Function
      End If
    Else
      EsisteGiocoSelezionato = True
      Exit Function
    End If
  End If
End Function
Public Function Taglia()
  If EsisteGiocoSelezionato(True) = True Then
    Call Copia
    MDIScacchi.ActiveForm.Scacchiera.SvuotaScacchiera True
  End If
End Function
Public Function Incolla()
  If EsisteGiocoSelezionato(True) = True Then
    If Dir("Clip.tmp") <> "" Then
      MDIScacchi.ActiveForm.Scacchiera.ReadToFile "Clip.tmp"
    Else
      MsgBox "Operazione non riuscita perchË in memoria non esiste nessuna situazione di gioco che possa essere incollata!", vbInformation
    End If
  End If
End Function
Public Function Copia()
  If EsisteGiocoSelezionato(True) = True Then
    Clipboard.Clear
    Clipboard.SetText Elaborazione.FEN(MDIScacchi.ActiveForm.Scacchiera)
    MDIScacchi.ActiveForm.Scacchiera.SaveToFile "Clip.tmp"
  End If
End Function
Public Function Duplica()
  If EsisteGiocoSelezionato(True) = True Then
    Call Copia
    Call Nuovo
    Call Incolla
  End If
End Function
Public Function Edita()
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = Editabile
    AggiornaEditabilit‡
  End If
End Function
Public Function Blocca()
  If EsisteGiocoSelezionato(True) = True Then
    MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = NonEditabile
    AggiornaEditabilit‡
  End If
End Function
Public Sub InserisciPezzo(pezzo As pezzo)
  If EsisteGiocoSelezionato(True) = True Then
    If MDIScacchi.ActiveForm.AspettaSelezionePezzoPromosso = True Then
      If Abs(pezzo) = Dama Or Abs(pezzo) = Torre Or Abs(pezzo) = Alfiere Or Abs(pezzo) = Cavallo Then
        MDIScacchi.ActiveForm.Scacchiera.ImpostaPezzoPromosso Abs(pezzo)
        MDIScacchi.ActiveForm.AspettaSelezionePezzoPromosso = False
        MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = MDIScacchi.ActiveForm.Editabilit‡DurantePromozione
        MDIScacchi.ActiveForm.MossaEffettuata
      End If
    Else
      MDIScacchi.ActiveForm.Scacchiera.InserisciPezzo (pezzo)
    End If
  End If
End Sub
Public Sub Turno(Colore As Colore)
  If EsisteGiocoSelezionato(True) = True Then
    If Colore = Bianco Then
      MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = BiancoMossePossibili
    ElseIf Colore = Nero Then
      MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = NeroMossePossibili
    Else
      MDIScacchi.ActiveForm.Scacchiera.Editabilit‡ = NonEditabile
    End If
    AggiornaEditabilit‡
  End If
End Sub
Private Function NomeFile2Caption(File As String) As String
  If InStr(File, "\") = 0 Then
    NomeFile2Caption = File
  Else
    Dim N As Integer
    Do
      N = InStr(N + 1, File, "\")
      If N <> 0 Then
        NomeFile2Caption = Mid(File, N + 1)
      End If
    Loop Until N = 0
  End If
End Function
Public Sub CaricaFigure(Optional PthPz As String)
  If PthPz <> "" Then
    If Dir(PthPz, 16) <> "" Then
      Dim imgx As Object
      While MDIScacchi.FigureScacchi.ListImages.Count
        MDIScacchi.FigureScacchi.ListImages.Clear
      Wend
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.pedone.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.cavallo.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.alfiere.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.torre.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.dama.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "bianco.re.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.pedone.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.cavallo.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.alfiere.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.torre.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.dama.gif"))
      Set imgx = MDIScacchi.FigureScacchi.ListImages.Add(, , LoadPicture(PthPz & "\" & "nero.re.gif"))
      Dim N As IndexCasa
      Dim P As pezzo
      For N = 0 To 11
        P = MDIScacchi.Casa(N).figura
        MDIScacchi.Casa(N).PathFigure = PthPz
        MDIScacchi.Casa(N).figura = Nessun_Pezzo
        MDIScacchi.Casa(N).figura = P
      Next
      If EsisteGiocoSelezionato = True Then
        Dim f As Variant
        For Each f In Forms
          If f.Name = "Gioco" Then
            f.Scacchiera.RicaricaFigure
          End If
        Next
      End If
      Exit Sub
    End If
  End If
  SelezionaSetPezzi.Show 1
End Sub

Public Sub CaricaPartita(NomeFile As String, Optional SeekRecord As Long = 1, Optional FormGioco As Variant, Optional DisabilitaMSG As Boolean = False)
  If TypeName(FormGioco) = "Nothing" Or TypeName(FormGioco) = "Error" Then
    Nuovo , FormGioco
  End If
  Dim OLDAggAbb As Boolean
  OLDAggAbb = FormGioco.AggiornamentiAbilitati
  FormGioco.AggiornamentiAbilitati = False
  Elaborazione.CaricaPartita NomeFile, SeekRecord, FormGioco.Scacchiera, DisabilitaMSG
  FormGioco.AggiornamentiAbilitati = OLDAggAbb
End Sub

