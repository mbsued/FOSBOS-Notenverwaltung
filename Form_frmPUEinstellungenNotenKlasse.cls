VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPUEinstellungenNotenKlasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim strSpalte As String
Dim strHalbjahr As String
Dim lngArt As Long
Dim strGewichtung As String
Dim strDatum As String
Dim lngFach As Long

Private Sub cmdAbbrechen_Click()
'Schließen ohne Speichern
    On Error GoTo Err_cmdclose_Click

    DOCMD.Close acForm, "frmPUEinstellungenNotenKlasse"

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdSpeichern_Click()

    On Error GoTo Err_cmdSpeichern_Click
    
    If IsNull(Me.cmbGewichtung) Then
        MsgBox "Bitte die Gewichtung eingeben!", vbOKOnly, "Notenprogramm FOSBOS"
        GoTo Exit_cmdSpeichern_Click
    End If
    
    If IsNull(Me.cmbArt) Then
        MsgBox "Bitte die Art der Leistung eingeben!", vbOKOnly, "Notenprogramm FOSBOS"
        GoTo Exit_cmdSpeichern_Click
    End If
' gib die Einstellungen in den globalen Variablen zurück
    Forms("FrmNotenNachKlassen").RueckgabeAenderungUebernehmen = True
    Forms("FrmNotenNachKlassen").RueckgabeKlasseUidArt = Me.cmbArt
    Forms("FrmNotenNachKlassen").RueckgabeKlasseGewichtung = Me.cmbGewichtung
    Forms("FrmNotenNachKlassen").RueckgabeKlasseDatum = Nz(Me.txtDatum, "")
    
    cmdAbbrechen_Click
    
Exit_cmdSpeichern_Click:
    Exit Sub

Err_cmdSpeichern_Click:
    FehlermeldungAusgeben "Formular Einstellungen Noten", Err.Source, Err.Number, Err.Description
    Resume Exit_cmdSpeichern_Click
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strValue As Variant
    
    On Error GoTo Err_Form_Open
    
    If Not IsNull(Me.OpenArgs) Then
        strValue = Split(Me.OpenArgs, ";", , vbBinaryCompare)
        strSpalte = strValue(0)
        strHalbjahr = strValue(1)
        lngArt = strValue(2)
        strGewichtung = strValue(3)
        strDatum = strValue(4)
        lngFach = strValue(5)
        
        Me.txtSpalte = strSpalte
        Me.txtHalbjahr = strHalbjahr
        Me.Caption = "Einstellungen für sonstige Leistung " & strSpalte & " im " & strHalbjahr & ". Hj"
        If lngFach <> 35 Then   'uid 35 = fachpraktische Ausbildung
            Me.cmbArt.RowSource = "SELECT [TabArtLeistungsnachweise].[uid], [TabArtLeistungsnachweise].[bezeichnung_lang] FROM TabArtLeistungsnachweise " & _
                                    "WHERE [TabArtLeistungsnachweise].[uid] < 5 ORDER BY [bezeichnung_lang];"
            Me.cmbArt.RowSourceType = "Tabelle/Abfrage"
        Else
            Me.cmbArt.RowSource = "SELECT [TabArtLeistungsnachweise].[uid], [TabArtLeistungsnachweise].[bezeichnung_lang] FROM TabArtLeistungsnachweise " & _
                                    "WHERE [TabArtLeistungsnachweise].[uid] > 4 ORDER BY [bezeichnung_lang];"
            Me.cmbArt.RowSourceType = "Tabelle/Abfrage"
            Me.cmbArt.Enabled = False
            Me.cmbGewichtung.Enabled = False
        End If
        If lngArt > 0 Then
            Me.cmbArt = lngArt
        End If
        If Not IsNull(strGewichtung) And Not strGewichtung = "" Then
            Me.cmbGewichtung = strGewichtung
        End If
        If Not IsNull(strDatum) And Not strDatum = "" Then
            Me.txtDatum = strDatum
        End If
    End If
Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    FehlermeldungAusgeben "Formular Einstellungen Noten", Err.Source, Err.Number, Err.Description
    Resume Exit_Form_Open
    
End Sub
