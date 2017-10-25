VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmHalbjahresleistungenSchüler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mycolFaecherhalbjahr As colFaecherHalbjahr 'Sammlung der Halbjahresleistungen der Fächer eines Schülers
Private booBeenden As Boolean

Private Sub JahresnoteBerechnen()
' Berechne die Jahresnoten der einzelnen Fächer
Dim dblsumme As Double              'Summe der Halbjahre
Dim dblErgebnis As Double           'Ergebnis

Dim intCountFaecher As Integer      'Zähler für Fächer
Dim strControl1 As String           'Steuerelement Halbjahr 1
Dim strControl2 As String           'Steuerelement Halbjahr 2
Dim StrControl3 As String           'Steuerelement Jahresnote
    
    On Error GoTo Err_JahresnoteBerechnen
    
    For intCountFaecher = 1 To mycolFaecherhalbjahr.Count
        If Format(Date, "m") < 4 Then       'nur Halbjahresleistung 1 zur Berechnung hernehmen
            If mycolFaecherhalbjahr.Item(intCountFaecher).jahrgangsstufe_uid <> 3 Or (mycolFaecherhalbjahr.Item(intCountFaecher).jahrgangsstufe_uid And mycolFaecherhalbjahr.Item(intCountFaecher).schulart_uid = 2) Then
                strControl1 = "txt_hj1_1_" & intCountFaecher
            Else
                strControl1 = "txt_hj1_2_" & intCountFaecher
            End If
            StrControl3 = "txt_jn_" & intCountFaecher
            If Not Me.Controls(strControl1) = "-" Then
                Me.Controls(StrControl3) = Me.Controls(strControl1)
            Else
                Me.Controls(StrControl3) = "-"
            End If
        Else        ' mit beiden Halbjahren rechnen
            If mycolFaecherhalbjahr.Item(intCountFaecher).jahrgangsstufe_uid <> 3 Or (mycolFaecherhalbjahr.Item(intCountFaecher).jahrgangsstufe_uid And mycolFaecherhalbjahr.Item(intCountFaecher).schulart_uid = 2) Then
                strControl1 = "txt_hj1_1_" & intCountFaecher
                strControl2 = "txt_hj2_1_" & intCountFaecher
            Else
                strControl1 = "txt_hj1_2_" & intCountFaecher
                strControl2 = "txt_hj2_2_" & intCountFaecher
            End If
            StrControl3 = "txt_jn_" & intCountFaecher
            If Not Me.Controls(strControl1) = "-" And Not Me.Controls(strControl2) = "-" Then
                dblsumme = CDbl(Me.Controls(strControl1)) + CDbl(Me.Controls(strControl2))
                dblErgebnis = dblsumme / 2
                If dblErgebnis < 1 Then
                    Me.Controls(StrControl3) = 0
                Else
                    Me.Controls(StrControl3) = Runden(dblErgebnis, 0)
                End If
            Else
                Me.Controls(StrControl3) = "-"
            End If
        End If
    Next intCountFaecher
    
Exit_JahresnoteBerechnen:
    Exit Sub
    
Err_JahresnoteBerechnen:
    FehlermeldungAusgeben "Formular Halbjahresleistungen", Err.Source, Err.Number, Err.Description
    Resume Exit_JahresnoteBerechnen
    
End Sub
Private Sub cmdClose_Click()
'Schließen vorher nachfragen ob gespeichert werden soll
    On Error GoTo Err_cmdclose_Click
    booBeenden = True
    DOCMD.Close
    
Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdNotenblatt_Click()
' Öffne das Notenblatt des Schülers
' Übergabe von Art= 2, SchülerUid und JahrgangstufenUid
    
    DOCMD.OpenReport "rptNotenbogen", acViewPreview, , , acWindowNormal, "2;" & Me.lstSchueler.Column(0) & " ;" & mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid

End Sub

Private Sub cmdprobezeit_Click()
' prüfen, ob Probezeit bestanden ist
    Select Case BestehenPruefen(1, mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid, mycolFaecherhalbjahr)
        Case -1
            MsgBox "Die Halbjahresnoten sind nicht vollständig!", vbOKOnly, "FOSBOS Notenverwaltung"
        Case 0
            Me.lblbestanden.Caption = "Der Schüler hat die Probezeit nicht bestanden."
            Me.lblbestanden.BackColor = RGB(255, 0, 0)
            Me.lblbestanden.ForeColor = RGB(255, 255, 255)
            Me.lblbestanden.Visible = True
        Case 1
            Me.lblbestanden.Caption = "Der Schüler hat die Probezeit bestanden."
            Me.lblbestanden.BackColor = RGB(0, 255, 0)
            Me.lblbestanden.ForeColor = RGB(0, 0, 0)
            Me.lblbestanden.Visible = True
    End Select
    
End Sub

Private Sub cmdschuljahr_Click()
' prüfen, ob das Schuljahr bestanden ist
    Select Case BestehenPruefen(2, mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid, mycolFaecherhalbjahr)
        Case -1
            MsgBox "Die Halbjahresnoten sind nicht vollständig!", vbOKOnly, "FOSBOS Notenverwaltung"
        Case 0
            Me.lblbestanden.Caption = "Der Schüler hat das Schuljahr nicht bestanden."
            Me.lblbestanden.BackColor = RGB(255, 0, 0)
            Me.lblbestanden.ForeColor = RGB(255, 255, 255)
            Me.lblbestanden.Visible = True
        Case 1
            Me.lblbestanden.Caption = "Der Schüler hat das Schuljahr bestanden."
            Me.lblbestanden.BackColor = RGB(0, 255, 0)
            Me.lblbestanden.ForeColor = RGB(0, 0, 0)
            Me.lblbestanden.Visible = True
    End Select
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
'Formular Änderung auf falsch stellen
'1. Klassengruppe und 1.Fach einstellen
    booBeenden = False
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
    Me.lstKlasse = Me.lstKlasse.ItemData(0)
    lstSchueler_AfterUpdate
    
End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("frmStart") Then
        Forms.Item("frmStart").Visible = True
    End If
End Sub


Private Sub lstKlasse_AfterUpdate()
    Me.lstSchueler.Requery
    Me.lstSchueler = Me.lstSchueler.ItemData(0)
    lstSchueler_AfterUpdate
End Sub

Private Sub lstSchueler_AfterUpdate()
lstschueler_Start:

' jetzt die entsprechenden Daten lesen
    Set mycolFaecherhalbjahr = FaecherHalbjahresNotenLesen(Me.lstSchueler, Me.lstKlasse)
' keine Daten gefunden
    If mycolFaecherhalbjahr Is Nothing Then
        MsgBox "In dieser Klassengruppe belegt kein Schüler das Fach " & lstSchueler.Column(1) & "." & vbCrLf & "Es wird das erste Fach in der Liste ausgewählt", vbOKOnly, "FOSBOS Notenverwaltung"
        lstSchueler.Value = Me.lstSchueler.ItemData(0)
        GoTo lstschueler_Start
    End If
' erst mal wieder alles einblenden anschließend ausblenden
    AlleZeilenUndNotenfelderEinblenden
    ZeilenAusblenden
' nicht benötigte Spalten ausblenden
    NotenfelderAusblenden
    SpaltenUeberschriften

' Daten füllen
    SchuelerNameSchreiben
    FaecherNameSchreiben
    NotenfelderFuellen
    JahresnoteBerechnen
End Sub

Private Sub ZeilenAusblenden()
'Alle Zeilen ausblenden die nicht benötigt werden
Dim intCount As Integer         'allgemeiner Zähler
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim strcontrol4 As String
Dim strControl5 As String
Dim strControl6 As String
Dim strControl7 As String
Dim strcontrol8 As String
Dim strControl9 As String
Dim strControl10 As String

    For intCount = mycolFaecherhalbjahr.Count + 1 To 17
        strControl1 = "lbl_fach_" & intCount
        Me.Controls(strControl1).Visible = False
        strControl2 = "txt_hj1_1_" & intCount
        Me.Controls(strControl2).Visible = False
        StrControl3 = "ctl_hj1_1_" & intCount
        Me.Controls(StrControl3).Visible = False
        strcontrol4 = "txt_hj2_1_" & intCount
        Me.Controls(strcontrol4).Visible = False
        strControl5 = "ctl_hj2_1_" & intCount
        Me.Controls(strControl5).Visible = False
        strControl6 = "txt_hj1_2_" & intCount
        Me.Controls(strControl6).Visible = False
        strControl7 = "ctl_hj1_2_" & intCount
        Me.Controls(strControl7).Visible = False
        strcontrol8 = "txt_hj2_2_" & intCount
        Me.Controls(strcontrol8).Visible = False
        strControl9 = "ctl_hj2_2_" & intCount
        Me.Controls(strControl9).Visible = False
        strControl10 = "txt_jn_" & intCount
        Me.Controls(strControl10).Visible = False
    Next intCount

    Me.lblbestanden.Visible = False

End Sub
Private Sub SchuelerNameSchreiben()
'Schreibe den Schülernamen in die Kopfzeile
    Me.lblSchueler.Caption = "Halbjahresergebnisse - " & Me.lstSchueler.Column(1) & " " & Me.lstSchueler.Column(2)

End Sub
Private Sub FaecherNameSchreiben()
'Schreibe den Fachnamen in die jeweilige Zeile
Dim intCount As Integer         'allgemeiner Zähler
Dim strControl1 As String

    For intCount = 1 To mycolFaecherhalbjahr.Count
        strControl1 = "lbl_fach_" & intCount
        Me.Controls(strControl1).Caption = mycolFaecherhalbjahr.Item(intCount).fach
    Next intCount


End Sub
Private Sub NotenfelderFuellen()
'Schreibe die Halbjahresleistungen in die jeweiligen Felder abhängig von der Jahrgangsstufe
Dim intCount As Integer     'allgemeiner Zähler
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim strcontrol4 As String
Dim strControl5 As String
Dim strControl6 As String
Dim strControl7 As String
Dim strcontrol8 As String

    Select Case mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid
        Case 1, 5, 6, 7  'Vorklasse und Vorkurs
            For intCount = 1 To mycolFaecherhalbjahr.Count
                strControl1 = "txt_hj1_1_" & intCount
                strControl2 = "txt_hj2_1_" & intCount
                Me.Controls(strControl1).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1)
                Me.Controls(strControl2).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_vkl_hj2)
            Next intCount
        Case 2  '11. Jahrgangsstufe
            For intCount = 1 To mycolFaecherhalbjahr.Count
                strControl1 = "txt_hj1_1_" & intCount
                strControl2 = "txt_hj2_1_" & intCount
                strControl5 = "ctl_hj1_1_" & intCount
                strControl6 = "ctl_hj2_1_" & intCount
                Me.Controls(strControl1).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_11_hj1)
                Me.Controls(strControl2).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_11_hj2)
                Me.Controls(strControl5).Value = mycolFaecherhalbjahr.Item(intCount).ein_11_hj1
                Me.Controls(strControl6).Value = mycolFaecherhalbjahr.Item(intCount).ein_11_hj2
            Next intCount
        Case 3  '12. Jahrgangsstufe
            If mycolFaecherhalbjahr.Item(1).schulart_uid = 1 Then   'FOS
                For intCount = 1 To mycolFaecherhalbjahr.Count
                    strControl1 = "txt_hj1_1_" & intCount
                    strControl2 = "txt_hj2_1_" & intCount
                    StrControl3 = "txt_hj1_2_" & intCount
                    strcontrol4 = "txt_hj2_2_" & intCount
                    strControl5 = "ctl_hj1_1_" & intCount
                    strControl6 = "ctl_hj2_1_" & intCount
                    strControl7 = "ctl_hj1_2_" & intCount
                    strcontrol8 = "ctl_hj2_2_" & intCount
                    Me.Controls(strControl1).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_11_hj1)
                    Me.Controls(strControl2).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_11_hj2)
                    Me.Controls(StrControl3).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_12_hj1)
                    Me.Controls(strcontrol4).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_12_hj2)
                    Me.Controls(strControl5).Value = mycolFaecherhalbjahr.Item(intCount).ein_11_hj1
                    Me.Controls(strControl6).Value = mycolFaecherhalbjahr.Item(intCount).ein_11_hj2
                    Me.Controls(strControl7).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj1
                    Me.Controls(strcontrol8).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj2
                Next intCount
            Else
                For intCount = 1 To mycolFaecherhalbjahr.Count
                    strControl1 = "txt_hj1_1_" & intCount
                    strControl2 = "txt_hj2_1_" & intCount
                    strControl5 = "ctl_hj1_1_" & intCount
                    strControl6 = "ctl_hj2_1_" & intCount
                    Me.Controls(strControl1).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_12_hj1)
                    Me.Controls(strControl2).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_12_hj2)
                    Me.Controls(strControl5).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj1
                    Me.Controls(strControl6).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj2
                Next intCount
            End If
        Case 4  '13. Jahrgangsstufe
            For intCount = 1 To mycolFaecherhalbjahr.Count
                strControl1 = "txt_hj1_1_" & intCount
                strControl2 = "txt_hj2_1_" & intCount
                strControl5 = "ctl_hj1_1_" & intCount
                strControl6 = "ctl_hj2_1_" & intCount
                Me.Controls(strControl1).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_13_hj1)
                Me.Controls(strControl2).Value = HalbjahresleistungUebertragen(1, mycolFaecherhalbjahr.Item(intCount).n_13_hj2)
                Me.Controls(strControl5).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj1
                Me.Controls(strControl6).Value = mycolFaecherhalbjahr.Item(intCount).ein_12_hj2
            Next intCount
    End Select
    

End Sub

Private Sub NotenfelderAusblenden()
'Blende Notenfelder der Schüler aus
Dim intCount As Integer
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim strcontrol4 As String
' es handelt sich nicht um die Jahrgangsstufe 12 oder um Jahrgangsstufe 12 in der BOS , also 2.Block ausblenden
    If mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid <> 3 Or (mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid = 3 And mycolFaecherhalbjahr.Item(1).schulart_uid = 2) Then
        For intCount = 1 To mycolFaecherhalbjahr.Count
            strControl1 = "txt_hj1_2_" & intCount
            strControl2 = "ctl_hj1_2_" & intCount
            StrControl3 = "txt_hj2_2_" & intCount
            strcontrol4 = "ctl_hj2_2_" & intCount
            Me.Controls(strControl1).Visible = False
            Me.Controls(strControl2).Visible = False
            Me.Controls(StrControl3).Visible = False
            Me.Controls(strcontrol4).Visible = False
        Next intCount
    End If
' bei Vorklasse keine Auswahlkästchen einblenden
    If mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid = 1 Then
        For intCount = 1 To mycolFaecherhalbjahr.Count
            strControl2 = "ctl_hj1_1_" & intCount
            strcontrol4 = "ctl_hj2_1_" & intCount
            Me.Controls(strControl2).Visible = False
            Me.Controls(strcontrol4).Visible = False
        Next intCount
    End If
End Sub
Private Sub SpaltenUeberschriften()
' Fülle die Spaltenüberschriften je nach Jahrgangsstufe
    Select Case mycolFaecherhalbjahr.Item(1).jahrgangsstufe_uid
        Case 1  'Vorklasse
            Me.lbl_hj1_1.Visible = True
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "Vkl/1"
            Me.lbl_hj2_1.Caption = "Vkl/2"
        Case 2  '11. Jahrgangstufe
            Me.lbl_hj1_1.Visible = True
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "11/1"
            Me.lbl_hj2_1.Caption = "11/2"
        Case 3  '12. Jahrgangsstufe
            If mycolFaecherhalbjahr.Item(1).schulart_uid = 1 Then   'Schulart FOS
                Me.lbl_hj1_1.Visible = True
                Me.lbl_hj2_1.Visible = True
                Me.lbl_hj1_2.Visible = True
                Me.lbl_hj2_2.Visible = True
                Me.lbl_hj1_1.Caption = "11/1"
                Me.lbl_hj2_1.Caption = "11/2"
                Me.lbl_hj1_2.Caption = "12/1"
                Me.lbl_hj2_2.Caption = "12/2"
            Else                                                    'Schulart BOS
                Me.lbl_hj1_1.Visible = True
                Me.lbl_hj2_1.Visible = True
                Me.lbl_hj1_2.Visible = False
                Me.lbl_hj2_2.Visible = False
                Me.lbl_hj1_1.Caption = "12/1"
                Me.lbl_hj2_1.Caption = "12/2"
            End If
        Case 4  '13. Jahrgangsstufe
            Me.lbl_hj1_1.Visible = True
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "13/1"
            Me.lbl_hj2_1.Caption = "13/2"
        Case 5  'Vorkurs ganzjährig
            Me.lbl_hj1_1.Visible = True
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "VKU/1"
            Me.lbl_hj2_1.Caption = "VKU/2"
        Case 6  'Vorkurs halbjährig
            Me.lbl_hj1_1.Visible = False
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "VKUh"
        Case 7  'FOS Integrationsvorklasse
            Me.lbl_hj1_1.Visible = False
            Me.lbl_hj2_1.Visible = True
            Me.lbl_hj1_2.Visible = False
            Me.lbl_hj2_2.Visible = False
            Me.lbl_hj1_1.Caption = "VFI/1"
            Me.lbl_hj1_2.Caption = "VFI/2"
    End Select
End Sub
Private Sub AlleZeilenUndNotenfelderEinblenden()
'Blende alles ein
Dim intCount As Integer     'allgemeiner Zähler
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim strcontrol4 As String
Dim strControl5 As String
Dim strControl6 As String
Dim strControl7 As String
Dim strcontrol8 As String
Dim strControl9 As String
Dim strControl10 As String

    For intCount = 1 To 17
        strControl1 = "lbl_fach_" & intCount
        Me.Controls(strControl1).Visible = True
        strControl2 = "txt_hj1_1_" & intCount
        Me.Controls(strControl2).Visible = True
        StrControl3 = "ctl_hj1_1_" & intCount
        Me.Controls(StrControl3).Visible = True
        strcontrol4 = "txt_hj2_1_" & intCount
        Me.Controls(strcontrol4).Visible = True
        strControl5 = "ctl_hj2_1_" & intCount
        Me.Controls(strControl5).Visible = True
        strControl6 = "txt_hj1_2_" & intCount
        Me.Controls(strControl6).Visible = True
        strControl7 = "ctl_hj1_2_" & intCount
        Me.Controls(strControl7).Visible = True
        strcontrol8 = "txt_hj2_2_" & intCount
        Me.Controls(strcontrol8).Visible = True
        strControl9 = "ctl_hj2_2_" & intCount
        Me.Controls(strControl9).Visible = True
        strControl10 = "txt_jn_" & intCount
        Me.Controls(strControl10).Visible = True
    Next intCount
    
    Me.lblbestanden.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltfläche Formular schließen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If
End Sub

