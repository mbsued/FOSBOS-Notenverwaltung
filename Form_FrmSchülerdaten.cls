VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmSchülerdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private myclsSchülerdaten As clsSchülerdaten        'Schülerdaten
Private booFormularAenderung As Boolean             'Wurde eine Änderung durchgeführt?
Private booKlassengruppeAenderung As Boolean        'Hat sich die Klassengruppe geändert?
Private lngKlassengruppeUidAlt As Long              'uid_Klassengruppe_vor Änderung
Private booBeenden As Boolean

Private Sub cmbgeschlecht_AfterUpdate()
    myclsSchülerdaten.geschlecht = Me.cmbgeschlecht
    booFormularAenderung = True
End Sub

Private Sub cmbklassengruppe_AfterUpdate()
    myclsSchülerdaten.klassengruppe = Me.cmbklassengruppe
    booFormularAenderung = True
    booKlassengruppeAenderung = True
End Sub

Private Sub cmbSchueler_AfterUpdate()
' Prüfen ob vorher Änderung war, wenn ja speichern
' jetzt die entsprechenden Daten lesen

    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If

    Set myclsSchülerdaten = SchülerLesen(Me.cmbSchueler)
' Daten füllen
    SchülerdatenAnzeigen

End Sub

Private Sub cmdAbbrechen_Click()
' Formular schließen vorher nachfragen ob Änderungen verworfen werden sollen

    On Error GoTo Err_cmdclose_Click
    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen verworfen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            booBeenden = True
            DOCMD.Close
        End If
    Else
        booBeenden = True
        DOCMD.Close
    End If

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdSpeichern_Click()
    ' die vorgegebenen Daten speichern (Schülerdaten)
    If CheckAusgeschieden = True Then
        SchülerSchreiben myclsSchülerdaten
        ' Formular geändert zurücksetzen
        booFormularAenderung = False
        If booKlassengruppeAenderung Then
            If CheckReligion = True Then
                KlassengruppeAendern
                booKlassengruppeAenderung = False
            Else
                MsgBox "Das Fach Religion muss bei Änderung der Klassengruppe belegt sein!", vbCritical, "FOSBOS Notenverwaltung"
                Me.txtreligion.SetFocus
            End If
        End If
    Else
        MsgBox "Das Feld Ausgeschieden und das Feld Ausgeschieden am müssen beide belegt sein!", vbCritical, "FOSBOS Notenverwaltung"
        Me.ctlausgeschieden.SetFocus
    End If

End Sub

Private Sub ctlausgeschieden_AfterUpdate()
    myclsSchülerdaten.ausgeschieden = Me.ctlausgeschieden
    booFormularAenderung = True
End Sub

Private Sub ctlnotenschutz_AfterUpdate()
    myclsSchülerdaten.notenschutz = Me.ctlnotenschutz
    booFormularAenderung = True
End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("frmStart") Then
        Forms.Item("frmStart").Visible = True
    End If
End Sub
Private Sub Form_Load()
'Höhe und Breite einstellen
'Formular Änderung auf falsch stellen
'1. Klassengruppe einstellen
    
    booBeenden = False
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
            
    booFormularAenderung = False
    booKlassengruppeAenderung = False
    cmbSchueler_AfterUpdate
    
End Sub

Private Sub cmdClose_Click()
'Schließen vorher nachfragen ob gespeichert werden soll
    On Error GoTo Err_cmdclose_Click
    
    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If
    
    booBeenden = True
    DOCMD.Close
    
Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub
Private Sub SchülerdatenAnzeigen()
    Me.txtnachname = myclsSchülerdaten.nachname
    Me.txtrufname = myclsSchülerdaten.rufname
    Me.txtvornamen = myclsSchülerdaten.vornamen
    Me.cmbgeschlecht = myclsSchülerdaten.geschlecht
    Me.txtgebdat = myclsSchülerdaten.geburtsdatum
    Me.txtgebort = myclsSchülerdaten.geburtsort
    Me.txtgebland = myclsSchülerdaten.geburtsland
    Me.txtbekenntnis = myclsSchülerdaten.bekenntnis
    Me.txtreligion = myclsSchülerdaten.religion
    Me.ctlausgeschieden = myclsSchülerdaten.ausgeschieden
    Me.txtausgeschiedenam = myclsSchülerdaten.ausgeschiedenam
    Me.cmbklassengruppe = myclsSchülerdaten.klassengruppe
    lngKlassengruppeUidAlt = Me.cmbklassengruppe
    Me.txteingetreten = myclsSchülerdaten.eintritt
    Me.txtprobezeit = myclsSchülerdaten.probezeit
    Me.ctlnotenschutz = myclsSchülerdaten.notenschutz
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltflächen Formular schließen oder Abbrechen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If

End Sub

Private Sub txtausgeschiedenam_AfterUpdate()
    myclsSchülerdaten.ausgeschiedenam = Me.txtausgeschiedenam
    booFormularAenderung = True
End Sub

Private Sub txtbekenntnis_AfterUpdate()
    myclsSchülerdaten.bekenntnis = Me.txtbekenntnis
    booFormularAenderung = True
End Sub

Private Sub txteingetreten_AfterUpdate()
    myclsSchülerdaten.eintritt = Me.txteingetreten
    booFormularAenderung = True
End Sub

Private Sub txtgebdat_AfterUpdate()
    myclsSchülerdaten.geburtsdatum = Me.txtgebdat
    booFormularAenderung = True
End Sub

Private Sub txtgebland_AfterUpdate()
    myclsSchülerdaten.geburtsland = Me.txtgebland
    booFormularAenderung = True
End Sub

Private Sub txtgebort_AfterUpdate()
    myclsSchülerdaten.geburtsort = Me.txtgebort
    booFormularAenderung = True
End Sub

Private Sub txtnachname_AfterUpdate()
    myclsSchülerdaten.nachname = Me.txtnachname
    booFormularAenderung = True
End Sub

Private Sub txtprobezeit_AfterUpdate()
    myclsSchülerdaten.probezeit = Me.txtprobezeit
    booFormularAenderung = True
End Sub

Private Sub txtreligion_AfterUpdate()
    myclsSchülerdaten.religion = Me.txtreligion
    booFormularAenderung = True
End Sub

Private Sub txtrufname_AfterUpdate()
    myclsSchülerdaten.rufname = Me.txtrufname
    booFormularAenderung = True
End Sub

Private Sub txtvornamen_AfterUpdate()
    myclsSchülerdaten.vornamen = Me.txtvornamen
    booFormularAenderung = True
End Sub

Private Function CheckAusgeschieden() As Boolean
' Plausibilisierung Ausgeschieden
    If Me.ctlausgeschieden = True Then
        If IsNull(Me.txtausgeschiedenam) Or Me.txtausgeschiedenam = "" Then
            CheckAusgeschieden = False
            Exit Function
        Else
            CheckAusgeschieden = True
        End If
    Else
        If Not IsNull(Me.txtausgeschiedenam) And Not Me.txtausgeschiedenam = "" Then
            CheckAusgeschieden = False
            Exit Function
        Else
            CheckAusgeschieden = True
        End If
    End If
End Function

Private Function CheckReligion() As Boolean
' Plausibilisierung Religion, wenn Änderung Klassengruppe
Dim lngJahrgangsstufeUid As Long
Dim mydb As clsDB

    Set mydb = New clsDB
    lngJahrgangsstufeUid = mydb.JahrgangsstufeErmitteln(Me.cmbklassengruppe)
    Set mydb = Nothing
    
    If lngJahrgangsstufeUid <> 2 And (Me.txtreligion = "" Or IsNull(Me.txtreligion)) Then
        CheckReligion = False
    Else
        If Me.txtreligion = "RK" Or Me.txtreligion = "EV" Or Me.txtreligion = "ETH" Then
            CheckReligion = True
        Else
            CheckReligion = False
        End If
    End If

End Function

Private Sub KlassengruppeAendern()
' Alle Notensätze des Schülers löschen und für die neue Klassengruppe neuanlegen
Dim mydb As clsDB
Dim mycolSchuelerNotenKurz As colSchülerNotenKurz
Dim mycolSchuelerHalbjahrKurz As colSchülerHalbjahrKurz
Dim intCount As Integer
Dim lngJahrgangsstufeNeuUid As Long
Dim lngJahrgangsstufeAltUid As Long
Dim lngAusbNeuUid As Long
Dim lngAusbAltUid As Long

' Jahrgangsstufe und Ausbildungsrichtung neu und alt ermitteln
    Set mydb = New clsDB
    lngJahrgangsstufeNeuUid = mydb.JahrgangsstufeErmitteln(Me.cmbklassengruppe)
    lngJahrgangsstufeAltUid = mydb.JahrgangsstufeErmitteln(lngKlassengruppeUidAlt)
    lngAusbNeuUid = mydb.AusbildungsrichtungErmitteln(Me.cmbklassengruppe)
    lngAusbAltUid = mydb.AusbildungsrichtungErmitteln(lngKlassengruppeUidAlt)
    Set mydb = Nothing
    
' Prüfen, ob sich die Jahrgangsstufe und/oder die Ausbildungsrichtung geändert haben,
' dann entsprechend entweder nur die Klassengruppe bei den Notensätzen ändern oder die Notensätze löschen und neuanlegen

    If lngJahrgangsstufeAltUid <> lngJahrgangsstufeNeuUid Or (lngJahrgangsstufeAltUid = lngJahrgangsstufeNeuUid And lngAusbAltUid <> lngAusbNeuUid) Then
' zuerst einmal die alten Sätze löschen
        Set mydb = New clsDB
    
        Set mycolSchuelerNotenKurz = SchuelerNotenLesenKurz(myclsSchülerdaten.uid)
        For intCount = 1 To mycolSchuelerNotenKurz.Count
            mycolSchuelerNotenKurz.Item(intCount).geloescht = True
        Next intCount
    
        Set mydb.SchuelerNotenKurz = mycolSchuelerNotenKurz
        mydb.SchuelerNotenSpeichernKurz
    
        Set mycolSchuelerHalbjahrKurz = SchuelerHalbjahresNotenLesenKurz(myclsSchülerdaten.uid)
        For intCount = 1 To mycolSchuelerHalbjahrKurz.Count
            mycolSchuelerHalbjahrKurz.Item(intCount).geloescht = True
        Next intCount
    
        Set mydb.SchuelerHalbjahrKurz = mycolSchuelerHalbjahrKurz
        
        mydb.SchuelerHalbjahrSpeichernKurz
        Set mycolSchuelerHalbjahrKurz = Nothing

'jetzt die neuen Sätze anlegen
        mydb.SchuelerNotensaetzeNachKlassengruppenWechselAnlegen Me.cmbklassengruppe, myclsSchülerdaten.uid, lngJahrgangsstufeNeuUid, Me.txtreligion, Me.txtrufname, Me.txtnachname
' sind noch Schüler in der alten Klassengruppe?
' wenn nein, dann Klassengruppe  und NotenfachKlassengruppe löschen und ggf. die Klasse
        If mydb.SchuelerAnzahleinerKlassengruppeErmitteln(lngKlassengruppeUidAlt) = 0 Then
            mydb.KlassengruppeOhneSchuelerLoeschen (lngKlassengruppeUidAlt)
        End If

        Set mydb = Nothing
        
    Else
        Set mydb = New clsDB
    
        Set mycolSchuelerNotenKurz = SchuelerNotenLesenKurz(myclsSchülerdaten.uid)
        For intCount = 1 To mycolSchuelerNotenKurz.Count
            mycolSchuelerNotenKurz.Item(intCount).klassengruppe_uid = Me.cmbklassengruppe
        Next intCount
    
        Set mydb.SchuelerNotenKurz = mycolSchuelerNotenKurz
        mydb.SchuelerNotenSpeichernKurz
    
        Set mycolSchuelerNotenKurz = Nothing
    
        Set mycolSchuelerHalbjahrKurz = SchuelerHalbjahresNotenLesenKurz(myclsSchülerdaten.uid)
        For intCount = 1 To mycolSchuelerHalbjahrKurz.Count
            mycolSchuelerHalbjahrKurz.Item(intCount).klassengruppe_uid = Me.cmbklassengruppe
        Next intCount
    
        Set mydb.SchuelerHalbjahrKurz = mycolSchuelerHalbjahrKurz
        
        mydb.SchuelerHalbjahrSpeichernKurz
        Set mycolSchuelerHalbjahrKurz = Nothing

        Set mydb = Nothing
    End If
End Sub
