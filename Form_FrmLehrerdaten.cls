VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmLehrerdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private myclsLehrerdaten As clsLehrerDaten          'Lehrerdaten
Private booFormularAenderung As Boolean             'Wurde eine Änderung durchgeführt?
Private booBeenden As Boolean

Private Sub cmbgeschlecht_AfterUpdate()
    myclsLehrerdaten.geschlecht = Me.cmbgeschlecht
    booFormularAenderung = True
End Sub

Private Sub cmbLehrer_AfterUpdate()
' Prüfen ob vorher Änderung war, wenn ja speichern
' Daten der Fächer einlesen
' entsprechende Felder einblenden
' jetzt die entsprechenden Daten lesen

    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If

    Set myclsLehrerdaten = LehrerLesen(Me.cmbLehrer)
' Daten füllen
    LehrerdatenAnzeigen

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
' die vorgegebenen Daten speichern (Lehrerdaten)
    LehrerSchreiben myclsLehrerdaten
    ' Formular geändert zurücksetzen
    booFormularAenderung = False

End Sub

Private Sub ctlschulleitung_AfterUpdate()
    myclsLehrerdaten.schulleitung = Me.ctlschulleitung
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
    cmbLehrer_AfterUpdate
    
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
Private Sub LehrerdatenAnzeigen()
    Me.txtkuerzel = myclsLehrerdaten.kuerzel
    Me.txtnachname = myclsLehrerdaten.nachname
    Me.txtrufname = myclsLehrerdaten.rufname
    Me.txtamt = myclsLehrerdaten.amt
    Me.txttitel = myclsLehrerdaten.titel
    Me.txtemail = myclsLehrerdaten.email
    Me.cmbgeschlecht = myclsLehrerdaten.geschlecht
    Me.ctlschulleitung = myclsLehrerdaten.schulleitung
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltflächen Formular schließen oder Abbrechen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If
End Sub

Private Sub txtamt_AfterUpdate()
    myclsLehrerdaten.amt = Me.txtamt
    booFormularAenderung = True
End Sub

Private Sub txtemail_AfterUpdate()
    myclsLehrerdaten.email = Me.txtemail
    booFormularAenderung = True
End Sub

Private Sub txtkuerzel_AfterUpdate()
    If Not IsNull(Me.txtkuerzel) And Not Me.txtkuerzel = "" Then
        myclsLehrerdaten.kuerzel = Me.txtkuerzel
        booFormularAenderung = True
    Else
    End If
End Sub

Private Sub txtnachname_AfterUpdate()
    myclsLehrerdaten.nachname = Me.txtnachname
    booFormularAenderung = True
End Sub

Private Sub txtrufname_AfterUpdate()
    myclsLehrerdaten.rufname = Me.txtrufname
    booFormularAenderung = True
End Sub

Private Sub txttitel_AfterUpdate()
    myclsLehrerdaten.titel = Me.txttitel
    booFormularAenderung = True

End Sub
