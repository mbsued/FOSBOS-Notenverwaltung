VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmKlassendaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private myclsKlassendaten As clsKlasseDaten         'Klassendaten
Private booFormularAenderung As Boolean             'Wurde eine Änderung durchgeführt?
Private booBeenden As Boolean

Private Sub cmbKlasse_AfterUpdate()
' Prüfen ob vorher Änderung war, wenn ja speichern

    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If

    Set myclsKlassendaten = KlasseLesen(Me.cmbKlasse)
' Daten füllen
    KlassendatenAnzeigen

End Sub

Private Sub cmbklassleitung_AfterUpdate()
    myclsKlassendaten.klassleitung = Me.cmbklassleitung
    booFormularAenderung = True
End Sub

Private Sub cmbzeugnis_AfterUpdate()
    myclsKlassendaten.zeugnisunterzeichner = Me.cmbzeugnis
    booFormularAenderung = True
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
' die vorgegebenen Daten speichern (Klassendaten)
    KlasseSchreiben myclsKlassendaten
    ' Formular geändert zurücksetzen
    booFormularAenderung = False

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
    cmbKlasse_AfterUpdate
    
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
Private Sub KlassendatenAnzeigen()
    Me.txtbezeichnung = myclsKlassendaten.bezeichnung
    Me.cmbklassleitung = myclsKlassendaten.klassleitung
    Me.cmbzeugnis = myclsKlassendaten.zeugnisunterzeichner
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltflächen Formular schließen oder Abbrechen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If
End Sub
