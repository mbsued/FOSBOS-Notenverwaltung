VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmKlassenF�cher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private myColKlasseFaecher As colKlasseF�cher       'Sammlung der F�cher
Private booFormularAenderung As Boolean             'Wurde eine �nderung durchgef�hrt?
Private booFachHinzugefuegt As Boolean              'ist ein Fach hinzugef�gt worden?

Private booBeenden As Boolean
Private Const lngColorWhite As Long = 16777215
Private Const lngColorBlack As Long = 0
Private Const lngColorRed As Long = 255
Private Const lngColorBlue As Long = 15123357

Private Sub cmbKlasse_AfterUpdate()
' Pr�fen ob vorher �nderung war, wenn ja speichern
' Daten der F�cher einlesen
' entsprechende Felder einblenden
' jetzt die entsprechenden Daten lesen

    If booFormularAenderung Then
        If MsgBox("Sollen die �nderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If

    Set myColKlasseFaecher = KlasseFaecherLesen(Me.cmbKlasse)
' erst mal wieder alles einblenden anschlie�end ausblenden
    AlleFelderEinblenden
    FelderAusblenden
' Daten f�llen
    FaecherNameSchreiben

End Sub

Private Sub cmd_loesch_1_Click()
    EintragLoeschen 1
End Sub

Private Sub cmd_loesch_2_Click()
    EintragLoeschen 2
End Sub

Private Sub cmd_loesch_3_Click()
    EintragLoeschen 3
End Sub

Private Sub cmd_loesch_4_Click()
    EintragLoeschen 4
End Sub

Private Sub cmd_loesch_5_Click()
    EintragLoeschen 5
End Sub

Private Sub cmd_loesch_6_Click()
    EintragLoeschen 6
End Sub

Private Sub cmd_loesch_7_Click()
    EintragLoeschen 7
End Sub

Private Sub cmd_loesch_8_Click()
    EintragLoeschen 8
End Sub

Private Sub cmd_loesch_9_Click()
    EintragLoeschen 9
End Sub

Private Sub cmd_loesch_10_Click()
    EintragLoeschen 10
End Sub

Private Sub cmd_loesch_11_Click()
    EintragLoeschen 11
End Sub

Private Sub cmd_loesch_12_Click()
    EintragLoeschen 12
End Sub

Private Sub cmd_loesch_13_Click()
    EintragLoeschen 13
End Sub

Private Sub cmd_loesch_14_Click()
    EintragLoeschen 14
End Sub

Private Sub cmd_loesch_15_Click()
    EintragLoeschen 15
End Sub

Private Sub cmdAbbrechen_Click()
' Formular schlie�en vorher nachfragen ob �nderungen verworfen werden sollen

    On Error GoTo Err_cmdclose_Click
    If booFormularAenderung Then
        If MsgBox("Sollen die �nderungen verworfen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
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

Private Sub cmdHinzufuegen_Click()
' PopUp Formular �ffnen
' belegte F�cher ermitteln
Dim strVorhandeneFaecher As String
Dim intCount As Integer

    strVorhandeneFaecher = ";"
    
    For intCount = 1 To myColKlasseFaecher.Count
        If myColKlasseFaecher.Item(intCount).fach_uid > 10 And myColKlasseFaecher.Item(intCount).fach_uid < 33 Then
            strVorhandeneFaecher = strVorhandeneFaecher & myColKlasseFaecher.Item(intCount).fach_uid
            If intCount < myColKlasseFaecher.Count Then
                strVorhandeneFaecher = strVorhandeneFaecher & ","
            End If
        End If
    Next intCount

    booFachHinzugefuegt = False
    
    DOCMD.OpenForm "frmPUFachHinzufuegen", acNormal, , , acFormEdit, acDialog, Me.cmbKlasse.Column(0) & strVorhandeneFaecher
    
    If booFachHinzugefuegt = True Then  'Fach wurde hinzugef�gt
        cmbKlasse_AfterUpdate
    End If
    
End Sub

Private Sub cmdSpeichern_Click()
' die vorgegebenen Daten speichern (KlasseF�cher und Sch�lerdaten)
    KlasseFaecherSchreiben myColKlasseFaecher
    ' Formular ge�ndert zur�cksetzen
    booFormularAenderung = False

End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("frmStart") Then
        Forms.Item("frmStart").Visible = True
    End If
End Sub
Private Sub Form_Load()
'H�he und Breite einstellen
'Formular �nderung auf falsch stellen
'1. Klassengruppe einstellen

    booBeenden = False
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
            
    booFormularAenderung = False
    cmbKlasse_AfterUpdate
    
End Sub

Private Sub cmdClose_Click()
'Schlie�en vorher nachfragen ob gespeichert werden soll
    On Error GoTo Err_cmdclose_Click
    
    If booFormularAenderung Then
        If MsgBox("Sollen die �nderungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
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
Private Sub AlleFelderEinblenden()
Dim strControl1 As String
Dim strControl2  As String
Dim intCount As Integer

    For intCount = 1 To 15
    
        strControl1 = "lbl_fach_" & intCount
        strControl2 = "cmd_loesch_" & intCount
        
        Me.Controls(strControl1).Visible = True
        Me.Controls(strControl1).BackColor = lngColorBlue
        Me.Controls(strControl1).ForeColor = lngColorBlack
        
        Me.Controls(strControl2).Visible = True
        Me.Controls(strControl2).Enabled = True
        
    Next intCount
    
    Me.cmdHinzufuegen.Visible = True
    If Me.cmbKlasse.Column(3) <> 1 Then
        Me.cmdHinzufuegen.Visible = False
    End If
    
End Sub
Private Sub FelderAusblenden()
Dim strControl1 As String
Dim strControl2  As String
Dim intCount As Integer

    For intCount = myColKlasseFaecher.Count + 1 To 15
    
        strControl1 = "lbl_fach_" & intCount
        strControl2 = "cmd_loesch_" & intCount
        
        Me.Controls(strControl1).Visible = False
        Me.Controls(strControl2).Visible = False
        
    Next intCount
End Sub
Private Sub FaecherNameSchreiben()

Dim strControl1 As String
Dim strControl2  As String
Dim intCount As Integer

    For intCount = 1 To myColKlasseFaecher.Count
    
        strControl1 = "lbl_fach_" & intCount
        strControl2 = "cmd_loesch_" & intCount
        
        Me.Controls(strControl1).Caption = myColKlasseFaecher.Item(intCount).fachbezeichnung
        Me.Controls(strControl2).ControlTipText = "Ein Klick auf diese Schaltfl�che l�scht die Zuordnung des Faches " & myColKlasseFaecher.Item(intCount).fachbezeichnung & " zur Klassengruppe " & Me.cmbKlasse.Column(1) & "." & vbCrLf & _
                                                    "Bereits f�r dieses Fach vorgegebene Sch�lernoten werden ebenfalls gel�scht."
        
    Next intCount
End Sub
Private Sub EintragLoeschen(intFach As Integer)
'L�schen des Eintrags
Dim strControl1 As String
Dim strControl2  As String
Dim strMsg As String

    strMsg = "Wollen sie wirklich die Zuordnung des Faches " & myColKlasseFaecher.Item(intFach).fachbezeichnung & " zur Klassengruppe " & Me.cmbKlasse.Column(1) & " l�schen?" & vbCrLf & _
                "Bereits f�r dieses Fach vorgegebene Sch�lernoten werden ebenfalls gel�scht" & vbCrLf & _
                "Die endg�ltige L�schung erfolgt erst, wenn Sie speichern anklicken!"

    If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
        myColKlasseFaecher.Item(intFach).geloescht = True
        booFormularAenderung = True
        
        strControl1 = "lbl_fach_" & intFach
        strControl2 = "cmd_loesch_" & intFach
        
        Me.Controls(strControl1).BackColor = lngColorRed
        Me.Controls(strControl1).ForeColor = lngColorWhite
        Me.Controls(strControl2).Enabled = False

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schlie�en nur �ber die Schaltfl�chen Formular schlie�en oder Abbrechen m�glich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If
End Sub
' R�ckgabewert aus dem Formular FachHinzufuegen
Public Property Let FachHinzugefuegt(ByVal booVar As Boolean)
    booFachHinzugefuegt = booVar
End Property

