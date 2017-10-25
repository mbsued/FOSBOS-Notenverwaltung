VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmKlassenNotenbogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private booBeenden As Boolean

Private Sub cmbKlasse_AfterUpdate()
' Button Drucken mit Klassengruppe füllen
    Me.cmddrucken.Caption = "Notenbögen für die Klasse " & Me.cmbKlasse.Column(1) & " drucken"
End Sub

Private Sub cmddrucken_Click()
' Ausdruck der Notenbögen für die gewählte Klasse
' Übergabe von KlassenUid und JahrgangstufenUid
    
    DOCMD.OpenReport "rptNotenbogen", acViewPreview, , , acWindowNormal, "1;" & Me.cmbKlasse.Column(0) & " ;" & Me.cmbKlasse.Column(4)
    
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
    
    cmbKlasse_AfterUpdate
    
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

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltfläche Formular schließen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If
End Sub
