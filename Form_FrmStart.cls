VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private booBeenden As Boolean   'Beenden von ACCESS nur wenn Anwendung beenden geklickt wird

Private Sub cmdBeenden_Click()
' Anwendung beenden
    booBeenden = True
    DOCMD.Quit
    
End Sub

Private Sub cmdhalbjahr_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular Halbjahresleistungen des Schülers öffnen
    DOCMD.OpenForm "FrmHalbjahresleistungenSchüler", acNormal, , , acFormEdit
End Sub

Private Sub cmdImport_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Importformular  öffnen
    DOCMD.OpenForm "FrmImport", acNormal, , , acFormEdit
End Sub

Private Sub cmdKlasseFaecher_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular zum Löschen der Fächer öffnen
    DOCMD.OpenForm "FrmKlassenFächer", acNormal, , , acFormEdit
End Sub

Private Sub cmdKlassenPflegen_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular Notenbögen klassenweise öffnen
    DOCMD.OpenForm "FrmKlassendaten", acNormal, , , acFormEdit
End Sub

Private Sub cmdLehrerdatenPflegen_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular Notenbögen klassenweise öffnen
    DOCMD.OpenForm "FrmLehrerdaten", acNormal, , , acFormEdit
End Sub

Private Sub cmdnotenbogen_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular Notenbögen klassenweise öffnen
    DOCMD.OpenForm "FrmKlassenNotenbogen", acNormal, , , acFormEdit
End Sub

Private Sub cmdNotenEingabe_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' NotenEingabe öffnen
    DOCMD.OpenForm "FrmNotenNachKlassen", acNormal, , , acFormEdit
End Sub

Private Sub cmdSchuelerPflegen_Click()
    ' Formular ausblenden
    Forms.Item("frmStart").Visible = False
    ' Formular Notenbögen klassenweise öffnen
    DOCMD.OpenForm "FrmSchülerdaten", acNormal, , , acFormEdit
End Sub

Private Sub Form_Load()
Dim booRelinkDb As Boolean
On Error GoTo Fehler
'beenden auf Falsch setzen
    booBeenden = False
'Höhe und Breite einstellen
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
' Kontextmenü erzeugen
    ErzeugeKontextmenue
'Tabellen neu verknüpfen
    booRelinkDb = RelinkDb(CurrentProject.Path)
' Ribbon ausblenden
'    DOCMD.ShowToolbar "Ribbon", acToolbarNo
'Navigationsfenster ausblenden
    DOCMD.SelectObject acTable, vbNullString, True
    DOCMD.RunCommand acCmdWindowHide
Ende:
    On Error Resume Next
    Exit Sub
    
Fehler:
    MsgBox "Fehler-Nummer:" & Err.Number & vbCrLf & "Beschreibung: " & Err.Description, vbOKOnly, "FOSBOS Notenverwaltung"
    Resume Ende
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        If IstFormularGeoeffnet("FrmStart") Then
        Forms.Item("FrmStart").Visible = True
        Cancel = True
    End If
     
    End If
End Sub
