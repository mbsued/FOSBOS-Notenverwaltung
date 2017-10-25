VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPUEinstellungenSchulaufgabenNotenKlasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim strSpalte As String
Dim strHalbjahr As String
Dim strDatum As String

Private Sub cmdAbbrechen_Click()
'Schließen ohne Speichern
    On Error GoTo Err_cmdclose_Click

    DOCMD.Close

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdSpeichern_Click()
' gib die Einstellungen in den globalen Variablen zurück
    Forms("FrmNotenNachKlassen").RueckgabeAenderungUebernehmen = True
    Forms("FrmNotenNachKlassen").RueckgabeKlasseDatum = Nz(Me.txtDatum, "")
    
    cmdAbbrechen_Click
    
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strValue As Variant

    If Not IsNull(Me.OpenArgs) Then
        strValue = Split(Me.OpenArgs, ";", , vbBinaryCompare)
        strSpalte = strValue(0)
        strHalbjahr = strValue(1)
        strDatum = strValue(2)
        
        Me.txtSpalte = strSpalte
        Me.txtHalbjahr = strHalbjahr
        Me.Caption = "Einstellungen Schulaufgabe " & strSpalte & " im " & strHalbjahr & ". Hj"
        If Not IsNull(strDatum) And Not strDatum = "" Then
            Me.txtDatum = strDatum
        End If
    End If
    
End Sub
