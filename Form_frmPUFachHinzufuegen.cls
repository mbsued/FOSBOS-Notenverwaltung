VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPUFachHinzufuegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim lngKlassengruppeUid As Long

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
' Speichere das Fach in der Tabelle TabNotenFachKlassengruppe und lege die Notensätze für alle Schüler dieser Klasse an.
' melde anschließend an das Hauptformular zurück
Dim myClsDb As clsDB

    Set myClsDb = New clsDB
    myClsDb.KlassenNotenFachHinzufuegen lngKlassengruppeUid, Me.cmbfach
    myClsDb.SchuelerNotenEinzelnesFachAnlegen lngKlassengruppeUid, Me.cmbfach
    Forms("FrmKlassenFächer").FachHinzugefuegt = True
    
    cmdAbbrechen_Click
    
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strValue As Variant
Dim strFaecherUids As Variant
Dim strSql As String
Dim strSelect As String
Dim strWhere As String
Dim strWhere2 As String
Dim strOrder As String
Dim intCount As Integer

    strWhere2 = ""
    If Not IsNull(Me.OpenArgs) Then
        strValue = Split(Me.OpenArgs, ";", , vbBinaryCompare)
        lngKlassengruppeUid = strValue(0)
        strFaecherUids = Split(strValue(1), ",", , vbBinaryCompare)
        For intCount = 0 To UBound(strFaecherUids)
            strWhere2 = strWhere2 & " AND TabFach.uid <> " & strFaecherUids(intCount)
        Next intCount
    End If
    
    strSelect = "SELECT TabFach.uid, TabFach.bezeichnung_lang FROM TabFach "
    strOrder = "ORDER BY TabFach.sort_uid, TabFach.kuerzel;"
    strWhere = "WHERE TabFach.profil=True " & strWhere2 & " "
    strSql = strSelect & strWhere & strOrder
    Me.cmbfach.RowSource = strSql
    
End Sub
