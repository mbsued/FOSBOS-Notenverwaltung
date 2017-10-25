VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSchülerdatenImportieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private booBeenden As Boolean

Private Sub cmdauswaehlen_Click()
    txtdatei = vbGetOpenFilename("Textdateien" & vbNullChar & "*.txt;*.csv", "*.txt", CurrentProject.Path)

End Sub

Private Sub cmdClose_Click()
On Error GoTo Err_cmdclose_Click

    booBeenden = True
    DOCMD.Close

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdImport_Click()

    If Not IsNull(Me.txtdatei) And Me.txtdatei <> "" Then
        SchuelerImportieren Me.txtdatei
    Else
        MsgBox "Keine Datei zum Importieren ausgewählt!", vbCritical, "Schüler importieren"
    End If

End Sub

Private Sub Form_Load()
    booBeenden = False
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub
Private Sub Form_Close()
    If IstFormularGeoeffnet("FrmImport") Then
        Forms.Item("FrmImport").Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltfläche Formular schließen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If

End Sub
