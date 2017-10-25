VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmHalbjahresleistungenImportieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdBeenden_Click()
On Error GoTo Err_cmdclose_Click

    DOCMD.Close

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click
End Sub

Private Sub cmdfachweise_Click()
    ' Formular ausblenden
    Forms.Item("frmHalbjahresleistungenImportieren").Visible = False
    ' Importformular für Halbjahresleistungen fachweise öffnen
    DOCMD.OpenForm "frmHalbjahresleistungenFachweiseImportieren", acNormal, , , acFormEdit
End Sub

Private Sub cmdkomplett_Click()
    ' Formular ausblenden
    Forms.Item("frmHalbjahresleistungenImportieren").Visible = False
    ' Importformular für Halbjahresleistungen komplett öffnen
    DOCMD.OpenForm "frmHalbjahresleistungenKomplettImportieren", acNormal, , , acFormEdit
End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("FrmImport") Then
        Forms.Item("FrmImport").Visible = True
    End If
End Sub

Private Sub Form_Load()
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub

