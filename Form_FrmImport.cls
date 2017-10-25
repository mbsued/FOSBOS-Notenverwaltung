VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmImport"
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

Private Sub cmdhalbjahrimportieren_Click()
    ' Formular ausblenden
    Forms.Item("frmImport").Visible = False
    ' Importformular für Halbjahresleistungen öffnen
    DOCMD.OpenForm "frmHalbjahresleistungenImportieren", acNormal, , , acFormEdit
End Sub

Private Sub cmdLehrerImport_Click()
    ' Formular ausblenden
    Forms.Item("frmImport").Visible = False
    ' Importformular für Schülerdaten öffnen
    DOCMD.OpenForm "FrmLehrerdatenImportieren", acNormal, , , acFormEdit
End Sub

Private Sub cmdSchuleImport_Click()
    ' Formular ausblenden
    Forms.Item("frmImport").Visible = False
    ' Importformular für Schuldaten öffnen
    DOCMD.OpenForm "FrmSchuldatenImportieren", acNormal, , , acFormEdit
End Sub

Private Sub cmdSchülerImport_Click()
    ' Formular ausblenden
    Forms.Item("frmImport").Visible = False
    ' Importformular für Schülerdaten öffnen
    DOCMD.OpenForm "FrmSchülerdatenImportieren", acNormal, , , acFormEdit
End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("FrmStart") Then
        Forms.Item("FrmStart").Visible = True
    End If
End Sub

Private Sub Form_Load()
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width

End Sub

