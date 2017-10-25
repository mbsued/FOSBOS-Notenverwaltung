Attribute VB_Name = "ModModulErzeugen"
Option Compare Database
Option Explicit

Sub ErzeugeModule()

Dim frm As Form
Dim mdl As Module
Dim ctl As Control

Dim intCountSchueler
Dim intCountHalbjahr
Dim intCountSpalte
Dim lngReturn As Long
Dim StrControl As String
    
    Set frm = Form_FrmNotenNachKlassen
    
    Set mdl = frm.Module
    For intCountSpalte = 1 To 7
        For intCountHalbjahr = 1 To 2
            For intCountSchueler = 1 To 35
                StrControl = "txt_md" & intCountSpalte & "_" & intCountHalbjahr & "_" & intCountSchueler
                lngReturn = mdl.CreateEventProc("AfterUpdate", StrControl)
                mdl.InsertLines lngReturn + 1, vbTab & "CheckArtUndGewichtung 2," & intCountSpalte & ", " & intCountHalbjahr & ", " & intCountSchueler
    
            Next intCountSchueler
        Next intCountHalbjahr
    Next intCountSpalte
    
    Set mdl = Nothing
    Set frm = Nothing
    
End Sub
Sub formular()
    Dim frm As Form
    Dim intCount As Integer
        For intCount = 1 To Forms.Count
            Debug.Print Forms.Item(intCount).name
        Next intCount
End Sub

