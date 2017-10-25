VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPUEinstellungenNotenSchueler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim StrSchueler As String
Dim strSchuelerName  As String
Dim strAnzahlSa1 As String
Dim strAnzahlSa2 As String
Dim strAnzahlSon1 As String
Dim strAnzahlSon2 As String
Dim varSaHj1(0 To 1) As Variant
Dim varSaHj2(0 To 1) As Variant
Dim varSonsHj1(0 To 20) As Variant
Dim varSonsHj2(0 To 20) As Variant
Private booAenderung As Boolean
Private mycolSchuelerdaten As colSchülerNoten
Private myclsKlasseNoten As clsKlasseNoten

Private Sub cmbgew1_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew1_1, 1, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew2_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew2_1, 4, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew3_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew3_1, 7, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew4_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew4_1, 10, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew5_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew5_1, 13, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew6_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew6_1, 16, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew7_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew7_1, 19, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew1_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew1_2, 1, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew2_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew2_2, 4, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew3_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew3_2, 7, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew4_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew4_2, 10, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew5_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew5_2, 13, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew6_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew6_2, 16, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmbgew7_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(cmbgew7_2, 19, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub cmdAbbrechen_Click()
'Schließen ohne Speichern
    On Error GoTo Err_cmdclose_Click

    DOCMD.Close acForm, "frmPUEinstellungenNotenSchueler"

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdSpeichern_Click()
' gib die Einstellungen zurück
    If booAenderung = True Then
        Forms("FrmNotenNachKlassen").RueckgabeAenderungUebernehmen = True
        VariablenSetzen
    End If
    
    cmdAbbrechen_Click
    
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strValue As Variant
Dim intCount As Integer
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim strcontrol4 As String
Dim strControl5 As String
Dim strControl6 As String
Dim strControl7 As String
Dim intCountSons As Integer

    booAenderung = False
    
    Set mycolSchuelerdaten = Forms("FrmNotenNachKlassen").Schuelerdaten
    Set myclsKlasseNoten = Forms("FrmNotenNachKlassen").Klassendaten
    
    If Not IsNull(Me.OpenArgs) Then
        StrSchueler = Me.OpenArgs
        strAnzahlSa1 = myclsKlasseNoten.anz_sa_hj1
        strAnzahlSa2 = myclsKlasseNoten.anz_sa_hj2
        strAnzahlSon1 = myclsKlasseNoten.anz_son_hj1
        strAnzahlSon2 = myclsKlasseNoten.anz_son_hj2
        varSaHj1(0) = mycolSchuelerdaten.Item(StrSchueler).dat_sa1_hj1
        varSaHj1(1) = mycolSchuelerdaten.Item(StrSchueler).dat_sa2_hj1
        varSaHj2(0) = mycolSchuelerdaten.Item(StrSchueler).dat_sa1_hj2
        varSaHj2(1) = mycolSchuelerdaten.Item(StrSchueler).dat_sa2_hj2
        
        varSonsHj1(0) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son1_hj1
        varSonsHj1(1) = mycolSchuelerdaten.Item(StrSchueler).gew_son1_hj1
        varSonsHj1(2) = mycolSchuelerdaten.Item(StrSchueler).dat_son1_hj1
        varSonsHj1(3) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son2_hj1
        varSonsHj1(4) = mycolSchuelerdaten.Item(StrSchueler).gew_son2_hj1
        varSonsHj1(5) = mycolSchuelerdaten.Item(StrSchueler).dat_son2_hj1
        varSonsHj1(6) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son3_hj1
        varSonsHj1(7) = mycolSchuelerdaten.Item(StrSchueler).gew_son3_hj1
        varSonsHj1(8) = mycolSchuelerdaten.Item(StrSchueler).dat_son3_hj1
        varSonsHj1(9) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son4_hj1
        varSonsHj1(10) = mycolSchuelerdaten.Item(StrSchueler).gew_son4_hj1
        varSonsHj1(11) = mycolSchuelerdaten.Item(StrSchueler).dat_son4_hj1
        varSonsHj1(12) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son5_hj1
        varSonsHj1(13) = mycolSchuelerdaten.Item(StrSchueler).gew_son5_hj1
        varSonsHj1(14) = mycolSchuelerdaten.Item(StrSchueler).dat_son5_hj1
        varSonsHj1(15) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son6_hj1
        varSonsHj1(16) = mycolSchuelerdaten.Item(StrSchueler).gew_son6_hj1
        varSonsHj1(17) = mycolSchuelerdaten.Item(StrSchueler).dat_son6_hj1
        varSonsHj1(18) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son7_hj1
        varSonsHj1(19) = mycolSchuelerdaten.Item(StrSchueler).gew_son7_hj1
        varSonsHj1(20) = mycolSchuelerdaten.Item(StrSchueler).dat_son7_hj1
        
        varSonsHj2(0) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son1_hj2
        varSonsHj2(1) = mycolSchuelerdaten.Item(StrSchueler).gew_son1_hj2
        varSonsHj2(2) = mycolSchuelerdaten.Item(StrSchueler).dat_son1_hj2
        varSonsHj2(3) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son2_hj2
        varSonsHj2(4) = mycolSchuelerdaten.Item(StrSchueler).gew_son2_hj2
        varSonsHj2(5) = mycolSchuelerdaten.Item(StrSchueler).dat_son2_hj2
        varSonsHj2(6) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son3_hj2
        varSonsHj2(7) = mycolSchuelerdaten.Item(StrSchueler).gew_son3_hj2
        varSonsHj2(8) = mycolSchuelerdaten.Item(StrSchueler).dat_son3_hj2
        varSonsHj2(9) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son4_hj2
        varSonsHj2(10) = mycolSchuelerdaten.Item(StrSchueler).gew_son4_hj2
        varSonsHj2(11) = mycolSchuelerdaten.Item(StrSchueler).dat_son4_hj2
        varSonsHj2(12) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son5_hj2
        varSonsHj2(13) = mycolSchuelerdaten.Item(StrSchueler).gew_son5_hj2
        varSonsHj2(14) = mycolSchuelerdaten.Item(StrSchueler).dat_son5_hj2
        varSonsHj2(15) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son6_hj2
        varSonsHj2(16) = mycolSchuelerdaten.Item(StrSchueler).gew_son6_hj2
        varSonsHj2(17) = mycolSchuelerdaten.Item(StrSchueler).dat_son6_hj2
        varSonsHj2(18) = mycolSchuelerdaten.Item(StrSchueler).uid_art_son7_hj2
        varSonsHj2(19) = mycolSchuelerdaten.Item(StrSchueler).gew_son7_hj2
        varSonsHj2(20) = mycolSchuelerdaten.Item(StrSchueler).dat_son7_hj2
       
        strSchuelerName = mycolSchuelerdaten.Item(StrSchueler).nachname & " " & mycolSchuelerdaten.Item(StrSchueler).rufname
        
        Me.Caption = "Einstellungen für den Schüler " & strSchuelerName

        Me.txtSchueler = StrSchueler
        Me.txtAnzahlSaHj1 = strAnzahlSa1
        Me.txtAnzahlSAHj2 = strAnzahlSa2
        Me.txtAnzahlSonHj1 = strAnzahlSon1
        Me.txtAnzahlSonHj2 = strAnzahlSon2
        Me.txtdatumsa1_1 = Nz(varSaHj1(0), "")
        Me.txtdatumsa2_1 = Nz(varSaHj1(1), "")
        Me.txtdatumsa1_2 = Nz(varSaHj2(0), "")
        Me.txtdatumsa2_2 = Nz(varSaHj2(1), "")
        
        intCountSons = 0
        
        For intCount = 1 To strAnzahlSon1
            strControl1 = "cmbart" & intCount & "_1"
            strControl2 = "cmbgew" & intCount & "_1"
            StrControl3 = "txtdatson" & intCount & "_1"
            Me.Controls(strControl1).Value = varSonsHj1(intCountSons)
            Me.Controls(strControl2).Value = varSonsHj1(intCountSons + 1)
            Me.Controls(StrControl3).Value = varSonsHj1(intCountSons + 2)
            intCountSons = intCountSons + 3
        Next intCount
        
        intCountSons = 0
        
        For intCount = 1 To strAnzahlSon2
            strcontrol4 = "cmbart" & intCount & "_2"
            strControl5 = "cmbgew" & intCount & "_2"
            strControl6 = "txtdatson" & intCount & "_2"
            Me.Controls(strcontrol4).Value = varSonsHj2(intCountSons)
            Me.Controls(strControl5).Value = varSonsHj2(intCountSons + 1)
            Me.Controls(strControl6).Value = varSonsHj2(intCountSons + 2)
            intCountSons = intCountSons + 3
        Next intCount
        
        Select Case strAnzahlSa1
            Case 1
                Me.lbldatsa1_1.Visible = True
                Me.txtdatumsa1_1.Visible = True
                Me.lbldatsa2_1.Visible = False
                Me.txtdatumsa2_1.Visible = False
            Case 2
                Me.lbldatsa1_1.Visible = True
                Me.txtdatumsa1_1.Visible = True
                Me.lbldatsa2_1.Visible = True
                Me.txtdatumsa2_1.Visible = True
            Case Else
                Me.lbldatsa1_1.Visible = False
                Me.txtdatumsa1_1.Visible = False
                Me.lbldatsa2_1.Visible = False
                Me.txtdatumsa2_1.Visible = False
        End Select
        
        Select Case strAnzahlSa2
            Case 1
                Me.lbldatsa1_2.Visible = True
                Me.txtdatumsa1_2.Visible = True
                Me.lbldatsa2_2.Visible = False
                Me.txtdatumsa2_2.Visible = False
            Case 2
                Me.lbldatsa1_2.Visible = True
                Me.txtdatumsa1_2.Visible = True
                Me.lbldatsa2_2.Visible = True
                Me.txtdatumsa2_2.Visible = True
            Case Else
                Me.lbldatsa1_2.Visible = False
                Me.txtdatumsa1_2.Visible = False
                Me.lbldatsa2_2.Visible = False
                Me.txtdatumsa2_2.Visible = False
        End Select
        
        For intCount = strAnzahlSon1 + 1 To 7
            strControl1 = "lblson" & intCount & "_1"
            strControl2 = "lblart" & intCount & "_1"
            StrControl3 = "cmbart" & intCount & "_1"
            strcontrol4 = "lblgew" & intCount & "_1"
            strControl5 = "cmbgew" & intCount & "_1"
            strControl6 = "lbldat" & intCount & "_1"
            strControl7 = "txtdatson" & intCount & "_1"
            Me.Controls(strControl1).Visible = False
            Me.Controls(strControl2).Visible = False
            Me.Controls(StrControl3).Visible = False
            Me.Controls(strcontrol4).Visible = False
            Me.Controls(strControl5).Visible = False
            Me.Controls(strControl6).Visible = False
            Me.Controls(strControl7).Visible = False
        Next intCount
        
        For intCount = strAnzahlSon1 + 1 To 7
            strControl1 = "lblson" & intCount & "_2"
            strControl2 = "lblart" & intCount & "_2"
            StrControl3 = "cmbart" & intCount & "_2"
            strcontrol4 = "lblgew" & intCount & "_2"
            strControl5 = "cmbgew" & intCount & "_2"
            strControl6 = "lbldat" & intCount & "_2"
            strControl7 = "txtdatson" & intCount & "_2"
            Me.Controls(strControl1).Visible = False
            Me.Controls(strControl2).Visible = False
            Me.Controls(StrControl3).Visible = False
            Me.Controls(strcontrol4).Visible = False
            Me.Controls(strControl5).Visible = False
            Me.Controls(strControl6).Visible = False
            Me.Controls(strControl7).Visible = False
        Next intCount
    End If
    
End Sub
Public Function CheckGeaendert(ctl As Control, intElement As Integer, intHalbjahr As Integer) As Boolean
' prüft, ob eine Änderung war?
    Select Case intHalbjahr
        Case 1  '1.Halbjahr Prüfung gegen Variable varSonsHj1
            If varSonsHj1(intElement) <> ctl.Value Then
                CheckGeaendert = True
            Else
                CheckGeaendert = False
            End If
        Case 2  '2.Halbjahr Prüfung gegen Variable varSonsHj2
            If varSonsHj2(intElement) <> ctl.Value Then
                CheckGeaendert = True
            Else
                CheckGeaendert = False
            End If
    End Select
End Function
Public Function CheckGeaendertSa(ctl As Control, intElement As Integer, intHalbjahr As Integer) As Boolean
' prüft, ob eine Änderung war?
    Select Case intHalbjahr
        Case 1  '1.Halbjahr Prüfung gegen Variable varSaHj1
            If varSaHj1(intElement) <> ctl.Value Then
                CheckGeaendertSa = True
            Else
                CheckGeaendertSa = False
            End If
        Case 2  '2.Halbjahr Prüfung gegen Variable varSaHj2
            If varSaHj2(intElement) <> ctl.Value Then
                CheckGeaendertSa = True
            Else
                CheckGeaendertSa = False
            End If
    End Select
End Function

Private Sub txtdatson1_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson1_1, 2, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson2_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson2_1, 5, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson3_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson3_1, 8, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson4_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson4_1, 11, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson5_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson5_1, 14, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson6_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson6_1, 17, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson7_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson7_1, 20, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson1_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson1_2, 2, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson2_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson2_2, 5, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson3_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson3_2, 8, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson4_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson4_2, 11, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson5_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson5_2, 14, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson6_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson6_2, 17, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatson7_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendert(txtdatson7_2, 20, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatumsa1_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendertSa(txtdatumsa1_1, 0, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatumsa2_1_BeforeUpdate(Cancel As Integer)
    If CheckGeaendertSa(txtdatumsa2_1, 1, 1) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatumsa1_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendertSa(txtdatumsa1_2, 0, 2) Then
        booAenderung = True
    End If
End Sub

Private Sub txtdatumsa2_2_BeforeUpdate(Cancel As Integer)
    If CheckGeaendertSa(txtdatumsa2_2, 1, 2) Then
        booAenderung = True
    End If
End Sub
Private Sub VariablenSetzen()
' schreibe die Änderungen in die Schülerdaten
Dim intCount As Integer
Dim strControl1 As String
Dim strControl2 As String

    Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSchulaufgabenHj1(0) = Nz(Me.txtdatumsa1_1, "")
    Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSchulaufgabenHj1(1) = Nz(Me.txtdatumsa2_1, "")
    Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSchulaufgabenHj2(0) = Nz(Me.txtdatumsa1_2, "")
    Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSchulaufgabenHj2(1) = Nz(Me.txtdatumsa2_2, "")
    
    For intCount = 1 To strAnzahlSon1
        strControl1 = "cmbgew" & intCount & "_1"
        strControl2 = "txtdatson" & intCount & "_1"
        Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSonstigeGewHj1(intCount - 1) = Me.Controls(strControl1).Value
        Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSonstigeDatHj1(intCount - 1) = Nz(Me.Controls(strControl2).Value, "")
    Next intCount
    
    For intCount = 1 To strAnzahlSon2
        strControl1 = "cmbgew" & intCount & "_2"
        strControl2 = "txtdatson" & intCount & "_2"
        Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSonstigeGewHj2(intCount - 1) = Me.Controls(strControl1).Value
        Forms("FrmNotenNachKlassen").RueckgabeSchuelerdatenSonstigeDatHj2(intCount - 1) = Nz(Me.Controls(strControl2).Value, "")
    Next intCount
    
End Sub
