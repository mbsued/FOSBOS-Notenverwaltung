VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptNotenbogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)

Dim intCount As Integer
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String

    Select Case Me.anz_sa_hj1
        Case 0
            Me.txt_sa1_hj1.Visible = False
            Me.txt_sa2_hj1.Visible = False
        Case 1
            Me.txt_sa1_hj1.Visible = True
            Me.txt_sa2_hj1.Visible = False
        Case 2
            Me.txt_sa1_hj1.Visible = True
            Me.txt_sa2_hj1.Visible = True
        Case Else
            Me.txt_sa1_hj1.Visible = False
            Me.txt_sa2_hj1.Visible = False
    End Select
    
    Select Case Me.anz_sa_hj2
        Case 0
            Me.txt_sa1_hj2.Visible = False
            Me.txt_sa2_hj2.Visible = False
        Case 1
            Me.txt_sa1_hj2.Visible = True
            Me.txt_sa2_hj2.Visible = False
        Case 2
            Me.txt_sa1_hj2.Visible = True
            Me.txt_sa2_hj2.Visible = True
        Case Else
            Me.txt_sa1_hj2.Visible = False
            Me.txt_sa2_hj2.Visible = False
    End Select
   
End Sub

Private Sub Report_Open(Cancel As Integer)
' Einstellen der richtigen Klasse und der richtigen Halbjahresfelder
Dim varOpenargs As Variant
Dim lngArt As Long                  'woher kommt der Aufruf? 1 = Klassenweise ausdrucken, 2 = einzelner Schüler
Dim lngUidKlasse As Long
Dim lngUidJahrgangsstufe As Long
Dim lngUidSchueler As Long
Dim StrFelder As String
Dim StrHalbjahresFelder As String
Dim strTable As String
Dim strWhere As String
Dim strOrder As String
Dim strSql As String

    On Error GoTo Err_Report_Open
    
    If Not IsNull(Me.OpenArgs) Then
        varOpenargs = Split(Me.OpenArgs, ";", , vbBinaryCompare)
        lngArt = varOpenargs(0)
        Select Case lngArt
            Case 1
                lngUidKlasse = varOpenargs(1)
                lngUidJahrgangsstufe = varOpenargs(2)
                
                strWhere = "TabSchueler.uid_klassengruppe=" & lngUidKlasse & " AND TabNotenSchuelerFachKlassengruppe.geloescht=FALSE AND TabNotenHalbjahrSchuelerFach.geloescht=FALSE"
            Case 2
                lngUidSchueler = varOpenargs(1)
                lngUidJahrgangsstufe = varOpenargs(2)
                
                strWhere = "TabSchueler.uid=" & lngUidSchueler & " AND TabNotenSchuelerFachKlassengruppe.geloescht=FALSE AND TabNotenHalbjahrSchuelerFach.geloescht=FALSE"
        End Select
        
        strOrder = "TabFach.sort_uid"
        
        StrFelder = "TabSchule.schulnummer, TabSchule.bezeichnung_kurz, TabSchueler.nachname , TabSchueler.rufname, TabSchueler.uid_geschlecht, TabSchueler.gebdat, TabSchueler.gebort, TabKlasse.bezeichnung, TabJahrgangsstufe.bezeichnung_kurz, TabAusbildungsrichtung.bezeichnung_kurz,TabFach.kuerzel,TabFach.sort_uid," & _
                    "TabSchueler.bekenntnis, TabSchueler.religion_unterricht, TabSchueler.eintritt_schule, TabSchueler.probezeit_bis, " & _
                    "TabNotenSchuelerFachKlassengruppe.anz_sa_hj1,TabNotenSchuelerFachKlassengruppe.anz_sa_hj2," & _
                    "TabNotenSchuelerFachKlassengruppe.uid_art_son1_hj1,TabNotenSchuelerFachKlassengruppe.uid_art_son2_hj1,TabNotenSchuelerFachKlassengruppe.uid_art_son3_hj1,TabNotenSchuelerFachKlassengruppe.uid_art_son4_hj1, TabNotenSchuelerFachKlassengruppe.uid_art_son5_hj1, TabNotenSchuelerFachKlassengruppe.uid_art_son6_hj1, TabNotenSchuelerFachKlassengruppe.uid_art_son7_hj1, TabNotenSchuelerFachKlassengruppe.uid_art_son1_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son2_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son3_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son4_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son5_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son6_hj2, TabNotenSchuelerFachKlassengruppe.uid_art_son7_hj2," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son1_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son1_hj1) AS gew_so1_hj1,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son2_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son2_hj1) AS gew_so2_hj1,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son3_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son3_hj1) AS gew_so3_hj1," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son4_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son4_hj1) AS gew_so4_hj1,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son5_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son5_hj1) AS gew_so5_hj1,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son6_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son6_hj1) AS gew_so6_hj1," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son7_hj1=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son7_hj1) AS gew_so7_hj1," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son1_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son1_hj2) AS gew_so1_hj2,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son2_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son2_hj2) AS gew_so2_hj2,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son3_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son3_hj2) AS gew_so3_hj2," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son4_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son4_hj2) AS gew_so4_hj2,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son5_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son5_hj2) AS gew_so5_hj2,IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son6_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son6_hj2) AS gew_so6_hj2," & _
                    "IIf(TabNotenSchuelerFachKlassengruppe.uid_art_son7_hj2=0,"""",TabNotenSchuelerFachKlassengruppe.gew_son7_hj2) AS gew_so7_hj2," & _
                    "TabNotenSchuelerFachKlassengruppe.n_sa1_hj1,TabNotenSchuelerFachKlassengruppe.n_sa2_hj1,TabNotenSchuelerFachKlassengruppe.n_son1_hj1,TabNotenSchuelerFachKlassengruppe.n_son2_hj1,TabNotenSchuelerFachKlassengruppe.n_son3_hj1,TabNotenSchuelerFachKlassengruppe.n_son4_hj1,TabNotenSchuelerFachKlassengruppe.n_son5_hj1,TabNotenSchuelerFachKlassengruppe.n_son6_hj1,TabNotenSchuelerFachKlassengruppe.n_son7_hj1," & _
                    "TabNotenSchuelerFachKlassengruppe.n_sa1_hj2,TabNotenSchuelerFachKlassengruppe.n_sa2_hj2,TabNotenSchuelerFachKlassengruppe.n_son1_hj2,TabNotenSchuelerFachKlassengruppe.n_son2_hj2,TabNotenSchuelerFachKlassengruppe.n_son3_hj2,TabNotenSchuelerFachKlassengruppe.n_son4_hj2,TabNotenSchuelerFachKlassengruppe.n_son5_hj2,TabNotenSchuelerFachKlassengruppe.n_son6_hj2,TabNotenSchuelerFachKlassengruppe.n_son7_hj2,"
        
        Select Case lngUidJahrgangsstufe
            Case 1, 5, 6 'Vorklasse und Vorkurs
                StrHalbjahresFelder = "TabNotenHalbjahrSchuelerFach.n_vkl_hj1 AS n_hj1, TabNotenHalbjahrSchuelerFach.n_vkl_hj2 AS n_hj2"
            Case 2  '11.Klasse
                StrHalbjahresFelder = "TabNotenHalbjahrSchuelerFach.n_11_hj1 AS n_hj1, TabNotenHalbjahrSchuelerFach.n_11_hj2 AS n_hj2"
            Case 3  '12.Klasse
                StrHalbjahresFelder = "TabNotenHalbjahrSchuelerFach.n_12_hj1 AS n_hj1, TabNotenHalbjahrSchuelerFach.n_12_hj2 AS n_hj2"
            Case 4  '13.Klasse
                StrHalbjahresFelder = "TabNotenHalbjahrSchuelerFach.n_13_hj1 AS n_hj1, TabNotenHalbjahrSchuelerFach.n_13_hj2 AS n_hj2"
        End Select
        
        strTable = "TabSchule INNER JOIN (TabKlasse INNER JOIN (((TabJahrgangsstufe INNER JOIN (TabAusbildungsrichtung INNER JOIN TabKlassengruppe ON TabAusbildungsrichtung.uid = TabKlassengruppe.uid_ausb) " & _
                    "ON TabJahrgangsstufe.uid = TabKlassengruppe.uid_jahrgangsstufe) INNER JOIN " & _
                    "((TabFach INNER JOIN TabNotenHalbjahrSchuelerFach ON TabFach.uid = TabNotenHalbjahrSchuelerFach.uid_fach) INNER JOIN " & _
                    "TabNotenSchuelerFachKlassengruppe ON TabFach.uid = TabNotenSchuelerFachKlassengruppe.uid_fach) ON " & _
                    "(TabKlassengruppe.uid = TabNotenSchuelerFachKlassengruppe.uid_klassengruppe) AND (TabKlassengruppe.uid = TabNotenHalbjahrSchuelerFach.uid_klassengruppe)) " & _
                    "INNER JOIN TabSchueler ON (TabSchueler.uid = TabNotenSchuelerFachKlassengruppe.uid_schueler) AND " & _
                    "(TabSchueler.uid = TabNotenHalbjahrSchuelerFach.uid_schueler) AND (TabKlassengruppe.uid = TabSchueler.uid_klassengruppe)) ON TabKlasse.uid = TabKlassengruppe.uid_klasse) " & _
                    " ON TabSchule.uid_schulart = TabKlassengruppe.uid_schulart"
        strSql = "SELECT " & StrFelder & StrHalbjahresFelder & " FROM " & strTable & " WHERE " & strWhere & " ORDER BY " & strOrder
        Me.RecordSource = strSql
    End If
    
Exit_Report_Open:
    Exit Sub
    
Err_Report_Open:
    FehlermeldungAusgeben "Notenbogen drucken", Err.Source, Err.Number, Err.Description
    Resume Exit_Report_Open

End Sub
