VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul der Datenbankoperationen
Private myconn As ADODB.Connection
Private myColSchueler As colSchüler
Private myColLehrer As colLehrer
Private myColSchule As colSchule
Private mcolSchuelerNoten As colSchülerNoten
Private mcolSchuelerNotenKurz As colSchülerNotenKurz
Private mcolSchuelerHalbjahr As colSchülerHalbjahr
Private mcolSchuelerHalbjahrKurz As colSchülerHalbjahrKurz
Private mcolKlasseFaecher As colKlasseFächer
Private mcolHalbjahr As colHalbjahrNoten
Private mcolHalbjahrKomplett As colHalbjahrNotenKomplett

Private Sub Class_Initialize()
    Set myconn = CurrentProject.Connection
End Sub

Private Sub Class_Terminate()
    Set myconn = Nothing
End Sub

Public Function SchuelerNotenLesenKurz(lngSchuelerUid As Long) As ADODB.Recordset
'Funktion liest alle Notendatensätze eines bestimmten Schülers und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenSchuelerFachKlassengruppe.*, TabKlassengruppe.uid_jahrgangsstufe " & _
                "FROM TabKlassengruppe INNER JOIN TabNotenSchuelerFachKlassengruppe ON TabKlassengruppe.uid = TabNotenSchuelerFachKlassengruppe.uid_klassengruppe " & _
                "WHERE ((TabNotenSchuelerFachKlassengruppe.uid_schueler)=" & lngSchuelerUid & ")"
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set SchuelerNotenLesenKurz = myrs
    Set myrs = Nothing

End Function

Public Function SchuelerNotenSpeichernKurz() As Boolean
'Funktion speichert die Daten der Collection SchülerNotenKurz in der Tabelle SchülerNoten
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenSchuelerFachKlassengruppe.* " & _
                "FROM TabNotenSchuelerFachKlassengruppe  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    For intCount = 1 To mcolSchuelerNotenKurz.Count
        myrs.Find "uid=" & mcolSchuelerNotenKurz.Item(intCount).uid
        ' wurde der Schüler gefunden?
        If Not myrs.EOF Then
            myrs.Fields("uid_klassengruppe") = mcolSchuelerNotenKurz.Item(intCount).klassengruppe_uid
            myrs.Fields("geloescht") = mcolSchuelerNotenKurz.Item(intCount).geloescht
            myrs.Update
            myrs.MoveFirst
        End If
    Next intCount

End Function

Public Function SchuelerNotenLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As ADODB.Recordset
'Funktion liest die Daten des Schülers und seine Notendaten und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.nachname, TabSchueler.rufname, TabNotenSchuelerFachKlassengruppe.* " & _
                "FROM TabSchueler INNER JOIN TabNotenSchuelerFachKlassengruppe ON TabSchueler.uid = TabNotenSchuelerFachKlassengruppe.uid_schueler " & _
                "WHERE (((TabNotenSchuelerFachKlassengruppe.uid_fach)=" & lngFachUid & ") AND ((TabNotenSchuelerFachKlassengruppe.uid_klassengruppe)=" & lngKlassengruppenUid & ") AND ((TabNotenSchuelerFachKlassengruppe.geloescht=false)) AND (TabSchueler.ausgeschieden=FALSE)) " & _
                "ORDER BY TabSchueler.nachname, TabSchueler.rufname"
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set SchuelerNotenLesen = myrs
    Set myrs = Nothing

End Function
Public Function SchuelerNotenSpeichern() As Boolean
'Funktion speichert die Daten der Collection in der Tabelle SchülerNoten
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenSchuelerFachKlassengruppe.* " & _
                "FROM TabNotenSchuelerFachKlassengruppe  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    For intCount = 1 To mcolSchuelerNoten.Count
        myrs.Find "uid=" & mcolSchuelerNoten.Item(intCount).uid
        ' wurde der Schüler gefunden?
        If Not myrs.EOF Then
            myrs.Fields("ind_einstellung") = mcolSchuelerNoten.Item(intCount).ind_einstellung
            myrs.Fields("anz_sa_hj1") = mcolSchuelerNoten.Item(intCount).anz_sa_hj1
            myrs.Fields("anz_sa_hj2") = mcolSchuelerNoten.Item(intCount).anz_sa_hj2
            myrs.Fields("anz_son_hj1") = mcolSchuelerNoten.Item(intCount).anz_son_hj1
            myrs.Fields("anz_son_hj2") = mcolSchuelerNoten.Item(intCount).anz_son_hj2
            
            myrs.Fields("uid_art_son1_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son1_hj1
            myrs.Fields("uid_art_son2_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son2_hj1
            myrs.Fields("uid_art_son3_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son3_hj1
            myrs.Fields("uid_art_son4_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son4_hj1
            myrs.Fields("uid_art_son5_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son5_hj1
            myrs.Fields("uid_art_son6_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son6_hj1
            myrs.Fields("uid_art_son7_hj1") = mcolSchuelerNoten.Item(intCount).uid_art_son7_hj1
            myrs.Fields("uid_art_son1_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son1_hj2
            myrs.Fields("uid_art_son2_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son2_hj2
            myrs.Fields("uid_art_son3_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son3_hj2
            myrs.Fields("uid_art_son4_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son4_hj2
            myrs.Fields("uid_art_son5_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son5_hj2
            myrs.Fields("uid_art_son6_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son6_hj2
            myrs.Fields("uid_art_son7_hj2") = mcolSchuelerNoten.Item(intCount).uid_art_son7_hj2
            
            myrs.Fields("dat_sa1_hj1") = mcolSchuelerNoten.Item(intCount).dat_sa1_hj1
            myrs.Fields("dat_sa2_hj1") = mcolSchuelerNoten.Item(intCount).dat_sa2_hj1
            myrs.Fields("dat_sa1_hj2") = mcolSchuelerNoten.Item(intCount).dat_sa1_hj2
            myrs.Fields("dat_sa2_hj2") = mcolSchuelerNoten.Item(intCount).dat_sa2_hj2
            myrs.Fields("dat_son1_hj1") = mcolSchuelerNoten.Item(intCount).dat_son1_hj1
            myrs.Fields("dat_son2_hj1") = mcolSchuelerNoten.Item(intCount).dat_son2_hj1
            myrs.Fields("dat_son3_hj1") = mcolSchuelerNoten.Item(intCount).dat_son3_hj1
            myrs.Fields("dat_son4_hj1") = mcolSchuelerNoten.Item(intCount).dat_son4_hj1
            myrs.Fields("dat_son5_hj1") = mcolSchuelerNoten.Item(intCount).dat_son5_hj1
            myrs.Fields("dat_son6_hj1") = mcolSchuelerNoten.Item(intCount).dat_son6_hj1
            myrs.Fields("dat_son7_hj1") = mcolSchuelerNoten.Item(intCount).dat_son7_hj1
            myrs.Fields("dat_son1_hj2") = mcolSchuelerNoten.Item(intCount).dat_son1_hj2
            myrs.Fields("dat_son2_hj2") = mcolSchuelerNoten.Item(intCount).dat_son2_hj2
            myrs.Fields("dat_son3_hj2") = mcolSchuelerNoten.Item(intCount).dat_son3_hj2
            myrs.Fields("dat_son4_hj2") = mcolSchuelerNoten.Item(intCount).dat_son4_hj2
            myrs.Fields("dat_son5_hj2") = mcolSchuelerNoten.Item(intCount).dat_son5_hj2
            myrs.Fields("dat_son6_hj2") = mcolSchuelerNoten.Item(intCount).dat_son6_hj2
            myrs.Fields("dat_son7_hj2") = mcolSchuelerNoten.Item(intCount).dat_son7_hj2
            
            myrs.Fields("gew_sa1_hj1") = mcolSchuelerNoten.Item(intCount).gew_sa1_hj1
            myrs.Fields("gew_sa2_hj1") = mcolSchuelerNoten.Item(intCount).gew_sa2_hj1
            myrs.Fields("gew_sa1_hj2") = mcolSchuelerNoten.Item(intCount).gew_sa1_hj2
            myrs.Fields("gew_sa2_hj2") = mcolSchuelerNoten.Item(intCount).gew_sa2_hj2
            myrs.Fields("gew_son1_hj1") = mcolSchuelerNoten.Item(intCount).gew_son1_hj1
            myrs.Fields("gew_son2_hj1") = mcolSchuelerNoten.Item(intCount).gew_son2_hj1
            myrs.Fields("gew_son3_hj1") = mcolSchuelerNoten.Item(intCount).gew_son3_hj1
            myrs.Fields("gew_son4_hj1") = mcolSchuelerNoten.Item(intCount).gew_son4_hj1
            myrs.Fields("gew_son5_hj1") = mcolSchuelerNoten.Item(intCount).gew_son5_hj1
            myrs.Fields("gew_son6_hj1") = mcolSchuelerNoten.Item(intCount).gew_son6_hj1
            myrs.Fields("gew_son7_hj1") = mcolSchuelerNoten.Item(intCount).gew_son7_hj1
            myrs.Fields("gew_son1_hj2") = mcolSchuelerNoten.Item(intCount).gew_son1_hj2
            myrs.Fields("gew_son2_hj2") = mcolSchuelerNoten.Item(intCount).gew_son2_hj2
            myrs.Fields("gew_son3_hj2") = mcolSchuelerNoten.Item(intCount).gew_son3_hj2
            myrs.Fields("gew_son4_hj2") = mcolSchuelerNoten.Item(intCount).gew_son4_hj2
            myrs.Fields("gew_son5_hj2") = mcolSchuelerNoten.Item(intCount).gew_son5_hj2
            myrs.Fields("gew_son6_hj2") = mcolSchuelerNoten.Item(intCount).gew_son6_hj2
            myrs.Fields("gew_son7_hj2") = mcolSchuelerNoten.Item(intCount).gew_son7_hj2
            
            myrs.Fields("n_sa1_hj1") = mcolSchuelerNoten.Item(intCount).n_sa1_hj1
            myrs.Fields("n_sa2_hj1") = mcolSchuelerNoten.Item(intCount).n_sa2_hj1
            myrs.Fields("n_sa1_hj2") = mcolSchuelerNoten.Item(intCount).n_sa1_hj2
            myrs.Fields("n_sa2_hj2") = mcolSchuelerNoten.Item(intCount).n_sa2_hj2
            myrs.Fields("n_son1_hj1") = mcolSchuelerNoten.Item(intCount).n_son1_hj1
            myrs.Fields("n_son2_hj1") = mcolSchuelerNoten.Item(intCount).n_son2_hj1
            myrs.Fields("n_son3_hj1") = mcolSchuelerNoten.Item(intCount).n_son3_hj1
            myrs.Fields("n_son4_hj1") = mcolSchuelerNoten.Item(intCount).n_son4_hj1
            myrs.Fields("n_son5_hj1") = mcolSchuelerNoten.Item(intCount).n_son5_hj1
            myrs.Fields("n_son6_hj1") = mcolSchuelerNoten.Item(intCount).n_son6_hj1
            myrs.Fields("n_son7_hj1") = mcolSchuelerNoten.Item(intCount).n_son7_hj1
            myrs.Fields("n_son1_hj2") = mcolSchuelerNoten.Item(intCount).n_son1_hj2
            myrs.Fields("n_son2_hj2") = mcolSchuelerNoten.Item(intCount).n_son2_hj2
            myrs.Fields("n_son3_hj2") = mcolSchuelerNoten.Item(intCount).n_son3_hj2
            myrs.Fields("n_son4_hj2") = mcolSchuelerNoten.Item(intCount).n_son4_hj2
            myrs.Fields("n_son5_hj2") = mcolSchuelerNoten.Item(intCount).n_son5_hj2
            myrs.Fields("n_son6_hj2") = mcolSchuelerNoten.Item(intCount).n_son6_hj2
            myrs.Fields("n_son7_hj2") = mcolSchuelerNoten.Item(intCount).n_son7_hj2
            myrs.Fields("geloescht") = mcolSchuelerNoten.Item(intCount).geloescht
            myrs.Update
            myrs.MoveFirst
        End If
    Next intCount

End Function
Public Function KlasseNotenLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As ADODB.Recordset
'Funktion liest die Noteneinstellungsdaten der Klasse
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenFachKlassengruppe.* " & _
                "FROM TabNotenFachKlassengruppe " & _
                "WHERE (((TabNotenFachKlassengruppe.uid_fach)=" & lngFachUid & ") AND ((TabNotenFachKlassengruppe.uid_klassengruppe)=" & lngKlassengruppenUid & ")) "
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set KlasseNotenLesen = myrs
    Set myrs = Nothing

End Function
Public Function KlasseNotenSpeichern(mclsKlasseNoten As clsKlasseNoten) As Boolean
'Funktion speichert die Noteneinstellungen in der Tabelle NotenFachKlassengruppe
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenFachKlassengruppe.* " & _
                "FROM TabNotenFachKlassengruppe  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    myrs.Find "uid=" & mclsKlasseNoten.uid
        ' wurde der Eintrag gefunden?
        If Not myrs.EOF Then
            myrs.Fields("anz_sa_hj1") = mclsKlasseNoten.anz_sa_hj1
            myrs.Fields("anz_sa_hj2") = mclsKlasseNoten.anz_sa_hj2
            
            myrs.Fields("anz_son_hj1") = mclsKlasseNoten.anz_son_hj1
            myrs.Fields("anz_son_hj2") = mclsKlasseNoten.anz_son_hj2
            
            myrs.Fields("uid_art_son1_hj1") = mclsKlasseNoten.uid_art_son1_hj1
            myrs.Fields("uid_art_son2_hj1") = mclsKlasseNoten.uid_art_son2_hj1
            myrs.Fields("uid_art_son3_hj1") = mclsKlasseNoten.uid_art_son3_hj1
            myrs.Fields("uid_art_son4_hj1") = mclsKlasseNoten.uid_art_son4_hj1
            myrs.Fields("uid_art_son5_hj1") = mclsKlasseNoten.uid_art_son5_hj1
            myrs.Fields("uid_art_son6_hj1") = mclsKlasseNoten.uid_art_son6_hj1
            myrs.Fields("uid_art_son7_hj1") = mclsKlasseNoten.uid_art_son7_hj1
            
            myrs.Fields("uid_art_son1_hj2") = mclsKlasseNoten.uid_art_son1_hj2
            myrs.Fields("uid_art_son2_hj2") = mclsKlasseNoten.uid_art_son2_hj2
            myrs.Fields("uid_art_son3_hj2") = mclsKlasseNoten.uid_art_son3_hj2
            myrs.Fields("uid_art_son4_hj2") = mclsKlasseNoten.uid_art_son4_hj2
            myrs.Fields("uid_art_son5_hj2") = mclsKlasseNoten.uid_art_son5_hj2
            myrs.Fields("uid_art_son6_hj2") = mclsKlasseNoten.uid_art_son6_hj2
            myrs.Fields("uid_art_son7_hj2") = mclsKlasseNoten.uid_art_son7_hj2
            
            myrs.Fields("dat_sa1_hj1") = mclsKlasseNoten.dat_sa1_hj1
            myrs.Fields("dat_sa2_hj1") = mclsKlasseNoten.dat_sa2_hj1
            myrs.Fields("dat_sa1_hj2") = mclsKlasseNoten.dat_sa1_hj2
            myrs.Fields("dat_sa2_hj2") = mclsKlasseNoten.dat_sa2_hj2
            
            myrs.Fields("dat_son1_hj1") = mclsKlasseNoten.dat_son1_hj1
            myrs.Fields("dat_son2_hj1") = mclsKlasseNoten.dat_son2_hj1
            myrs.Fields("dat_son3_hj1") = mclsKlasseNoten.dat_son3_hj1
            myrs.Fields("dat_son4_hj1") = mclsKlasseNoten.dat_son4_hj1
            myrs.Fields("dat_son5_hj1") = mclsKlasseNoten.dat_son5_hj1
            myrs.Fields("dat_son6_hj1") = mclsKlasseNoten.dat_son6_hj1
            myrs.Fields("dat_son7_hj1") = mclsKlasseNoten.dat_son7_hj1
            
            myrs.Fields("dat_son1_hj2") = mclsKlasseNoten.dat_son1_hj2
            myrs.Fields("dat_son2_hj2") = mclsKlasseNoten.dat_son2_hj2
            myrs.Fields("dat_son3_hj2") = mclsKlasseNoten.dat_son3_hj2
            myrs.Fields("dat_son4_hj2") = mclsKlasseNoten.dat_son4_hj2
            myrs.Fields("dat_son5_hj2") = mclsKlasseNoten.dat_son5_hj2
            myrs.Fields("dat_son6_hj2") = mclsKlasseNoten.dat_son6_hj2
            myrs.Fields("dat_son7_hj2") = mclsKlasseNoten.dat_son7_hj2
            
            myrs.Fields("gew_son1_hj1") = mclsKlasseNoten.gew_son1_hj1
            myrs.Fields("gew_son2_hj1") = mclsKlasseNoten.gew_son2_hj1
            myrs.Fields("gew_son3_hj1") = mclsKlasseNoten.gew_son3_hj1
            myrs.Fields("gew_son4_hj1") = mclsKlasseNoten.gew_son4_hj1
            myrs.Fields("gew_son5_hj1") = mclsKlasseNoten.gew_son5_hj1
            myrs.Fields("gew_son6_hj1") = mclsKlasseNoten.gew_son6_hj1
            myrs.Fields("gew_son7_hj1") = mclsKlasseNoten.gew_son7_hj1
            
            myrs.Fields("gew_son1_hj2") = mclsKlasseNoten.gew_son1_hj2
            myrs.Fields("gew_son2_hj2") = mclsKlasseNoten.gew_son2_hj2
            myrs.Fields("gew_son3_hj2") = mclsKlasseNoten.gew_son3_hj2
            myrs.Fields("gew_son4_hj2") = mclsKlasseNoten.gew_son4_hj2
            myrs.Fields("gew_son5_hj2") = mclsKlasseNoten.gew_son5_hj2
            myrs.Fields("gew_son6_hj2") = mclsKlasseNoten.gew_son6_hj2
            myrs.Fields("gew_son7_hj2") = mclsKlasseNoten.gew_son7_hj2
            
            myrs.Update
            myrs.MoveFirst
        End If

End Function

Public Function SchuelerHalbjahrLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As ADODB.Recordset
'Funktion liest die Halbjahresdaten des Schülers und die aktuelle Jahrgangsstufe seiner Klassengruppe und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT  TabNotenHalbjahrSchuelerFach.*, TabKlassengruppe.uid_jahrgangsstufe " & _
                "FROM (TabKlassengruppe INNER JOIN TabNotenHalbjahrSchuelerFach ON TabKlassengruppe.uid = TabNotenHalbjahrSchuelerFach.uid_klassengruppe) INNER JOIN TabSchueler ON (TabSchueler.uid = TabNotenHalbjahrSchuelerFach.uid_schueler) AND (TabKlassengruppe.uid = TabSchueler.uid_klassengruppe) " & _
                "WHERE (((TabNotenHalbjahrSchuelerFach.uid_fach)=" & lngFachUid & ") AND ((TabNotenHalbjahrSchuelerFach.uid_klassengruppe)=" & lngKlassengruppenUid & ")) " & _
                "ORDER BY TabSchueler.nachname, TabSchueler.rufname"
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set SchuelerHalbjahrLesen = myrs
    Set myrs = Nothing

End Function

Public Function SchuelerHalbjahrLesenKurz(lngSchuelerUid As Long) As ADODB.Recordset
'Funktion liest die Halbjahresdaten des Schülers und die aktuelle Jahrgangsstufe seiner Klassengruppe und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT  TabNotenHalbjahrSchuelerFach.*, TabKlassengruppe.uid_jahrgangsstufe " & _
                "FROM TabKlassengruppe INNER JOIN TabNotenHalbjahrSchuelerFach ON TabKlassengruppe.uid = TabNotenHalbjahrSchuelerFach.uid_klassengruppe " & _
                "WHERE TabNotenHalbjahrSchuelerFach.uid_schueler=" & lngSchuelerUid
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set SchuelerHalbjahrLesenKurz = myrs
    Set myrs = Nothing

End Function


Public Function FaecherHalbjahrLesen(lngSchuelerUid As Long, lngKlassengruppenUid As Long) As ADODB.Recordset
'Funktion liest die Halbjahresdaten der Fächer eines Schülers und die aktuelle Jahrgangsstufe und Schulart seiner Klassengruppe und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT  TabNotenHalbjahrSchuelerFach.*, TabKlassengruppe.uid_jahrgangsstufe, TabKlassengruppe.uid_schulart, TabFach.bezeichnung_lang " & _
                "FROM TabFach INNER JOIN (TabKlassengruppe INNER JOIN TabNotenHalbjahrSchuelerFach ON TabKlassengruppe.uid = TabNotenHalbjahrSchuelerFach.uid_klassengruppe) ON TabFach.uid = TabNotenHalbjahrSchuelerFach.uid_fach " & _
                "WHERE (((TabNotenHalbjahrSchuelerFach.uid_schueler)=" & lngSchuelerUid & ") AND ((TabNotenHalbjahrSchuelerFach.uid_klassengruppe)=" & lngKlassengruppenUid & ") AND ((TabNotenHalbjahrSchuelerFach.geloescht)=0)) " & _
                "ORDER BY TabFach.sort_uid"
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set FaecherHalbjahrLesen = myrs
    Set myrs = Nothing

End Function

Public Function SchuelerHalbjahrSpeichern() As Boolean
'Funktion speichert die Daten der Collection in der Tabelle SchülerHalbjahr abhängig von der Jahrgangsstufe in der der Schüler gerade ist
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenHalbjahrSchuelerFach.* " & _
                "FROM TabNotenHalbjahrSchuelerFach " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    For intCount = 1 To mcolSchuelerHalbjahr.Count
        myrs.Find "uid=" & mcolSchuelerHalbjahr.Item(intCount).uid
        ' wurde der Schüler gefunden?
        If Not myrs.EOF Then
            Select Case mcolSchuelerHalbjahr.Item(intCount).jahrgangsstufe_uid
                Case 1, 5, 6, 7 'Vorklassen und Vorkurs
                    myrs.Fields("n_vkl_hj1") = mcolSchuelerHalbjahr.Item(intCount).n_vkl_hj1
                    myrs.Fields("n_vkl_hj2") = mcolSchuelerHalbjahr.Item(intCount).n_vkl_hj2
                Case 2  '11. Klasse
                    myrs.Fields("n_11_hj1") = mcolSchuelerHalbjahr.Item(intCount).n_11_hj1
                    myrs.Fields("n_11_hj2") = mcolSchuelerHalbjahr.Item(intCount).n_11_hj2
                    myrs.Fields("n_11_hj1_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_11_hj1
                    myrs.Fields("n_11_hj2_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_11_hj2
                Case 3  '12.Klasse
                    myrs.Fields("n_12_hj1") = mcolSchuelerHalbjahr.Item(intCount).n_12_hj1
                    myrs.Fields("n_12_hj2") = mcolSchuelerHalbjahr.Item(intCount).n_12_hj2
                    myrs.Fields("n_12_hj1_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_12_hj1
                    myrs.Fields("n_12_hj2_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_12_hj2
                Case 4  '13.Klasse
                    myrs.Fields("n_13_hj1") = mcolSchuelerHalbjahr.Item(intCount).n_13_hj1
                    myrs.Fields("n_13_hj2") = mcolSchuelerHalbjahr.Item(intCount).n_13_hj2
                    myrs.Fields("n_13_hj1_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_13_hj1
                    myrs.Fields("n_13_hj2_abgewaehlt") = mcolSchuelerHalbjahr.Item(intCount).ein_13_hj2
            End Select
            myrs.Update
            myrs.MoveFirst
        End If
    Next intCount

End Function

Public Function SchuelerHalbjahrSpeichernKurz() As Boolean
'Funktion speichert die Daten der Collection in der Tabelle SchülerHalbjahr
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenHalbjahrSchuelerFach.* " & _
                "FROM TabNotenHalbjahrSchuelerFach " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    For intCount = 1 To mcolSchuelerHalbjahrKurz.Count
        myrs.Find "uid=" & mcolSchuelerHalbjahrKurz.Item(intCount).uid
        ' wurde der Schüler gefunden?
        If Not myrs.EOF Then
            myrs.Fields("uid_klassengruppe") = mcolSchuelerHalbjahrKurz.Item(intCount).klassengruppe_uid
            myrs.Fields("geloescht") = mcolSchuelerHalbjahrKurz.Item(intCount).geloescht
            myrs.Update
            myrs.MoveFirst
        End If
    Next intCount

End Function

Public Function KlasseFaecherLesen(lngKlassengruppenUid As Long) As ADODB.Recordset
'Funktion liest die Daten der Fächer, die einer Klassengruppe zugeordnet sind und gib den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenFachKlassengruppe.uid, TabNotenFachKlassengruppe.uid_fach, TabNotenFachKlassengruppe.uid_klassengruppe, TabNotenFachKlassengruppe.geloescht, TabFach.bezeichnung_kurz " & _
                "FROM TabFach INNER JOIN TabNotenFachKlassengruppe ON TabFach.uid = TabNotenFachKlassengruppe.uid_fach " & _
                "WHERE (TabNotenFachKlassengruppe.uid_klassengruppe=" & lngKlassengruppenUid & ") AND (TabNotenFachKlassengruppe.geloescht=0)" & _
                "ORDER BY TabFach.sort_uid"
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set KlasseFaecherLesen = myrs
    Set myrs = Nothing

End Function
Public Function KlasseFaecherSchreiben() As Boolean
' schreibe die geänderten Daten zurück auch in der Tabelle TabNotenSchuelerFachKlassengruppe und TabNotenHalbjahrSchuelerFach
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String
Dim myrs1 As ADODB.Recordset
Dim myrs2 As ADODB.Recordset
Dim strSql1 As String
Dim strSql2 As String
Dim intCount1 As Integer

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenFachKlassengruppe.* " & _
                "FROM TabNotenFachKlassengruppe " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    For intCount = 1 To mcolKlasseFaecher.Count
        If mcolKlasseFaecher.Item(intCount).geloescht = True Then
            myrs.Find "uid=" & mcolKlasseFaecher.Item(intCount).uid
            ' wurde der Eintrag gefunden?
            If Not myrs.EOF Then
                myrs.Fields("geloescht") = mcolKlasseFaecher.Item(intCount).geloescht
                myrs.Update
                myrs.MoveFirst
            End If
            Set myrs1 = New ADODB.Recordset
            strSql1 = "SELECT TabNotenSchuelerFachKlassengruppe.* " & _
                        "FROM TabNotenSchuelerFachKlassengruppe " & _
                        "WHERE uid_fach = " & mcolKlasseFaecher.Item(intCount).fach_uid & " AND " & _
                        "uid_klassengruppe = " & mcolKlasseFaecher.Item(intCount).klassengruppe_uid
            myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
            If myrs1.RecordCount > 0 Then
                myrs1.MoveFirst
                For intCount1 = 1 To myrs1.RecordCount
                    myrs1.Fields("geloescht") = True
                    myrs1.Update
                    myrs1.MoveNext
                Next intCount1
            End If
            Set myrs2 = New ADODB.Recordset
            strSql2 = "SELECT TabNotenHalbjahrSchuelerFach.* " & _
                        "FROM TabNotenHalbjahrSchuelerFach " & _
                        "WHERE uid_fach = " & mcolKlasseFaecher.Item(intCount).fach_uid & " AND " & _
                        "uid_klassengruppe = " & mcolKlasseFaecher.Item(intCount).klassengruppe_uid
            myrs2.Open strSql2, myconn, adOpenKeyset, adLockOptimistic
            If myrs2.RecordCount > 0 Then
                myrs2.MoveFirst
                For intCount1 = 1 To myrs2.RecordCount
                    myrs2.Fields("geloescht") = True
                    myrs2.Update
                    myrs2.MoveNext
                Next intCount1
            End If
        End If
    Next intCount
    Set myrs = Nothing
    
End Function
Public Function SchülerdatenLesen(lngSchuelerUid As Long) As ADODB.Recordset
'Funktion liest die Daten des Schülers und liefert den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.* " & _
                "FROM TabSchueler " & _
                "WHERE ((TabSchueler.uid)=" & lngSchuelerUid & ") "
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set SchülerdatenLesen = myrs
    Set myrs = Nothing
    
End Function
Public Function SchülerdatenSpeichern(myclsSchuelerdaten As clsSchülerdaten) As Boolean
'Funktion speichert die Schülerdaten in der Tabelle TabSchueler
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.* " & _
                "FROM TabSchueler  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    myrs.Find "uid=" & myclsSchuelerdaten.uid
        ' wurde der Eintrag gefunden?
        If Not myrs.EOF Then
            myrs.Fields("nachname") = myclsSchuelerdaten.nachname
            myrs.Fields("rufname") = myclsSchuelerdaten.rufname
            myrs.Fields("vornamen") = myclsSchuelerdaten.vornamen
            myrs.Fields("gebdat") = myclsSchuelerdaten.geburtsdatum
            myrs.Fields("gebort") = myclsSchuelerdaten.geburtsort
            myrs.Fields("gebland") = myclsSchuelerdaten.geburtsland
            myrs.Fields("bekenntnis") = myclsSchuelerdaten.bekenntnis
            myrs.Fields("religion_unterricht") = myclsSchuelerdaten.religion
            myrs.Fields("ausgeschieden") = myclsSchuelerdaten.ausgeschieden
            If Not IsNull(myclsSchuelerdaten.ausgeschiedenam) And Not myclsSchuelerdaten.ausgeschiedenam = "" Then
                myrs.Fields("ausgeschieden_am") = myclsSchuelerdaten.ausgeschiedenam
            End If
            myrs.Fields("uid_geschlecht") = myclsSchuelerdaten.geschlecht
            myrs.Fields("uid_klassengruppe") = myclsSchuelerdaten.klassengruppe
            If Not IsNull(myclsSchuelerdaten.eintritt) And Not myclsSchuelerdaten.eintritt = "" Then
                myrs.Fields("eintritt_schule") = myclsSchuelerdaten.eintritt
            End If
            If Not IsNull(myclsSchuelerdaten.probezeit) And Not myclsSchuelerdaten.probezeit = "" Then
                myrs.Fields("probezeit_bis") = myclsSchuelerdaten.probezeit
            End If
            myrs.Fields("notenschutz") = myclsSchuelerdaten.notenschutz
            myrs.Update
            myrs.MoveFirst
        End If

End Function

Public Function LehrerdatenLesen(lngLehrerUid As Long) As ADODB.Recordset
'Funktion liest die Daten des Lehrers und liefert den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabLehrer.* " & _
                "FROM TabLehrer " & _
                "WHERE ((TabLehrer.uid)=" & lngLehrerUid & ") "
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set LehrerdatenLesen = myrs
    Set myrs = Nothing
    
End Function
Public Function LehrerdatenSpeichern(myclsLehrerdaten As clsLehrerDaten) As Boolean
'Funktion speichert die Lehrerdaten in der Tabelle TabLehrer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabLehrer.* " & _
                "FROM TabLehrer  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    myrs.Find "uid=" & myclsLehrerdaten.uid
        ' wurde der Eintrag gefunden?
        If Not myrs.EOF Then
            myrs.Fields("kuerzel") = myclsLehrerdaten.kuerzel
            myrs.Fields("familienname") = myclsLehrerdaten.nachname
            myrs.Fields("rufname") = myclsLehrerdaten.rufname
            myrs.Fields("amt") = myclsLehrerdaten.amt
            myrs.Fields("titel") = myclsLehrerdaten.titel
            myrs.Fields("uid_geschlecht") = myclsLehrerdaten.geschlecht
            myrs.Fields("email") = myclsLehrerdaten.email
            myrs.Fields("schulleitung") = myclsLehrerdaten.schulleitung
            myrs.Update
            myrs.MoveFirst
        End If

End Function
Public Function KlassendatenLesen(lngKlasseUid As Long) As ADODB.Recordset
'Funktion liest die Daten der Klasse und liefert den Recordset zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlasse.* " & _
                "FROM TabKlasse " & _
                "WHERE ((TabKlasse.uid)=" & lngKlasseUid & ") "
    
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    Set KlassendatenLesen = myrs
    Set myrs = Nothing
    
End Function
Public Function KlassendatenSpeichern(myclsKlassendaten As clsKlasseDaten) As Boolean
'Funktion speichert die Klassendaten in der Tabelle TabKlasse
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlasse.* " & _
                "FROM TabKlasse  " & _
                "ORDER BY uid"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    myrs.MoveFirst

    myrs.Find "uid=" & myclsKlassendaten.uid
        ' wurde der Eintrag gefunden?
        If Not myrs.EOF Then
            myrs.Fields("bezeichnung") = myclsKlassendaten.bezeichnung
            myrs.Fields("klassleitung") = myclsKlassendaten.klassleitung
            myrs.Fields("zeugnisunterzeichnung") = myclsKlassendaten.zeugnisunterzeichner
            myrs.Update
            myrs.MoveFirst
        End If

End Function

' Collection SchülerNoten
Public Property Set SchuelerNoten(ByVal vNewValue As Variant)
    Set mcolSchuelerNoten = vNewValue
End Property

' Collection SchülerNotenKurz
Public Property Set SchuelerNotenKurz(ByVal vNewValue As Variant)
    Set mcolSchuelerNotenKurz = vNewValue
End Property

' Collection SchülerHalbjahresleistungen
Public Property Set SchuelerHalbjahr(ByVal vNewValue As Variant)
    Set mcolSchuelerHalbjahr = vNewValue
End Property
' Collection SchülerHalbjahresleistungenKurz
Public Property Set SchuelerHalbjahrKurz(ByVal vNewValue As Variant)
    Set mcolSchuelerHalbjahrKurz = vNewValue
End Property
' Collection KlasseFaecher
Public Property Set KlasseFaecher(ByVal vNewValue As Variant)
    Set mcolKlasseFaecher = vNewValue
End Property

' Collection Schüler
Public Property Set Schueler(ByVal vNewValue As Variant)
    Set myColSchueler = vNewValue
End Property

' Collection Lehrer
Public Property Set Lehrer(ByVal vNewValue As Variant)
    Set myColLehrer = vNewValue
End Property

' Collection Schule
Public Property Set Schule(ByVal vNewValue As Variant)
    Set myColSchule = vNewValue
End Property

' Collection Halbjahresnoten
Public Property Set Halbjahr(ByVal vNewValue As Variant)
    Set mcolHalbjahr = vNewValue
End Property

' Collection Halbjahresnoten Komplett
Public Property Set HalbjahrKomplett(ByVal vNewValue As Variant)
    Set mcolHalbjahrKomplett = vNewValue
End Property

Public Sub SchuelerImportieren()
'hier werden die Schüler in die Datenbank importiert
Dim lngCountSchueler As Long    'Anzahl der Schüler
Dim lngSchuelerUid As Long      'Uid des Schülers
Dim lngKlasseUid As Long        'Uid der Klasse
Dim lngKlassengruppeUid As Long 'Uid der Klassengruppe
Dim lngSchuelerKlassengruppe As Long    'Ist Schüler in Klassengruppe? 0 = Nein, 1 = Ja
Dim lngKlassengruppeUidAlt As Long  'Uid der alten Klassengruppe eines Schülers
Dim lngAusbUidAlt As Long           'Uid der alten Ausbildungsrichtung eines Schülers
Dim lngJahrgangUidAlt As Long       'Uid der alten Jahrgangsstufe eines Schülers
    
    On Error GoTo Err_Schueler_Importieren
    
    SysCmd acSysCmdInitMeter, "Schüler importieren: ", myColSchueler.Count
    
    For lngCountSchueler = 1 To myColSchueler.Count
    
        SysCmd acSysCmdUpdateMeter, lngCountSchueler
        
        lngSchuelerUid = SchuelerVorhanden(myColSchueler.Item(lngCountSchueler).schuelernummer)
        'ist der Schüler vorhanden?
        If lngSchuelerUid = 0 Then  'Schüler nicht da.
            lngKlasseUid = KlasseVorhanden(myColSchueler.Item(lngCountSchueler).klasse)
            If lngKlasseUid = 0 Then    'Klasse ist nicht vorhanden, dann anlegen, auch Klassengruppe
                lngKlasseUid = KlasseAnlegen(myColSchueler.Item(lngCountSchueler).klasse)
                lngKlassengruppeUid = KlassengruppeAnlegen(myColSchueler.Item(lngCountSchueler).klasse, lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung_kurz)
                If KlassenNotenAnlegen(lngKlassengruppeUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart) = 0 Then
                    Err.Raise 514, "Notenverwaltung FOSBOS", "Kein Standardeintrag für Ausbildungsrichtung: " & myColSchueler.Item(lngCountSchueler).ausbildungsrichtung & _
                                    ", Jahrgangsstufe:  " & myColSchueler.Item(lngCountSchueler).jahrgangsstufe & " und Schulart: " & myColSchueler.Item(lngCountSchueler).schulart
                Else
                End If
            Else    'Klasse ist vorhanden
                lngKlassengruppeUid = KlassengruppeVorhanden(lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart)
                If lngKlassengruppeUid = 0 Then 'Klassengruppe nicht vorhanden, dann anlegen
                    lngKlassengruppeUid = KlassengruppeAnlegen(myColSchueler.Item(lngCountSchueler).klasse, lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung_kurz)
                    If KlassenNotenAnlegen(lngKlassengruppeUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart) = 0 Then
                        Err.Raise 514, "Notenverwaltung FOSBOS", "Kein Standardeintrag für Ausbildungsrichtung: " & myColSchueler.Item(lngCountSchueler).ausbildungsrichtung & _
                                        ", Jahrgangsstufe:  " & myColSchueler.Item(lngCountSchueler).jahrgangsstufe & " und Schulart: " & myColSchueler.Item(lngCountSchueler).schulart
                    Else
                    End If
                End If
            End If
            ' den Schüler neu anlegen mitsamt allen Fächern und Noten
            lngSchuelerUid = SchuelerAnlegen(myColSchueler.Item(lngCountSchueler).schuelernummer, myColSchueler.Item(lngCountSchueler).nachname, myColSchueler.Item(lngCountSchueler).rufname, myColSchueler.Item(lngCountSchueler).vornamen, _
                                                myColSchueler.Item(lngCountSchueler).geburtsdatum, myColSchueler.Item(lngCountSchueler).geburtsort, myColSchueler.Item(lngCountSchueler).geburtsland, _
                                                myColSchueler.Item(lngCountSchueler).bekenntnis, myColSchueler.Item(lngCountSchueler).religion, _
                                                lngKlassengruppeUid, myColSchueler.Item(lngCountSchueler).geschlecht, myColSchueler.Item(lngCountSchueler).eintrittdatum, myColSchueler.Item(lngCountSchueler).probezeit)
            If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
                Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
            End If
            If SchuelerNotenHalbjahrAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe) = 0 Then
                Err.Raise 516, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
            End If
            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                Else
                If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                    Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                End If
                If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                    Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                End If
            End If
            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                If SchuelerNotenSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                    Err.Raise 519, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                End If
                If SchuelerNotenHalbjahrSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                    Err.Raise 520, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                End If
            End If
        Else        ' der Schüler ist bereits da
            lngKlasseUid = KlasseVorhanden(myColSchueler.Item(lngCountSchueler).klasse)
            If lngKlasseUid = 0 Then    'Klasse ist nicht vorhanden, dann anlegen, auch Klassengruppe
                lngKlasseUid = KlasseAnlegen(myColSchueler.Item(lngCountSchueler).klasse)
                lngKlassengruppeUid = KlassengruppeAnlegen(myColSchueler.Item(lngCountSchueler).klasse, lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung_kurz)
                If KlassenNotenAnlegen(lngKlassengruppeUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart) = 0 Then
                    Err.Raise 514, "Notenverwaltung FOSBOS", "Kein Standardeintrag für Ausbildungsrichtung: " & myColSchueler.Item(lngCountSchueler).ausbildungsrichtung & _
                                    ", Jahrgangsstufe:  " & myColSchueler.Item(lngCountSchueler).jahrgangsstufe & " und Schulart: " & myColSchueler.Item(lngCountSchueler).schulart
                Else
                End If
            Else
                lngKlassengruppeUid = KlassengruppeVorhanden(lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart)
                If lngKlassengruppeUid = 0 Then 'Klassengruppe nicht vorhanden, dann anlegen
                    lngKlassengruppeUid = KlassengruppeAnlegen(myColSchueler.Item(lngCountSchueler).klasse, lngKlasseUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung_kurz)
                    If KlassenNotenAnlegen(lngKlassengruppeUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).ausbildungsrichtung, myColSchueler.Item(lngCountSchueler).schulart) = 0 Then
                        Err.Raise 514, "Notenverwaltung FOSBOS", "Kein Standardeintrag für Ausbildungsrichtung: " & myColSchueler.Item(lngCountSchueler).ausbildungsrichtung & _
                                        ", Jahrgangsstufe:  " & myColSchueler.Item(lngCountSchueler).jahrgangsstufe & " und Schulart: " & myColSchueler.Item(lngCountSchueler).schulart
                    Else
                    End If
                End If
            End If
            ' Klassengruppe und Klasse sind beide da jetzt prüfen ob der Schüler bereits in dieser Klassengruppe ist
            If SchuelerInKlassengruppe(lngSchuelerUid, lngKlassengruppeUid) = 0 Then       'nicht da
                ' Ermittle die momentan gespeicherte Klassengruppe des Schülers
                lngKlassengruppeUidAlt = KlassengruppeEinesSchuelersErmitteln(lngSchuelerUid)
                If lngKlassengruppeUidAlt > 0 Then           'Schüler gefunden
                    lngJahrgangUidAlt = JahrgangsstufeErmitteln(lngKlassengruppeUidAlt)     'Alte Jahrgangsstufe des Schülers
                    lngAusbUidAlt = AusbildungsrichtungErmitteln(lngKlassengruppeUidAlt)    'Alte Ausbildungsrichtung des Schülers
                    If lngJahrgangUidAlt = myColSchueler.Item(lngCountSchueler).jahrgangsstufe Then
                        If lngAusbUidAlt = myColSchueler.Item(lngCountSchueler).ausbildungsrichtung Then
                        ' Jahrgangsstufe und Ausbildungsrichtung haben sich nicht geändert
                            ' Klassengruppe beim Schüler austauschen im Schülersatz
                            If KlassengruppeEinesSchuelersAendern(lngSchuelerUid, lngKlassengruppeUid) = 0 Then
                                Err.Raise 531, "Notenverwaltung FOSBOS", "Klassengruppe des Schülers " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname & " konnte nicht geändert werden!"
                            End If
                            ' die alten Notensätze löschen und neue anlegen
                            NotensaetzeEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
                                Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            ' im Halbjahresnotensatz die Klassengruppe austauschen und gefüllte Noten zurücksetzen, da der Schüler wiederholt.
                            NotensaetzeHalbjahrEinesSchuelersAendern lngSchuelerUid, lngKlassengruppeUid, lngKlassengruppeUidAlt, myColSchueler.Item(lngCountSchueler).jahrgangsstufe
                        Else
                        ' Jahrgangsstufe bleibt gleich und die Ausbildungsrichung ändert sich
                            ' Klassengruppe beim Schüler austauschen im Schülersatz
                            If KlassengruppeEinesSchuelersAendern(lngSchuelerUid, lngKlassengruppeUid) = 0 Then
                                Err.Raise 531, "Notenverwaltung FOSBOS", "Klassengruppe des Schülers " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname & " konnte nicht geändert werden!"
                            End If
                            ' die alten Notensätze löschen und neue anlegen
                            NotensaetzeEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
                                Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                If SchuelerNotenSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                    Err.Raise 519, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            ' die alten Halbjahressätze löschen und neue anlegen
                            NotensaetzeHalbjahrEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenHalbjahrAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe) = 0 Then
                                Err.Raise 516, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                If SchuelerNotenHalbjahrSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                    Err.Raise 520, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                        End If
                    Else
                        If lngAusbUidAlt = myColSchueler.Item(lngCountSchueler).ausbildungsrichtung Then
                        ' Jahrgangsstufe hat sich geändert, Ausbildungsrichtung bleibt
                            ' Klassengruppe beim Schüler austauschen im Schülersatz
                            If KlassengruppeEinesSchuelersAendern(lngSchuelerUid, lngKlassengruppeUid) = 0 Then
                                Err.Raise 531, "Notenverwaltung FOSBOS", "Klassengruppe des Schülers " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname & " konnte nicht geändert werden!"
                            End If
                            ' die alten Notensätze löschen und neue anlegen
                            NotensaetzeEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
                                Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                If SchuelerNotenSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                    Err.Raise 519, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            ' im Halbjahresnotensatz die Klassengruppe austauschen und die fehlenden Notensätze für die neue Jahrgangsstufe ergänzen wenn es sich
                            ' bei der neuen Jahrgangsstufe um die 12.FOS (Jahrgangsstufe = 3 und Schulart = 1) handelt, ansonsten die alten Sätze löschen und neue anlegen.
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 3 And myColSchueler.Item(lngCountSchueler).schulart = 1 Then
                            ' Klassengruppe tauschen und fehlende Sätze ergänzen
                                NotensaetzeHalbjahrEinesSchuelersErgaenzen lngSchuelerUid, lngKlassengruppeUid, lngKlassengruppeUidAlt
                                If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            Else
                            ' Alte löschen und neue anlegen
                                NotensaetzeHalbjahrEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                                If SchuelerNotenHalbjahrAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe) = 0 Then
                                    Err.Raise 516, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                                If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                    Else
                                    If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                        Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                    End If
                                End If
                                If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                    If SchuelerNotenHalbjahrSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                        Err.Raise 520, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                    End If
                                End If
                            End If
                        Else
                        ' Jahrgangsstufe hat sich geändert und Ausbildungsrichtung auch (geht nur bei Vorklasse)
                            ' Klassengruppe beim Schüler austauschen im Schülersatz
                            If KlassengruppeEinesSchuelersAendern(lngSchuelerUid, lngKlassengruppeUid) = 0 Then
                                Err.Raise 531, "Notenverwaltung FOSBOS", "Klassengruppe des Schülers " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname & " konnte nicht geändert werden!"
                            End If
                            ' die alten Notensätze löschen und neue anlegen
                            NotensaetzeEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
                                Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                If SchuelerNotenSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                    Err.Raise 519, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            ' die alten Halbjahressätze löschen und neue anlegen
                            NotensaetzeHalbjahrEinesSchuelersLoeschen lngSchuelerUid, lngKlassengruppeUidAlt
                            If SchuelerNotenHalbjahrAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe) = 0 Then
                                Err.Raise 516, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                            End If
                            If myColSchueler.Item(lngCountSchueler).jahrgangsstufe = 2 Or myColSchueler.Item(lngCountSchueler).jahrgangsstufe > 4 Then 'Religion
                                Else
                                If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).jahrgangsstufe, myColSchueler.Item(lngCountSchueler).religion) = 0 Then
                                    Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                            If myColSchueler.Item(lngCountSchueler).ausbildungsrichtung = 4 And myColSchueler.Item(lngCountSchueler).sprache > 0 Then   'Sprachunterricht Int. Wirtschaft
                                If SchuelerNotenHalbjahrSpracheAnlegen(lngKlassengruppeUid, lngSchuelerUid, myColSchueler.Item(lngCountSchueler).sprache) = 0 Then
                                    Err.Raise 520, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Sprachunterricht für den Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname
                                End If
                            End If
                        End If
                    End If
                Else
                    Err.Raise 530, "Notenverwaltung FOSBOS", "Schüler " & myColSchueler.Item(lngCountSchueler).rufname & " " & myColSchueler.Item(lngCountSchueler).nachname & " nicht gefunden!"
                End If
            Else
            ' Schüler bereits in Klassengruppe nix machen
            End If
        End If
    Next lngCountSchueler
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Schueler_Importieren:
    Exit Sub
    
Err_Schueler_Importieren:
    FehlermeldungAusgeben "Datenbank", "Schüler importieren", Err.Number, Err.Description
    Resume Exit_Schueler_Importieren
    
End Sub
Public Sub LehrerImportieren()
'hier werden die Lehrer in die Datenbank importiert
Dim lngCountLehrer As Long    'Anzahl der Lehrer
Dim lngLehrerUid As Long      'Uid des Lehrers

    On Error GoTo Err_Lehrer_Importieren
    
    SysCmd acSysCmdInitMeter, "Lehrer importieren: ", myColLehrer.Count
    
    For lngCountLehrer = 1 To myColLehrer.Count
    
        SysCmd acSysCmdUpdateMeter, lngCountLehrer
        lngLehrerUid = LehrerVorhanden(myColLehrer.Item(lngCountLehrer).kuerzel)
        If lngLehrerUid = 0 Then    'Lehrer nicht vorhanden, also anlegen
            lngLehrerUid = LehrerAnlegen(myColLehrer.Item(lngCountLehrer).kuerzel, myColLehrer.Item(lngCountLehrer).nachname, myColLehrer.Item(lngCountLehrer).rufname, _
                                            myColLehrer.Item(lngCountLehrer).amt, myColLehrer.Item(lngCountLehrer).titel, myColLehrer.Item(lngCountLehrer).geschlecht)
        Else                        'Daten updaten
            lngLehrerUid = LehrerAendern(lngLehrerUid, myColLehrer.Item(lngCountLehrer).kuerzel, myColLehrer.Item(lngCountLehrer).nachname, myColLehrer.Item(lngCountLehrer).rufname, _
                                            myColLehrer.Item(lngCountLehrer).amt, myColLehrer.Item(lngCountLehrer).titel, myColLehrer.Item(lngCountLehrer).geschlecht)
        End If
        
    Next lngCountLehrer
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Lehrer_Importieren:
    Exit Sub
    
Err_Lehrer_Importieren:
    FehlermeldungAusgeben "Datenbank", "Lehrer importieren", Err.Number, Err.Description
    Resume Exit_Lehrer_Importieren
    
End Sub

Public Sub SchuleImportieren()
'hier werden die Schulen in die Datenbank importiert
Dim lngCountSchule As Long    'Anzahl der Schulen
Dim lngSchulUid As Long      'Uid der Schule

    On Error GoTo Err_Schule_Importieren
    
    SysCmd acSysCmdInitMeter, "Schule importieren: ", myColSchule.Count
    
    For lngCountSchule = 1 To myColSchule.Count
    
        SysCmd acSysCmdUpdateMeter, lngCountSchule
        lngSchulUid = SchuleVorhanden(myColSchule.Item(lngCountSchule).schulnummer)
        If lngSchulUid = 0 Then    'Schule nicht vorhanden, also anlegen
            lngSchulUid = SchuleAnlegen(myColSchule.Item(lngCountSchule).schulnummer, myColSchule.Item(lngCountSchule).schulart, myColSchule.Item(lngCountSchule).schulname, myColSchule.Item(lngCountSchule).schulnamekurz)
        Else                        'Daten updaten
            lngSchulUid = SchuleAendern(lngSchulUid, myColSchule.Item(lngCountSchule).schulnummer, myColSchule.Item(lngCountSchule).schulart, myColSchule.Item(lngCountSchule).schulname, myColSchule.Item(lngCountSchule).schulnamekurz)
        End If
        
    Next lngCountSchule
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Schule_Importieren:
    Exit Sub
    
Err_Schule_Importieren:
    FehlermeldungAusgeben "Datenbank", "Schule importieren", Err.Number, Err.Description
    Resume Exit_Schule_Importieren
    
End Sub

Public Sub HalbjahresNotenImportieren(ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal lngJahrgangsstufeUid As Long)
'hier werden die Halbjahresnoten in die Datenbank importiert
Dim lngCountSchueler As Long    'Anzahl der Schüler
Dim lngSchuelerUid As Long      'Uid des Schülers
Dim lngNotenUid As Long         'Uid des Notenhalbjahressatzes

    On Error GoTo Err_Halbjahr_Importieren
    
    SysCmd acSysCmdInitMeter, "Halbjahresnoten importieren: ", mcolHalbjahr.Count
    
    For lngCountSchueler = 1 To mcolHalbjahr.Count
    
        SysCmd acSysCmdUpdateMeter, lngCountSchueler
        lngSchuelerUid = SchuelerVorhanden(mcolHalbjahr.Item(lngCountSchueler).schuelernummer)  'gibt es den Schüler?
        If lngSchuelerUid <> 0 Then
            lngNotenUid = SchuelerNotenHalbjahrVorhanden(lngSchuelerUid, lngFachUid, lngKlassengruppeUid)   'Notendatensatz vorhanden
            If lngNotenUid <> 0 Then
                lngNotenUid = SchuelerNotenHalbjahrAendern(lngNotenUid, lngJahrgangsstufeUid, mcolHalbjahr.Item(lngCountSchueler).halbjahr1, mcolHalbjahr.Item(lngCountSchueler).halbjahr2, mcolHalbjahr.Item(lngCountSchueler).schuelernummer)
            Else
                Err.Raise 531, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Keinen Halbjahressatz für den Schüler mit der Nummer " & mcolHalbjahr.Item(lngCountSchueler).schuelernummer & " gefunden!"
            End If
        Else
            Err.Raise 530, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Schüler mit der Nummer " & mcolHalbjahr.Item(lngCountSchueler).schuelernummer & " nicht gefunden!"
        End If
        
    Next lngCountSchueler
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Halbjahr_Importieren:
    Exit Sub
    
Err_Halbjahr_Importieren:
    FehlermeldungAusgeben "Datenbank", "Halbjahresleistungen importieren", Err.Number, Err.Description
    Resume Exit_Halbjahr_Importieren
    
End Sub

Public Sub HalbjahresNotenKomplettImportieren()
'hier werden die Halbjahresnoten in die Datenbank importiert
Dim lngCountSchueler As Long        'Anzahl der Schüler
Dim lngSchuelerUid As Long          'Uid des Schülers
Dim lngFachUid As Long              'Uid des Faches
Dim lngKlassengruppeUid As Long     'Uid der Klassengruppe
Dim lngJahrgangsstufeUid As Long    'Uid der Jahrgangsstufe
Dim lngNotenUid As Long             'Uid des Notenhalbjahressatzes

    On Error GoTo Err_Halbjahr_Importieren
    
    SysCmd acSysCmdInitMeter, "Halbjahresnoten importieren: ", mcolHalbjahrKomplett.Count
    
    For lngCountSchueler = 1 To mcolHalbjahrKomplett.Count
    
        SysCmd acSysCmdUpdateMeter, lngCountSchueler
        lngSchuelerUid = SchuelerVorhanden(mcolHalbjahrKomplett.Item(lngCountSchueler).schuelernummer)  'gibt es den Schüler?
        If lngSchuelerUid <> 0 Then
            lngFachUid = FachVorhanden(mcolHalbjahrKomplett.Item(lngCountSchueler).fach)                ' gibt es das Fach?
            If lngFachUid <> 0 Then
                lngKlassengruppeUid = SchuelerInKlasse(lngSchuelerUid, mcolHalbjahrKomplett.Item(lngCountSchueler).klasse)  ' ist der Schüler in der Klasse?
                If lngKlassengruppeUid <> 0 Then
                    lngNotenUid = SchuelerNotenHalbjahrVorhanden(lngSchuelerUid, lngFachUid, lngKlassengruppeUid)   'Notendatensatz vorhanden
                    If lngNotenUid <> 0 Then
                        lngJahrgangsstufeUid = JahrgangsstufeErmitteln(lngKlassengruppeUid)                         'Jahrgangsstufe ermitteln
                        lngNotenUid = SchuelerNotenHalbjahrAendern(lngNotenUid, lngJahrgangsstufeUid, mcolHalbjahrKomplett.Item(lngCountSchueler).halbjahr1, mcolHalbjahrKomplett.Item(lngCountSchueler).halbjahr2, mcolHalbjahrKomplett.Item(lngCountSchueler).schuelernummer)
                    Else
                        Err.Raise 531, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Keinen Halbjahressatz für den Schüler " & mcolHalbjahrKomplett.Item(lngCountSchueler).rufname & " " & mcolHalbjahrKomplett.Item(lngCountSchueler).name & " mit der Nummer " & mcolHalbjahr.Item(lngCountSchueler).schuelernummer & " gefunden!"
                    End If
                Else
                    Err.Raise 533, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Schüler " & mcolHalbjahrKomplett.Item(lngCountSchueler).rufname & " " & mcolHalbjahrKomplett.Item(lngCountSchueler).name & " mit der Nummer " & mcolHalbjahrKomplett.Item(lngCountSchueler).schuelernummer & " in der Klasse " & mcolHalbjahrKomplett.Item(lngCountSchueler).klasse & " nicht gefunden!"
                End If
            Else
                Err.Raise 532, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Fachkürzel " & mcolHalbjahrKomplett.Item(lngCountSchueler).fach & " nicht gefunden!"
            End If
        Else
            Err.Raise 530, "Notenverwaltung FOSBOS", "Fehler beim Import der Halbjahresleistungen: Schüler " & mcolHalbjahrKomplett.Item(lngCountSchueler).rufname & " " & mcolHalbjahrKomplett.Item(lngCountSchueler).name & " mit der Nummer " & mcolHalbjahrKomplett.Item(lngCountSchueler).schuelernummer & " nicht gefunden!"
        End If
        
    Next lngCountSchueler
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Halbjahr_Importieren:
    Exit Sub
    
Err_Halbjahr_Importieren:
    FehlermeldungAusgeben "Datenbank", "Halbjahresleistungen importieren", Err.Number, Err.Description
    Resume Exit_Halbjahr_Importieren
    
End Sub

Private Function SchuelerVorhanden(lngSchuelernummer As Long) As Long
' prüft ob es den Schüler bereits gibt, melde die Uid des Schülers oder 0 zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.uid FROM TabSchueler " & _
                "WHERE TabSchueler.schuelernummer = " & lngSchuelernummer & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        SchuelerVorhanden = myrs.Fields("uid")
    Else
        SchuelerVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function

Private Function KlasseVorhanden(strKlasse As String) As Long
' prüft, ob es die Klasse gibt, melde die Uid zurück oder 0
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlasse.uid FROM TabKlasse " & _
                "WHERE TabKlasse.bezeichnung = " & Chr(34) & strKlasse & Chr(34) & " AND TabKlasse.geloescht=FALSE"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        KlasseVorhanden = myrs.Fields("uid")
    Else
        KlasseVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function
Private Function KlassengruppeVorhanden(lngKlasse As Long, lngJgst As Long, lngAusb As Long, lngSchulart As Long) As Long
' prüft, ob es die Klassengruppe bereits gibt, melde die Uid zurück oder 0
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.uid " & _
                "FROM TabKlassengruppe " & _
                "WHERE TabKlassengruppe.uid_klasse = " & lngKlasse & " AND " & _
                "TabKlassengruppe.uid_schulart = " & lngSchulart & " AND " & _
                "TabKlassengruppe.uid_jahrgangsstufe = " & lngJgst & " AND " & _
                "TabKlassengruppe.uid_ausb = " & lngAusb & " AND " & _
                "TabKlassengruppe.geloescht = FALSE"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        KlassengruppeVorhanden = myrs.Fields("uid")
    Else
        KlassengruppeVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function
Private Function LehrerVorhanden(strKuerzel As String) As Long
' prüft ob es den Lehrer bereits gibt, melde die Uid des Lehrers oder 0 zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabLehrer.uid FROM TabLehrer " & _
                "WHERE tabLehrer.kuerzel = " & Chr(34) & strKuerzel & Chr(34) & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        LehrerVorhanden = myrs.Fields("uid")
    Else
        LehrerVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function
Private Function SchuleVorhanden(strSchulnummer As String) As Long
' prüft ob es die Schule bereits gibt, melde die Uid der Schule oder 0 zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchule.uid FROM TabSchule " & _
                "WHERE TabSchule.schulnummer = " & Chr(34) & strSchulnummer & Chr(34) & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        SchuleVorhanden = myrs.Fields("uid")
    Else
        SchuleVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function
Private Function FachVorhanden(strFach As String) As Long
' prüft ob es das Fach mit dem übergebene Kürzel gibt, melde Uid des Faches oder 0 zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabFach.uid FROM TabFach " & _
                "WHERE TabFach.kuerzel = " & Chr(34) & strFach & Chr(34) & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        FachVorhanden = myrs.Fields("uid")
    Else
        FachVorhanden = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function SchuelerNotenHalbjahrVorhanden(lngSchuelerUid As Long, lngFachUid As Long, lngKlassengruppeUid As Long) As Long
' prüft ob es den Halbjahresnotensatz eines Schülers gibt, melde die Uid oder 0 zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenHalbjahrSchuelerFach.uid FROM TabNotenHalbjahrSchuelerFach " & _
                "WHERE TabNotenHalbjahrSchuelerFach.uid_schueler = " & lngSchuelerUid & _
                " AND TabNotenHalbjahrSchuelerFach.uid_fach = " & lngFachUid & _
                " AND TabNotenHalbjahrSchuelerFach.uid_klassengruppe = " & lngKlassengruppeUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        SchuelerNotenHalbjahrVorhanden = myrs.Fields("uid")
    Else
        SchuelerNotenHalbjahrVorhanden = 0
    End If
    
    Set myrs = Nothing

End Function

Private Function SchuelerInKlassengruppe(lngSchueler As Long, lngKlassengruppe As Long) As Long
' prüft, ob der Schüler in der Klassengruppe ist 0 = Nein, 1 = Ja
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.uid " & _
                "FROM TabSchueler " & _
                "WHERE TabSchueler.uid = " & lngSchueler & " AND " & _
                "TabSchueler.uid_Klassengruppe = " & lngKlassengruppe
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        SchuelerInKlassengruppe = 1
    Else
        SchuelerInKlassengruppe = 0
    End If
    
    Set myrs = Nothing

End Function

Private Function SchuelerInKlasse(lngSchueler As Long, strKlasse As String) As Long
' prüft, ob der Schüler in der Klasse ist und liefert die Uid der Klassengruppe zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.uid " & _
                "FROM (TabKlasse INNER JOIN TabKlassengruppe ON TabKlasse.uid = TabKlassengruppe.uid_klasse) INNER JOIN TabSchueler ON TabKlassengruppe.uid = TabSchueler.uid_klassengruppe " & _
                "WHERE TabSchueler.uid = " & lngSchueler & " AND " & _
                "TabKlasse.bezeichnung = " & Chr(34) & strKlasse & Chr(34)
    myrs.Open strSql, myconn, adOpenKeyset, adLockReadOnly
    
    If myrs.RecordCount > 0 Then
        SchuelerInKlasse = myrs.Fields("uid")
    Else
        SchuelerInKlasse = 0
    End If
    
    Set myrs = Nothing

End Function

Private Function SchuelerAnlegen(lngSchuelernummer As Long, strNachname As String, strRufname As String, strVornamen As String, _
                                    strGebdat As String, strGebort As String, strGebland As String, _
                                    strBekenntnis As String, strReligion As String, _
                                    lngKlassengruppe As Long, lngGeschlecht As Long, _
                                    strEintrittsdatum As String, strProbezeit As String) As Long
' schreibt die Schülerdaten in die Tabelle TabSchueler und liefert die Uid des Schülers zurück.
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.* FROM TabSchueler"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
    
        .AddNew
        .Fields("schuelernummer") = lngSchuelernummer
        .Fields("nachname") = strNachname
        .Fields("rufname") = strRufname
        .Fields("vornamen") = strVornamen
        .Fields("gebdat") = strGebdat
        .Fields("gebort") = strGebort
        .Fields("gebland") = strGebland
        .Fields("bekenntnis") = strBekenntnis
        .Fields("religion_unterricht") = strReligion
        .Fields("uid_klassengruppe") = lngKlassengruppe
        .Fields("uid_geschlecht") = lngGeschlecht
        .Fields("eintritt_schule") = strEintrittsdatum
        If Not IsNull(strProbezeit) And Not strProbezeit = "" Then
            .Fields("probezeit_bis") = strProbezeit
        End If
        .Update
        SchuelerAnlegen = .Fields("uid")

    End With
    Set myrs = Nothing
    
End Function
Private Function KlasseAnlegen(strBezeichnung As String) As Long
'schreibt die Klassendaten in die Tabelle TabKlasse und liefert die Uid der Klasse zurück.
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlasse.* FROM TabKlasse"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
    
        .AddNew
        
        .Fields("bezeichnung") = strBezeichnung
        
        .Update
        KlasseAnlegen = .Fields("uid")

    End With
    Set myrs = Nothing
End Function
Private Function KlassengruppeAnlegen(strKlasse As String, lngKlasse As Long, lngJgst As Long, lngAusb As Long, lngSchulart As Long, strAusbKurz As String) As Long
'schreibt die Klassengruppe in die Tabelle TabKlassengruppe und liefert die Uid der Klassengruppe zurück.
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.* FROM TabKlassengruppe"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
    
        .AddNew
        
        .Fields("bezeichnung") = strKlasse & "_" & strAusbKurz
        .Fields("uid_klasse") = lngKlasse
        .Fields("uid_jahrgangsstufe") = lngJgst
        .Fields("uid_ausb") = lngAusb
        .Fields("uid_schulart") = lngSchulart
        
        .Update
        KlassengruppeAnlegen = .Fields("uid")

    End With
    Set myrs = Nothing

End Function
Private Function KlassenNotenAnlegen(lngKlassengruppe As Long, lngJgst As Long, lngAusb As Long, lngSchulart As Long) As Long
'schreibe den Datensatz für die Tabelle TabNotenFachKlassengruppe, der als Grundlage für den Datensatz TabNotenSchuelerFachKlassengruppe dient
' dazu die Standardeinstellungen für jedes Fach aus der Tabelle TabNotenFachJahrgangAusbSchulart holen
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim myrs2 As ADODB.Recordset
Dim strSql2 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenFachJahrgangAusbSchulart " & _
                "WHERE uid_jahrgang = " & lngJgst & " AND " & _
                "uid_ausb = " & lngAusb & " AND " & _
                "uid_schulart = " & lngSchulart
    Set myrs2 = New ADODB.Recordset
    strSql2 = "SELECT * FROM TabNotenFachKlassengruppe"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockReadOnly
    myrs2.Open strSql2, myconn, adOpenKeyset, adLockOptimistic
    
    If myrs1.RecordCount > 0 Then
        myrs1.MoveFirst
        For intCount = 1 To myrs1.RecordCount
            myrs2.AddNew
            myrs2.Fields("uid_klassengruppe") = lngKlassengruppe
            myrs2.Fields("uid_fach") = myrs1.Fields("uid_fach")
            myrs2.Fields("anz_sa_hj1") = myrs1.Fields("anz_sa_hj1")
            myrs2.Fields("anz_sa_hj2") = myrs1.Fields("anz_sa_hj2")
            myrs2.Fields("anz_son_hj1") = myrs1.Fields("anz_son_hj1")
            myrs2.Fields("anz_son_hj2") = myrs1.Fields("anz_son_hj2")
            myrs2.Fields("gew_son1_hj1") = 1
            myrs2.Fields("gew_son2_hj1") = 1
            myrs2.Fields("gew_son1_hj2") = 1
            myrs2.Fields("gew_son2_hj2") = 1
            If myrs1.Fields("anz_son_hj1") = 3 Then
                myrs2.Fields("gew_son3_hj1") = 1
            End If
            If myrs1.Fields("anz_son_hj2") = 3 Then
                myrs2.Fields("gew_son3_hj2") = 1
            End If
            If myrs1.Fields("uid_fach") = 35 Then   'fachpraktische Ausbildung
                myrs2.Fields("gew_son1_hj1") = 2
                myrs2.Fields("gew_son1_hj2") = 2
                myrs2.Fields("uid_art_son1_hj1") = 5
                myrs2.Fields("uid_art_son1_hj2") = 5
                myrs2.Fields("uid_art_son2_hj1") = 6
                myrs2.Fields("uid_art_son2_hj2") = 6
                myrs2.Fields("uid_art_son3_hj1") = 7
                myrs2.Fields("uid_art_son3_hj2") = 7
            End If
            myrs2.Update
            myrs1.MoveNext
        Next intCount
        KlassenNotenAnlegen = 1
    Else
        'nix gefunden
        KlassenNotenAnlegen = 0
    End If

    Set myrs1 = Nothing
    Set myrs2 = Nothing
End Function
Private Function SchuelerNotenAnlegen(lngKlassengruppe As Long, lngSchueler As Long) As Long
'schreibe den Datensatz für die Tabelle TabNotenSchuelerFachKlassengruppe
' dazu die Standardeinstellungen für jedes Fach aus der Tabelle TabNotenFachKlassengruppe holen
' keine Datensätze für Religionsfächer (uid= 8,9 und 10) und Sprachen (uid= 19, 20, 41 und 42), diese werden seperat für die Schüler geschrieben
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim myrs2 As ADODB.Recordset
Dim strSql2 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenFachKlassengruppe " & _
                "WHERE uid_Klassengruppe = " & lngKlassengruppe & " AND " & _
                "(NOT uid_fach=8 AND NOT uid_fach=9 AND NOT uid_fach=10" & _
                " AND NOT uid_fach=19 AND NOT uid_fach=20 AND NOT uid_fach=41 AND NOT uid_fach=42)"
    Set myrs2 = New ADODB.Recordset
    strSql2 = "SELECT * FROM TabNotenSchuelerFachKlassengruppe"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockReadOnly
    myrs2.Open strSql2, myconn, adOpenKeyset, adLockOptimistic
    
    If myrs1.RecordCount > 0 Then
        myrs1.MoveFirst
        For intCount = 1 To myrs1.RecordCount
            myrs2.AddNew
            myrs2.Fields("uid_klassengruppe") = lngKlassengruppe
            myrs2.Fields("uid_schueler") = lngSchueler
            myrs2.Fields("uid_fach") = myrs1.Fields("uid_fach")
            myrs2.Fields("anz_sa_hj1") = myrs1.Fields("anz_sa_hj1")
            myrs2.Fields("anz_sa_hj2") = myrs1.Fields("anz_sa_hj2")
            myrs2.Fields("anz_son_hj1") = myrs1.Fields("anz_son_hj1")
            myrs2.Fields("anz_son_hj2") = myrs1.Fields("anz_son_hj2")
            myrs2.Fields("gew_son1_hj1") = myrs1.Fields("gew_son1_hj1")
            myrs2.Fields("gew_son2_hj1") = myrs1.Fields("gew_son2_hj1")
            myrs2.Fields("gew_son1_hj2") = myrs1.Fields("gew_son1_hj2")
            myrs2.Fields("gew_son2_hj2") = myrs1.Fields("gew_son2_hj2")
            If myrs1.Fields("anz_son_hj1") = 3 Then
                myrs2.Fields("gew_son3_hj1") = myrs1.Fields("gew_son3_hj1")
            End If
            If myrs1.Fields("anz_son_hj2") = 3 Then
                myrs2.Fields("gew_son3_hj2") = myrs1.Fields("gew_son3_hj2")
            End If
            If myrs1.Fields("uid_fach") = 35 Then   'fachpraktische Ausbildung
                myrs2.Fields("uid_art_son1_hj1") = 5
                myrs2.Fields("uid_art_son1_hj2") = 5
                myrs2.Fields("uid_art_son2_hj1") = 6
                myrs2.Fields("uid_art_son2_hj2") = 6
                myrs2.Fields("uid_art_son3_hj1") = 7
                myrs2.Fields("uid_art_son3_hj2") = 7
            End If
            myrs2.Update
            myrs1.MoveNext
        Next intCount
        
        SchuelerNotenAnlegen = 1
    Else
        'nix gefunden
        SchuelerNotenAnlegen = 0
    End If

    Set myrs1 = Nothing
    Set myrs2 = Nothing
    
End Function
Private Function SchuelerNotenReligionAnlegen(lngKlassengruppe As Long, lngSchueler As Long, strReligion As String) As Long
'schreibe den Datensatz für die Tabelle TabNotenSchuelerFachKlassengruppe für den Religionsunterricht
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenSchuelerFachKlassengruppe"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
    
    myrs1.AddNew
    myrs1.Fields("uid_klassengruppe") = lngKlassengruppe
    myrs1.Fields("uid_schueler") = lngSchueler
    Select Case UCase(strReligion)
        Case "RK"
            myrs1.Fields("uid_fach") = 8
        Case "EV"
            myrs1.Fields("uid_fach") = 9
        Case "ETH"
            myrs1.Fields("uid_fach") = 10
        Case Else
            myrs1.Fields("uid_fach") = 10
    End Select
    myrs1.Fields("anz_sa_hj1") = 0
    myrs1.Fields("anz_sa_hj2") = 0
    myrs1.Fields("anz_son_hj1") = 3
    myrs1.Fields("anz_son_hj2") = 3
    myrs1.Fields("gew_son1_hj1") = 1
    myrs1.Fields("gew_son2_hj1") = 1
    myrs1.Fields("gew_son3_hj1") = 1
    myrs1.Fields("gew_son1_hj2") = 1
    myrs1.Fields("gew_son2_hj2") = 1
    myrs1.Fields("gew_son3_hj2") = 1
    myrs1.Update
        
    SchuelerNotenReligionAnlegen = 1

    Set myrs1 = Nothing

End Function

Private Function SchuelerNotenSpracheAnlegen(lngKlassengruppe As Long, lngSchueler As Long, lngJahrgangsstufe As Long, lngSprache As Long) As Long
'schreibe den Datensatz für die Tabelle TabNotenSchuelerFachKlassengruppe für den Sprachunterricht in der Ausbildungsrichtung Int. Wirtschaft
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenSchuelerFachKlassengruppe"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
    
    myrs1.AddNew
    myrs1.Fields("uid_klassengruppe") = lngKlassengruppe
    myrs1.Fields("uid_schueler") = lngSchueler
    myrs1.Fields("uid_fach") = lngSprache
    If lngJahrgangsstufe > 2 Then       'ab Jahrgangsstufe 12 Schulaufgaben
        myrs1.Fields("anz_sa_hj1") = 1
        myrs1.Fields("anz_sa_hj2") = 1
        myrs1.Fields("anz_son_hj1") = 2
        myrs1.Fields("anz_son_hj2") = 2
        myrs1.Fields("gew_son1_hj1") = 1
        myrs1.Fields("gew_son2_hj1") = 1
        myrs1.Fields("gew_son1_hj2") = 1
        myrs1.Fields("gew_son2_hj2") = 1
    Else
        myrs1.Fields("anz_sa_hj1") = 0
        myrs1.Fields("anz_sa_hj2") = 0
        myrs1.Fields("anz_son_hj1") = 3
        myrs1.Fields("anz_son_hj2") = 3
        myrs1.Fields("gew_son1_hj1") = 1
        myrs1.Fields("gew_son2_hj1") = 1
        myrs1.Fields("gew_son3_hj1") = 1
        myrs1.Fields("gew_son1_hj2") = 1
        myrs1.Fields("gew_son2_hj2") = 1
        myrs1.Fields("gew_son3_hj2") = 1
    End If
    myrs1.Update
        
    SchuelerNotenSpracheAnlegen = 1

    Set myrs1 = Nothing

End Function


Private Function SchuelerNotenHalbjahrAnlegen(lngKlassengruppe As Long, lngSchueler As Long, lngJahrgangsstufe As Long) As Long
'schreibe den Datensatz für die Tabelle TabNotenHalbjahrSchuelerFach
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim myrs2 As ADODB.Recordset
Dim strSql2 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenFachKlassengruppe " & _
                "WHERE uid_Klassengruppe = " & lngKlassengruppe & " AND " & _
                "(Not uid_fach=8 And Not uid_fach=9 And Not uid_fach=10" & _
                " AND NOT uid_fach=19 AND NOT uid_fach=20 AND NOT uid_fach=41 AND NOT uid_fach=42)"
    Set myrs2 = New ADODB.Recordset
    strSql2 = "SELECT * FROM TabNotenHalbjahrSchuelerFach"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockReadOnly
    myrs2.Open strSql2, myconn, adOpenKeyset, adLockOptimistic
    
    If myrs1.RecordCount > 0 Then
        myrs1.MoveFirst
        For intCount = 1 To myrs1.RecordCount
            myrs2.AddNew
            myrs2.Fields("uid_klassengruppe") = lngKlassengruppe
            myrs2.Fields("uid_schueler") = lngSchueler
            myrs2.Fields("uid_fach") = myrs1.Fields("uid_fach")
            myrs2.Update
            myrs1.MoveNext
        Next intCount
        
        SchuelerNotenHalbjahrAnlegen = 1
    Else
        'nix gefunden
        SchuelerNotenHalbjahrAnlegen = 0
    End If

    Set myrs1 = Nothing
    Set myrs2 = Nothing
End Function
Private Function SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppe As Long, lngSchueler As Long, lngJahrgangsstufe As Long, strReligion As String) As Long
'schreibe den Datensatz für die Tabelle TabNotenHalbjahrSchuelerFach für den Religionsunterricht
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenHalbjahrSchuelerFach"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
    
    myrs1.AddNew
    myrs1.Fields("uid_klassengruppe") = lngKlassengruppe
    myrs1.Fields("uid_schueler") = lngSchueler
    Select Case UCase(strReligion)
        Case "RK"
            myrs1.Fields("uid_fach") = 8
        Case "EV"
            myrs1.Fields("uid_fach") = 9
        Case "ETH"
            myrs1.Fields("uid_fach") = 10
        Case Else
            myrs1.Fields("uid_fach") = 10
    End Select
    myrs1.Update
        
    SchuelerNotenHalbjahrReligionAnlegen = 1

    Set myrs1 = Nothing

End Function

Private Function SchuelerNotenHalbjahrSpracheAnlegen(lngKlassengruppe As Long, lngSchueler As Long, lngSprache As Long) As Long
'schreibe den Datensatz für die Tabelle TabNotenHalbjahrSchuelerFach für den Sprachunterricht Int. Wirtschaft
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenHalbjahrSchuelerFach"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
    
    myrs1.AddNew
    myrs1.Fields("uid_klassengruppe") = lngKlassengruppe
    myrs1.Fields("uid_schueler") = lngSchueler
    myrs1.Fields("uid_fach") = lngSprache
    myrs1.Update
        
    SchuelerNotenHalbjahrSpracheAnlegen = 1

    Set myrs1 = Nothing

End Function
Private Function SchuelerNotenHalbjahrAendern(lngUid As Long, lngJahrgangsstufeUid As Long, lngHalbjahr1 As Long, lngHalbjahr2 As Long, lngSchuelernummer As Long) As Long
' schreibe die importierten Halbjahresnoten in die Datenbank
Dim myrs As ADODB.Recordset
Dim strSql As String
Dim strMsg As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT * FROM TabNotenHalbjahrSchuelerFach WHERE TabNotenHalbjahrSchuelerFach.uid=" & lngUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        Select Case lngJahrgangsstufeUid
            Case 1, 5, 6, 7   'Vorklassen und Vorkurse
                If myrs.Fields("n_vkl_hj1") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_vkl_hj1") <> lngHalbjahr1 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_vkl_hj1") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr1 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_vkl_hj1") = lngHalbjahr1
                            myrs.Fields("n_vkl_hj1_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr1 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_vkl_hj1") = lngHalbjahr1
                        myrs.Fields("n_vkl_hj1_import") = True
                    End If
                End If
                If myrs.Fields("n_vkl_hj2") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_vkl_hj2") <> lngHalbjahr2 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_vkl_hj2") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr2 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_vkl_hj2") = lngHalbjahr2
                            myrs.Fields("n_vkl_hj2_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr2 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_vkl_hj2") = lngHalbjahr2
                        myrs.Fields("n_vkl_hj2_import") = True
                    End If
                End If
            Case 2      '11.Klasse
                If myrs.Fields("n_11_hj1") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_11_hj1") <> lngHalbjahr1 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_11_hj1") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr1 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_11_hj1") = lngHalbjahr1
                            myrs.Fields("n_11_hj1_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr1 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_11_hj1") = lngHalbjahr1
                        myrs.Fields("n_11_hj1_import") = True
                    End If
                End If
                If myrs.Fields("n_11_hj2") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_11_hj2") <> lngHalbjahr2 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_11_hj2") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr2 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_11_hj2") = lngHalbjahr2
                            myrs.Fields("n_11_hj2_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr2 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_11_hj2") = lngHalbjahr2
                        myrs.Fields("n_11_hj2_import") = True
                    End If
                End If
            Case 3      '12.Klasse
                If myrs.Fields("n_12_hj1") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_12_hj1") <> lngHalbjahr1 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_12_hj1") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr1 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_12_hj1") = lngHalbjahr1
                            myrs.Fields("n_12_hj1_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr1 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_12_hj1") = lngHalbjahr1
                        myrs.Fields("n_12_hj1_import") = True
                    End If
                End If
                If myrs.Fields("n_12_hj2") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_12_hj2") <> lngHalbjahr2 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_12_hj2") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr2 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_12_hj2") = lngHalbjahr2
                            myrs.Fields("n_12_hj2_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr2 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_12_hj2") = lngHalbjahr2
                        myrs.Fields("n_12_hj2_import") = True
                    End If
                End If
            Case 4      '13.Klasse
                If myrs.Fields("n_13_hj1") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_13_hj1") <> lngHalbjahr1 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_13_hj1") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr1 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_13_hj1") = lngHalbjahr1
                            myrs.Fields("n_13_hj1_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr1 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_13_hj1") = lngHalbjahr1
                        myrs.Fields("n_13_hj1_import") = True
                    End If
                End If
                If myrs.Fields("n_13_hj2") > -1 Then   'Note bereits vorhanden
                    If myrs.Fields("n_13_hj2") <> lngHalbjahr2 Then  'gleiche Note nix machen
                        strMsg = "Wollen Sie den gespeicherten Wert " & myrs.Fields("n_13_hj2") & " Punkte für den Schüler " & lngSchuelernummer & " mit dem Wert " & lngHalbjahr2 & " Punkte überschreiben?"
                        If MsgBox(strMsg, vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then 'nachfragen ob das geändert werden soll
                            myrs.Fields("n_13_hj2") = lngHalbjahr2
                            myrs.Fields("n_13_hj2_import") = True
                        End If
                    End If
                Else                                    'Note nicht vorhanden
                    If lngHalbjahr2 > -1 Then           'importierte Halbjahresnote belegt?
                        myrs.Fields("n_13_hj2") = lngHalbjahr2
                        myrs.Fields("n_13_hj2_import") = True
                    End If
                End If
        End Select
        .Update
        
        SchuelerNotenHalbjahrAendern = .Fields("uid")

    End With
    Set myrs = Nothing

End Function
Private Function LehrerAnlegen(strKuerzel As String, strNachname As String, strRufname As String, _
                                    strAmt As String, strTitel As String, lngGeschlecht As Long) As Long
' schreibt die Lehrerdaten in die Tabelle TabLehrer und liefert die Uid des Lehrers zurück.
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabLehrer.* FROM TabLehrer"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
    
        .AddNew
        
        .Fields("kuerzel") = strKuerzel
        .Fields("familienname") = strNachname
        .Fields("rufname") = strRufname
        .Fields("amt") = strAmt
        .Fields("titel") = strTitel
        .Fields("uid_geschlecht") = lngGeschlecht
        
        .Update
        LehrerAnlegen = .Fields("uid")

    End With
    Set myrs = Nothing
    
End Function
Private Function LehrerAendern(lngUid As Long, strKuerzel As String, strNachname As String, strRufname As String, _
                                    strAmt As String, strTitel As String, lngGeschlecht As Long) As Long
' ändert die Lehrerdaten in der Tabelle TabLehrer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabLehrer.* FROM TabLehrer WHERE TabLehrer.uid=" & lngUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        
        .Fields("kuerzel") = strKuerzel
        .Fields("familienname") = strNachname
        .Fields("rufname") = strRufname
        .Fields("amt") = strAmt
        .Fields("titel") = strTitel
        .Fields("uid_geschlecht") = lngGeschlecht
        
        .Update
        
        LehrerAendern = .Fields("uid")

    End With
    Set myrs = Nothing
    
End Function
Private Function SchuleAnlegen(strSchulnummer As String, strSchulart As String, strSchulname As String, _
                                    strSchulnamekurz As String) As Long
' schreibt die Schuldaten in die Tabelle TabSchule und liefert die Uid der Schule zurück.
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchule.* FROM TabSchule"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
    
        .AddNew
        
        .Fields("schulnummer") = strSchulnummer
        .Fields("bezeichnung_lang") = strSchulname
        .Fields("bezeichnung_kurz") = strSchulnamekurz
        Select Case strSchulart
            Case "FOS"
                .Fields("uid_schulart") = 1
            Case "BOS"
                .Fields("uid_schulart") = 2
        End Select
        
        .Update
        SchuleAnlegen = .Fields("uid")

    End With
    Set myrs = Nothing
    
End Function
Private Function SchuleAendern(lngUid As Long, strSchulnummer As String, strSchulart As String, strSchulname As String, _
                                    strSchulnamekurz As String) As Long
' ändert die Schuldaten in der Tabelle TabSchule
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchule.* FROM TabSchule WHERE TabSchule.uid=" & lngUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        
        .Fields("schulnummer") = strSchulnummer
        .Fields("bezeichnung_lang") = strSchulname
        .Fields("bezeichnung_kurz") = strSchulnamekurz
        Select Case strSchulart
            Case "FOS"
                .Fields("uid_schulart") = 1
            Case "BOS"
                .Fields("uid_schulart") = 2
        End Select
        
        .Update
        
        SchuleAendern = .Fields("uid")

    End With
    Set myrs = Nothing
    
End Function

Public Function JahrgangsstufeErmitteln(lngKlassengruppeUid As Long) As Long
' Funktion gibt die Jahrgangsstufe einer Klassengruppe zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT Tabklassengruppe.uid_jahrgangsstufe FROM TabKlassengruppe WHERE TabKlassengruppe.uid=" & lngKlassengruppeUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        JahrgangsstufeErmitteln = .Fields("uid_jahrgangsstufe")
    End With
    Set myrs = Nothing
End Function

Public Function AusbildungsrichtungErmitteln(lngKlassengruppeUid As Long) As Long
' Funktion gibt die Ausbildungsrichtung einer Klassengruppe zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT Tabklassengruppe.uid_ausb FROM TabKlassengruppe WHERE TabKlassengruppe.uid=" & lngKlassengruppeUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        AusbildungsrichtungErmitteln = .Fields("uid_ausb")
    End With
    Set myrs = Nothing
End Function

Public Sub SchuelerNotensaetzeNachKlassengruppenWechselAnlegen(lngKlassengruppeUid As Long, _
                                                                lngSchuelerUid As Long, _
                                                                lngJahrgangsstufeUid As Long, _
                                                                strReligion As String, _
                                                                strRufname As String, _
                                                                strNachname As String)
' die Notensätze eines Schülers neu Anlegen bei Klassengruppenwechsel
    If SchuelerNotenAnlegen(lngKlassengruppeUid, lngSchuelerUid) = 0 Then
        Err.Raise 515, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge für den Schüler " & strRufname & " " & strNachname
    End If
    If SchuelerNotenHalbjahrAnlegen(lngKlassengruppeUid, lngSchuelerUid, lngJahrgangsstufeUid) = 0 Then
        Err.Raise 516, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge für den Schüler " & strRufname & " " & strNachname
    End If
    If lngJahrgangsstufeUid <> 2 Then 'Religion
        If SchuelerNotenReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, strReligion) = 0 Then
            Err.Raise 517, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Noteneinträge Religionsunterricht für den Schüler " & strRufname & " " & strNachname
        End If
        If SchuelerNotenHalbjahrReligionAnlegen(lngKlassengruppeUid, lngSchuelerUid, lngJahrgangsstufeUid, strReligion) = 0 Then
            Err.Raise 518, "Notenverwaltung FOSBOS", "Fehler beim Anlegen der Halbjahreseinträge Religionsunterricht für den Schüler " & strRufname & " " & strNachname
        End If
    End If
End Sub
Public Function SchuelerAnzahleinerKlassengruppeErmitteln(lngKlassengruppeUid) As Long
' Funktion ermittelt die Schüleranzahl einer Klassengruppe
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT Count(*) FROM TabSchueler WHERE TabSchueler.uid_klassengruppe=" & lngKlassengruppeUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        SchuelerAnzahleinerKlassengruppeErmitteln = .Fields(0)
    End With
    Set myrs = Nothing
End Function

Public Sub KlassengruppeOhneSchuelerLoeschen(lngKlassengruppeUid As Long)

Dim lngKlasseUid As Long

' die Funktion setzt die Klassengruppe, die zugehörigen Sätze in der Tabelle TabNotenFachKlassengruppe auf geloescht und ggf. die Klasse auf gelöscht
    If KlassengruppeLoeschen(lngKlassengruppeUid) = 0 Then
        Err.Raise 519, "Notenverwaltung FOSBOS", "Fehler beim Löschen der Klassengruppe"
    End If
    If NotenFachKlassengruppeLoeschen(lngKlassengruppeUid) = 0 Then
        Err.Raise 520, "Notenverwaltung FOSBOS", "Fehler beim Löschen der Klassengruppe in der Tabelle NotenFachKlassengruppe"
    End If
    lngKlasseUid = KlasseEinerKlassengruppeErmitteln(lngKlassengruppeUid)
    If KlassengruppenAnzahleinerKlasseErmitteln(lngKlasseUid) = 0 Then 'keine weiteren Klassengruppen für diese Klasse also löschen
        If KlasseLoeschen(lngKlasseUid) = 0 Then
            Err.Raise 521, "Notenverwaltung FOSBOS", "Fehler beim Löschen der Klasse"
        End If
    End If
End Sub

Private Function KlasseEinerKlassengruppeErmitteln(lngKlassengruppeUid) As Long
' die Funktion ermittelt die Uid der Klasse einer Klassengruppe und liefert sie zurück
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.uid_klasse FROM TabKlassengruppe WHERE TabKlassengruppe.uid=" & lngKlassengruppeUid & ";"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        KlasseEinerKlassengruppeErmitteln = .Fields("uid_klasse")
    End With
    Set myrs = Nothing
End Function

Private Function KlassengruppenAnzahleinerKlasseErmitteln(lngKlasseUid) As Long
' Funktion ermittelt die Anzahl aktiver Klassengruppen einer Klasse
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT Count(*) FROM TabKlassengruppe WHERE TabKlassengruppe.uid_klasse=" & lngKlasseUid & " AND TabKlassengruppe.geloescht=FALSE;"
    With myrs
        .Open strSql, myconn, adOpenKeyset, adLockOptimistic
        KlassengruppenAnzahleinerKlasseErmitteln = .Fields(0)
    End With
    Set myrs = Nothing
End Function

Private Function KlassengruppeLoeschen(lngKlassengruppeUid As Long) As Long
' diese Funktion löscht die Klassengruppe
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.geloescht FROM TabKlassengruppe WHERE TabKlassengruppe.uid=" & lngKlassengruppeUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.Fields("geloescht") = True
        myrs.Update
        KlassengruppeLoeschen = 1
    Else
        KlassengruppeLoeschen = 0
    End If
    
    Set myrs = Nothing
    
End Function

Private Function NotenFachKlassengruppeLoeschen(lngKlassengruppeUid As Long) As Long
' diese Funktion löscht die Sätze der Klassengruppe in der Tabelle TabNotenFachKlassengruppe
Dim intCount As Integer
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlassengruppe.geloescht FROM TabKlassengruppe WHERE TabKlassengruppe.uid=" & lngKlassengruppeUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            myrs.Fields("geloescht") = True
            myrs.Update
            myrs.MoveNext
        Next intCount
        NotenFachKlassengruppeLoeschen = 1
    Else
        NotenFachKlassengruppeLoeschen = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function KlasseLoeschen(lngKlasseUid As Long) As Long
' diese Funktion löscht die Klasse
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabKlasse.geloescht FROM TabKlasse WHERE TabKlasse.uid=" & lngKlasseUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.Fields("geloescht") = True
        myrs.Update
        KlasseLoeschen = 1
    Else
        KlasseLoeschen = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function KlassengruppeEinesSchuelersErmitteln(lngSchuelerUid As Long) As Long
' diese Funktion ermittelt die Klassengruppe eines Schülers
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.uid_klassengruppe FROM TabSchueler WHERE TabSchueler.uid=" & lngSchuelerUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        KlassengruppeEinesSchuelersErmitteln = myrs.Fields("uid_klassengruppe")
    Else
        KlassengruppeEinesSchuelersErmitteln = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function KlassengruppeEinesSchuelersAendern(lngSchuelerUid As Long, lngKlassengruppeUid As Long) As Long
' diese Funktion ändert die Klassengruppe eines Schülers
Dim myrs As ADODB.Recordset
Dim strSql As String

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabSchueler.uid_klassengruppe FROM TabSchueler WHERE TabSchueler.uid=" & lngSchuelerUid & ";"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.Fields("uid_klassengruppe") = lngKlassengruppeUid
        myrs.Update
        KlassengruppeEinesSchuelersAendern = 1
    Else
        KlassengruppeEinesSchuelersAendern = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function NotensaetzeEinesSchuelersLoeschen(lngSchuelerUid As Long, lngKlassengruppeUid As Long) As Long
' diese Funktion löscht die Notensaetze eines Schülers
Dim myrs As ADODB.Recordset
Dim strSql As String
Dim intCount As Integer

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenSchuelerFachKlassengruppe.uid_schueler, TabNotenSchuelerFachKlassengruppe.uid_klassengruppe, TabNotenSchuelerFachKlassengruppe.geloescht " & _
                "FROM TabNotenSchuelerFachKlassengruppe WHERE " & _
                "TabNotenSchuelerFachKlassengruppe.uid_schueler=" & lngSchuelerUid & _
                " AND TabNotenSchuelerFachKlassengruppe.uid_klassengruppe=" & lngKlassengruppeUid & _
                " AND TabNotenSchuelerFachKlassengruppe.geloescht=FALSE;"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            myrs.Fields("geloescht") = True
            myrs.Update
            myrs.MoveNext
        Next intCount
        NotensaetzeEinesSchuelersLoeschen = 1
    Else
        NotensaetzeEinesSchuelersLoeschen = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function NotensaetzeHalbjahrEinesSchuelersLoeschen(lngSchuelerUid As Long, lngKlassengruppeUid As Long) As Long
' diese Funktion löscht die Notensaetze eines Schülers
Dim myrs As ADODB.Recordset
Dim strSql As String
Dim intCount As Integer

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenHalbjahrSchuelerFach.uid_schueler, TabNotenHalbjahrSchuelerFach.uid_klassengruppe, TabNotenHalbjahrSchuelerFach.geloescht " & _
                "FROM TabNotenHalbjahrSchuelerFach WHERE " & _
                "TabNotenHalbjahrSchuelerFach.uid_schueler=" & lngSchuelerUid & _
                " AND TabNotenHalbjahrSchuelerFach.uid_klassengruppe=" & lngKlassengruppeUid & _
                " AND TabNotenHalbjahrSchuelerFach.geloescht=FALSE;"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            myrs.Fields("geloescht") = True
            myrs.Update
            myrs.MoveNext
        Next intCount
        NotensaetzeHalbjahrEinesSchuelersLoeschen = 1
    Else
        NotensaetzeHalbjahrEinesSchuelersLoeschen = 0
    End If
    
    Set myrs = Nothing
End Function

Private Function NotensaetzeHalbjahrEinesSchuelersAendern(lngSchuelerUid As Long, lngKlassengruppeUid As Long, lngKlassengruppeUidAlt As Long, lngJahrgangsstufeUid) As Long
' diese Funktion ändert die Notensaetze eines Schülers: die Klassengruppe wird geändert und die Noten entsprechend der Jahrgangsstufe zurückgesetzt.
Dim myrs As ADODB.Recordset
Dim strSql As String
Dim intCount As Integer

    Set myrs = New ADODB.Recordset
    strSql = "SELECT TabNotenHalbjahrSchuelerFach.uid_schueler, TabNotenHalbjahrSchuelerFach.uid_klassengruppe, TabNotenHalbjahrSchuelerFach.geloescht " & _
                "FROM TabNotenHalbjahrSchuelerFach WHERE " & _
                "TabNotenHalbjahrSchuelerFach.uid_schueler=" & lngSchuelerUid & _
                " AND TabNotenHalbjahrSchuelerFach.uid_klassengruppe=" & lngKlassengruppeUidAlt & _
                " AND TabNotenHalbjahrSchuelerFach.geloescht=FALSE;"
    myrs.Open strSql, myconn, adOpenKeyset, adLockOptimistic
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            myrs.Fields("uid_klassengruppe") = lngKlassengruppeUid
            Select Case lngJahrgangsstufeUid
                Case 1      'Vorklasse
                    myrs.Fields("n_vkl_hj1") = -1
                    myrs.Fields("n_vkl_hj2") = -1
                    myrs.Fields("n_vkl_hj1_import") = False
                    myrs.Fields("n_vkl_hj2_import") = False
                Case 2      '11.Klasse
                    myrs.Fields("n_11_hj1") = -1
                    myrs.Fields("n_11_hj2") = -1
                    myrs.Fields("n_11_hj1_import") = False
                    myrs.Fields("n_11_hj2_import") = False
                Case 3      '12.Klasse
                    myrs.Fields("n_12_hj1") = -1
                    myrs.Fields("n_12_hj2") = -1
                    myrs.Fields("n_12_hj1_import") = False
                    myrs.Fields("n_12_hj2_import") = False
                    myrs.Fields("n_12_hj1_eingebracht") = False
                    myrs.Fields("n_12_hj2_eingebracht") = False
                    myrs.Fields("n_11_hj1_eingebracht") = False
                    myrs.Fields("n_11_hj2_eingebracht") = False
                Case 4      '13.Klasse
                    myrs.Fields("n_13_hj1") = -1
                    myrs.Fields("n_13_hj2") = -1
                    myrs.Fields("n_13_hj1_import") = False
                    myrs.Fields("n_13_hj2_import") = False
            End Select
            myrs.Update
            myrs.MoveNext
        Next intCount
        NotensaetzeHalbjahrEinesSchuelersAendern = 1
    Else
        NotensaetzeHalbjahrEinesSchuelersAendern = 0
    End If
    
    Set myrs = Nothing
End Function
Private Function NotensaetzeHalbjahrEinesSchuelersErgaenzen(lngSchuelerUid As Long, lngKlassengruppeUid As Long, lngKlassengruppeUidAlt As Long) As Long
' diese Funktion ändert vorhandene Notensätze eines Schülers auf die neue Klassengruppe und ergänzt fehlende Notensätze für die 12. Jahrgangsstufe
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim myrs2 As ADODB.Recordset
Dim strSql2 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenFachKlassengruppe " & _
                "WHERE uid_Klassengruppe = " & lngKlassengruppeUid & " AND " & _
                "(Not uid_fach=8 And Not uid_fach=9 And Not uid_fach=10)"
    Set myrs2 = New ADODB.Recordset
    strSql2 = "SELECT * FROM TabNotenHalbjahrSchuelerFach WHERE TabNotenHalbjahrSchuelerFach.uid_schueler=" & lngSchuelerUid & " AND TabNotenHalbjahrSchuelerFach.uid_klassengruppe=" & lngKlassengruppeUidAlt
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockReadOnly
    myrs2.Open strSql2, myconn, adOpenKeyset, adLockOptimistic
    ' zunächst bei allen vorhandenen Sätzen die Klassengruppe ändern
    If myrs2.RecordCount > 0 Then
        myrs2.MoveFirst
        For intCount = 1 To myrs2.RecordCount
            myrs2.Fields("uid_klassengruppe") = lngKlassengruppeUid
            myrs2.Update
            myrs2.MoveNext
        Next intCount
    End If
    
    If myrs1.RecordCount > 0 Then
        myrs1.MoveFirst
        For intCount = 1 To myrs1.RecordCount
            myrs2.MoveFirst
            myrs2.Find "uid_fach=" & myrs1.Fields("uid_fach")
            'wenn gefunden nix machen
            If Not myrs2.EOF And Not myrs2.BOF Then
            Else
            'wenn nicht gefunden Satz anlegen
                myrs2.AddNew
                myrs2.Fields("uid_klassengruppe") = lngKlassengruppeUid
                myrs2.Fields("uid_schueler") = lngSchuelerUid
                myrs2.Fields("uid_fach") = myrs1.Fields("uid_fach")
                myrs2.Update
            End If
            myrs1.MoveNext
        Next intCount
        
        NotensaetzeHalbjahrEinesSchuelersErgaenzen = 1
    Else
        'nix gefunden
        NotensaetzeHalbjahrEinesSchuelersErgaenzen = 0
    End If

    Set myrs1 = Nothing
    Set myrs2 = Nothing

End Function
Public Function KlassenNotenFachHinzufuegen(lngKlassengruppeUid As Long, lngFachUid As Long) As Boolean
' ein einzelnes Fach zur Tabelle TabNotenFachKlassengruppe hinzufügen
Dim myrs1 As ADODB.Recordset
Dim strSql1 As String
Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT * FROM TabNotenFachKlassengruppe"
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockOptimistic
    With myrs1
            .AddNew
            .Fields("uid_klassengruppe") = lngKlassengruppeUid
            .Fields("uid_fach") = lngFachUid
            .Fields("anz_sa_hj1") = 0
            .Fields("anz_sa_hj2") = 0
            .Fields("anz_son_hj1") = 3
            .Fields("anz_son_hj2") = 3
            .Fields("gew_son1_hj1") = 1
            .Fields("gew_son2_hj1") = 1
            .Fields("gew_son1_hj2") = 1
            .Fields("gew_son2_hj2") = 1
            .Fields("gew_son3_hj1") = 1
            .Fields("gew_son3_hj2") = 1
            .Update
    End With
        
    KlassenNotenFachHinzufuegen = 1

    Set myrs1 = Nothing
    
End Function
Public Function SchuelerNotenEinzelnesFachAnlegen(lngKlassengruppeUid As Long, lngFachUid As Long) As Boolean
' für alle Schüler einer Klassengruppe die Notensätze inklusive Halbjahressätze anlegen
' lies alle Schüler der Klassengruppe
' lies den Notensatz der Klassengruppe für dieses Fach
' füge den Notensatz für jeden Schüler hinzu

Dim myrs1 As ADODB.Recordset    ' Recordset Alle Schüler der Klassengruppe
Dim myrs2 As ADODB.Recordset    ' Recordset Noteneinstellungen der Klassengruppe für das einzelne Fach
Dim myrs3 As ADODB.Recordset    ' Recordset Schülernoten
Dim myrs4 As ADODB.Recordset    ' Recordset Halbjahresleistungen
Dim strSql1 As String
Dim strSql2 As String
Dim strSql3 As String
Dim strSql4 As String

Dim intCount As Integer

    Set myrs1 = New ADODB.Recordset
    strSql1 = "SELECT TabSchueler.uid FROM TabSchueler WHERE TabSchueler.uid_klassengruppe=" & lngKlassengruppeUid
    
    myrs1.Open strSql1, myconn, adOpenKeyset, adLockReadOnly
    
    Set myrs2 = New ADODB.Recordset
    strSql2 = "SELECT * FROM TabNotenFachKlassengruppe WHERE TabNotenFachKlassengruppe.uid_klassengruppe=" & lngKlassengruppeUid & " AND TabNotenFachKlassengruppe.uid_fach=" & lngFachUid
    
    myrs2.Open strSql2, myconn, adOpenKeyset, adLockReadOnly
    
    Set myrs3 = New ADODB.Recordset
    strSql3 = "SELECT * FROM TabNotenSchuelerFachKlassengruppe"
    
    myrs3.Open strSql3, myconn, adOpenKeyset, adLockOptimistic
    
    Set myrs4 = New ADODB.Recordset
    strSql4 = "SELECT * FROM TabNotenHalbjahrSchuelerFach"
    
    myrs4.Open strSql4, myconn, adOpenKeyset, adLockOptimistic

    myrs1.MoveFirst
    For intCount = 1 To myrs1.RecordCount
        With myrs3
            .AddNew
            .Fields("uid_klassengruppe") = lngKlassengruppeUid
            .Fields("uid_schueler") = myrs1.Fields("uid")
            .Fields("uid_fach") = lngFachUid
            .Fields("anz_sa_hj1") = myrs2.Fields("anz_sa_hj1")
            .Fields("anz_sa_hj2") = myrs2.Fields("anz_sa_hj2")
            .Fields("anz_son_hj1") = myrs2.Fields("anz_son_hj1")
            .Fields("anz_son_hj2") = myrs2.Fields("anz_son_hj2")
            .Fields("gew_son1_hj1") = myrs2.Fields("gew_son1_hj1")
            .Fields("gew_son2_hj1") = myrs2.Fields("gew_son2_hj1")
            .Fields("gew_son1_hj2") = myrs2.Fields("gew_son1_hj2")
            .Fields("gew_son2_hj2") = myrs2.Fields("gew_son2_hj2")
            .Fields("gew_son3_hj1") = myrs2.Fields("gew_son3_hj1")
            .Fields("gew_son3_hj2") = myrs2.Fields("gew_son3_hj2")
            .Update
        End With
        With myrs4
            .AddNew
            .Fields("uid_klassengruppe") = lngKlassengruppeUid
            .Fields("uid_schueler") = myrs1.Fields("uid")
            .Fields("uid_fach") = lngFachUid
            .Update
        End With
        myrs1.MoveNext
    Next intCount
    
    Set myrs1 = Nothing
    Set myrs2 = Nothing
    Set myrs3 = Nothing
    Set myrs4 = Nothing
    
End Function
