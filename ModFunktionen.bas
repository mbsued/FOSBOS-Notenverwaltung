Attribute VB_Name = "ModFunktionen"
Option Compare Database
Option Explicit

'Modul f�r alle Funktionen
Public Function SchuelerNotenLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As colSch�lerNoten
' diese Funktion liest die Daten aller Sch�ler entsprechend dem ausgew�hlten Fach und der ausgew�hlten Klassengruppe
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection Sch�lerNoten, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer
Dim varAnzLeistungsnachweise(0 To 3) As Variant
Dim varArtLeistungsnachweise(0 To 13) As Variant
Dim varDatumLeistungsnachweise(0 To 17) As Variant
Dim varGewLeistungsnachweise(0 To 17) As Variant
Dim varNotenLeistungsnachweise(0 To 17) As Variant

Dim mycolSch�lerNoten As colSch�lerNoten

    Set mydb = New clsDB
    Set myrs = mydb.SchuelerNotenLesen(lngFachUid, lngKlassengruppenUid)
    
    Set mycolSch�lerNoten = New colSch�lerNoten
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            varAnzLeistungsnachweise(0) = myrs.Fields("anz_sa_hj1")
            varAnzLeistungsnachweise(1) = myrs.Fields("anz_sa_hj2")
            varAnzLeistungsnachweise(2) = myrs.Fields("anz_son_hj1")
            varAnzLeistungsnachweise(3) = myrs.Fields("anz_son_hj2")
            
            varArtLeistungsnachweise(0) = myrs.Fields("uid_art_son1_hj1")
            varArtLeistungsnachweise(1) = myrs.Fields("uid_art_son2_hj1")
            varArtLeistungsnachweise(2) = myrs.Fields("uid_art_son3_hj1")
            varArtLeistungsnachweise(3) = myrs.Fields("uid_art_son4_hj1")
            varArtLeistungsnachweise(4) = myrs.Fields("uid_art_son5_hj1")
            varArtLeistungsnachweise(5) = myrs.Fields("uid_art_son6_hj1")
            varArtLeistungsnachweise(6) = myrs.Fields("uid_art_son7_hj1")
            varArtLeistungsnachweise(7) = myrs.Fields("uid_art_son1_hj2")
            varArtLeistungsnachweise(8) = myrs.Fields("uid_art_son2_hj2")
            varArtLeistungsnachweise(9) = myrs.Fields("uid_art_son3_hj2")
            varArtLeistungsnachweise(10) = myrs.Fields("uid_art_son4_hj2")
            varArtLeistungsnachweise(11) = myrs.Fields("uid_art_son5_hj2")
            varArtLeistungsnachweise(12) = myrs.Fields("uid_art_son6_hj2")
            varArtLeistungsnachweise(13) = myrs.Fields("uid_art_son7_hj2")
            
            varDatumLeistungsnachweise(0) = myrs.Fields("dat_sa1_hj1")
            varDatumLeistungsnachweise(1) = myrs.Fields("dat_sa2_hj1")
            varDatumLeistungsnachweise(2) = myrs.Fields("dat_sa1_hj2")
            varDatumLeistungsnachweise(3) = myrs.Fields("dat_sa2_hj2")
            varDatumLeistungsnachweise(4) = myrs.Fields("dat_son1_hj1")
            varDatumLeistungsnachweise(5) = myrs.Fields("dat_son2_hj1")
            varDatumLeistungsnachweise(6) = myrs.Fields("dat_son3_hj1")
            varDatumLeistungsnachweise(7) = myrs.Fields("dat_son4_hj1")
            varDatumLeistungsnachweise(8) = myrs.Fields("dat_son5_hj1")
            varDatumLeistungsnachweise(9) = myrs.Fields("dat_son6_hj1")
            varDatumLeistungsnachweise(10) = myrs.Fields("dat_son7_hj1")
            varDatumLeistungsnachweise(11) = myrs.Fields("dat_son1_hj2")
            varDatumLeistungsnachweise(12) = myrs.Fields("dat_son2_hj2")
            varDatumLeistungsnachweise(13) = myrs.Fields("dat_son3_hj2")
            varDatumLeistungsnachweise(14) = myrs.Fields("dat_son4_hj2")
            varDatumLeistungsnachweise(15) = myrs.Fields("dat_son5_hj2")
            varDatumLeistungsnachweise(16) = myrs.Fields("dat_son6_hj2")
            varDatumLeistungsnachweise(17) = myrs.Fields("dat_son7_hj2")
            
            varGewLeistungsnachweise(0) = myrs.Fields("gew_sa1_hj1")
            varGewLeistungsnachweise(1) = myrs.Fields("gew_sa2_hj1")
            varGewLeistungsnachweise(2) = myrs.Fields("gew_sa1_hj2")
            varGewLeistungsnachweise(3) = myrs.Fields("gew_sa2_hj2")
            varGewLeistungsnachweise(4) = myrs.Fields("gew_son1_hj1")
            varGewLeistungsnachweise(5) = myrs.Fields("gew_son2_hj1")
            varGewLeistungsnachweise(6) = myrs.Fields("gew_son3_hj1")
            varGewLeistungsnachweise(7) = myrs.Fields("gew_son4_hj1")
            varGewLeistungsnachweise(8) = myrs.Fields("gew_son5_hj1")
            varGewLeistungsnachweise(9) = myrs.Fields("gew_son6_hj1")
            varGewLeistungsnachweise(10) = myrs.Fields("gew_son7_hj1")
            varGewLeistungsnachweise(11) = myrs.Fields("gew_son1_hj2")
            varGewLeistungsnachweise(12) = myrs.Fields("gew_son2_hj2")
            varGewLeistungsnachweise(13) = myrs.Fields("gew_son3_hj2")
            varGewLeistungsnachweise(14) = myrs.Fields("gew_son4_hj2")
            varGewLeistungsnachweise(15) = myrs.Fields("gew_son5_hj2")
            varGewLeistungsnachweise(16) = myrs.Fields("gew_son6_hj2")
            varGewLeistungsnachweise(17) = myrs.Fields("gew_son7_hj2")
            
            varNotenLeistungsnachweise(0) = myrs.Fields("n_sa1_hj1")
            varNotenLeistungsnachweise(1) = myrs.Fields("n_sa2_hj1")
            varNotenLeistungsnachweise(2) = myrs.Fields("n_sa1_hj2")
            varNotenLeistungsnachweise(3) = myrs.Fields("n_sa2_hj2")
            varNotenLeistungsnachweise(4) = myrs.Fields("n_son1_hj1")
            varNotenLeistungsnachweise(5) = myrs.Fields("n_son2_hj1")
            varNotenLeistungsnachweise(6) = myrs.Fields("n_son3_hj1")
            varNotenLeistungsnachweise(7) = myrs.Fields("n_son4_hj1")
            varNotenLeistungsnachweise(8) = myrs.Fields("n_son5_hj1")
            varNotenLeistungsnachweise(9) = myrs.Fields("n_son6_hj1")
            varNotenLeistungsnachweise(10) = myrs.Fields("n_son7_hj1")
            varNotenLeistungsnachweise(11) = myrs.Fields("n_son1_hj2")
            varNotenLeistungsnachweise(12) = myrs.Fields("n_son2_hj2")
            varNotenLeistungsnachweise(13) = myrs.Fields("n_son3_hj2")
            varNotenLeistungsnachweise(14) = myrs.Fields("n_son4_hj2")
            varNotenLeistungsnachweise(15) = myrs.Fields("n_son5_hj2")
            varNotenLeistungsnachweise(16) = myrs.Fields("n_son6_hj2")
            varNotenLeistungsnachweise(17) = myrs.Fields("n_son7_hj2")
            
            mycolSch�lerNoten.Add myrs.Fields("uid"), myrs.Fields("nachname"), myrs.Fields("rufname"), myrs.Fields("uid_schueler"), _
                                    myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), myrs.Fields("ind_einstellung"), _
                                    varAnzLeistungsnachweise, varArtLeistungsnachweise, varDatumLeistungsnachweise, varGewLeistungsnachweise, varNotenLeistungsnachweise, myrs.Fields("geloescht")
            myrs.MoveNext
        Next intCount
        Set SchuelerNotenLesen = mycolSch�lerNoten
    Else
        Set SchuelerNotenLesen = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolSch�lerNoten = Nothing
    
End Function
Public Function SchuelerNotenLesenKurz(lngSchuelerUid As Long) As colSch�lerNotenKurz
' diese Funktion liest die Daten eines bestimmten Sch�lers anhand seiner Uid
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection Sch�lerNotenKurz, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer

Dim mycolSch�lerNotenKurz As colSch�lerNotenKurz

    Set mydb = New clsDB
    Set myrs = mydb.SchuelerNotenLesenKurz(lngSchuelerUid)
    
    Set mycolSch�lerNotenKurz = New colSch�lerNotenKurz
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            mycolSch�lerNotenKurz.Add myrs.Fields("uid"), myrs.Fields("uid_schueler"), myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), myrs.Fields("geloescht")
            myrs.MoveNext
        Next intCount
        Set SchuelerNotenLesenKurz = mycolSch�lerNotenKurz
    Else
        Set SchuelerNotenLesenKurz = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolSch�lerNotenKurz = Nothing
    
End Function
Public Function SchuelerHalbjahresNotenLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As colSch�lerHalbjahr
' diese Funktion liest die Halbjahresdaten aller Sch�ler entsprechend dem ausgew�hlten Fach und der ausgew�hlten Klassengruppe
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection Sch�lerhalbjahr, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer

Dim mycolSch�lerHalbjahr As colSch�lerHalbjahr

    Set mydb = New clsDB
    Set myrs = mydb.SchuelerHalbjahrLesen(lngFachUid, lngKlassengruppenUid)
    
    Set mycolSch�lerHalbjahr = New colSch�lerHalbjahr
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            mycolSch�lerHalbjahr.Add myrs.Fields("uid"), myrs.Fields("uid_schueler"), myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), myrs.Fields("uid_jahrgangsstufe"), _
                                        myrs.Fields("n_vkl_hj1"), myrs.Fields("n_vkl_hj2"), myrs.Fields("n_11_hj1"), myrs.Fields("n_11_hj2"), myrs.Fields("n_12_hj1"), myrs.Fields("n_12_hj2"), _
                                        myrs.Fields("n_13_hj1"), myrs.Fields("n_13_hj2"), _
                                        myrs.Fields("n_11_hj1_abgewaehlt"), myrs.Fields("n_11_hj2_abgewaehlt"), _
                                        myrs.Fields("n_12_hj1_abgewaehlt"), myrs.Fields("n_12_hj2_abgewaehlt"), _
                                        myrs.Fields("n_13_hj1_abgewaehlt"), myrs.Fields("n_13_hj2_abgewaehlt")
                                        
            myrs.MoveNext
        Next intCount
        Set SchuelerHalbjahresNotenLesen = mycolSch�lerHalbjahr
    Else
        Set SchuelerHalbjahresNotenLesen = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolSch�lerHalbjahr = Nothing
    
End Function
Public Function SchuelerHalbjahresNotenLesenKurz(lngSchuelerUid As Long) As colSch�lerHalbjahrKurz
' diese Funktion liest die Halbjahresdaten eines Sch�lers
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection Sch�lerhalbjahrkurz, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer

Dim mycolSch�lerHalbjahrKurz As colSch�lerHalbjahrKurz

    Set mydb = New clsDB
    Set myrs = mydb.SchuelerHalbjahrLesenKurz(lngSchuelerUid)
    
    Set mycolSch�lerHalbjahrKurz = New colSch�lerHalbjahrKurz
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            mycolSch�lerHalbjahrKurz.Add myrs.Fields("uid"), myrs.Fields("uid_schueler"), myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), myrs.Fields("uid_jahrgangsstufe"), myrs.Fields("geloescht")
                                        
            myrs.MoveNext
        Next intCount
        Set SchuelerHalbjahresNotenLesenKurz = mycolSch�lerHalbjahrKurz
    Else
        Set SchuelerHalbjahresNotenLesenKurz = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolSch�lerHalbjahrKurz = Nothing
    
End Function
Public Function FaecherHalbjahresNotenLesen(lngSchuelerUid As Long, lngKlassengruppenUid As Long) As colFaecherHalbjahr
' diese Funktion liest die Halbjahresdaten aller F�cher entsprechend dem ausgew�hlten Fach und des ausgew�hlten Sch�lers
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection F�cherHalbjahr, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer

Dim mycolFaecherhalbjahr As colFaecherHalbjahr

    Set mydb = New clsDB
    Set myrs = mydb.FaecherHalbjahrLesen(lngSchuelerUid, lngKlassengruppenUid)
    
    Set mycolFaecherhalbjahr = New colFaecherHalbjahr
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            mycolFaecherhalbjahr.Add myrs.Fields("uid"), myrs.Fields("uid_schueler"), myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), _
                                        myrs.Fields("uid_jahrgangsstufe"), myrs.Fields("uid_schulart"), myrs.Fields("bezeichnung_lang"), _
                                        myrs.Fields("n_vkl_hj1"), myrs.Fields("n_vkl_hj2"), myrs.Fields("n_11_hj1"), myrs.Fields("n_11_hj2"), myrs.Fields("n_12_hj1"), myrs.Fields("n_12_hj2"), _
                                        myrs.Fields("n_13_hj1"), myrs.Fields("n_13_hj2"), _
                                        myrs.Fields("n_11_hj1_abgewaehlt"), myrs.Fields("n_11_hj2_abgewaehlt"), _
                                        myrs.Fields("n_12_hj1_abgewaehlt"), myrs.Fields("n_12_hj2_abgewaehlt"), _
                                        myrs.Fields("n_13_hj1_abgewaehlt"), myrs.Fields("n_13_hj2_abgewaehlt")
                                        
            myrs.MoveNext
        Next intCount
        Set FaecherHalbjahresNotenLesen = mycolFaecherhalbjahr
    Else
        Set FaecherHalbjahresNotenLesen = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolFaecherhalbjahr = Nothing
    
End Function

Public Function NotenEinstellungenKlasseFachLesen(lngFachUid As Long, lngKlassengruppenUid As Long) As clsKlasseNoten
' diese Funktion liest die Einstellungsdaten gem�� dem ausgew�hlten Fach und der ausgew�hlten Klassengruppe
' zuerst Zugriff auf die Datenbank, anschlie�end das Objekt clsKlasseNoten f�llen und dies zur�ckliefern
Dim mydb As clsDB
Dim myrs As ADODB.Recordset

Dim myclsKlasseNoten As clsKlasseNoten

    Set mydb = New clsDB
    Set myrs = mydb.KlasseNotenLesen(lngFachUid, lngKlassengruppenUid)
    
    Set myclsKlasseNoten = New clsKlasseNoten
    myrs.MoveFirst
    With myclsKlasseNoten
        .uid = myrs.Fields("uid")
        .anz_sa_hj1 = myrs.Fields("anz_sa_hj1")
        .anz_sa_hj2 = myrs.Fields("anz_sa_hj2")
        .anz_son_hj1 = myrs.Fields("anz_son_hj1")
        .anz_son_hj2 = myrs.Fields("anz_son_hj2")
        .uid_art_son1_hj1 = myrs.Fields("uid_art_son1_hj1")
        .uid_art_son2_hj1 = myrs.Fields("uid_art_son2_hj1")
        .uid_art_son3_hj1 = myrs.Fields("uid_art_son3_hj1")
        .uid_art_son4_hj1 = myrs.Fields("uid_art_son4_hj1")
        .uid_art_son5_hj1 = myrs.Fields("uid_art_son5_hj1")
        .uid_art_son6_hj1 = myrs.Fields("uid_art_son6_hj1")
        .uid_art_son7_hj1 = myrs.Fields("uid_art_son7_hj1")
        .uid_art_son1_hj2 = myrs.Fields("uid_art_son1_hj2")
        .uid_art_son2_hj2 = myrs.Fields("uid_art_son2_hj2")
        .uid_art_son3_hj2 = myrs.Fields("uid_art_son3_hj2")
        .uid_art_son4_hj2 = myrs.Fields("uid_art_son4_hj2")
        .uid_art_son5_hj2 = myrs.Fields("uid_art_son5_hj2")
        .uid_art_son6_hj2 = myrs.Fields("uid_art_son6_hj2")
        .uid_art_son7_hj2 = myrs.Fields("uid_art_son7_hj2")
        .dat_sa1_hj1 = myrs.Fields("dat_sa1_hj1")
        .dat_sa2_hj1 = myrs.Fields("dat_sa2_hj1")
        .dat_sa1_hj2 = myrs.Fields("dat_sa1_hj2")
        .dat_sa2_hj2 = myrs.Fields("dat_sa2_hj2")
        .dat_son1_hj1 = myrs.Fields("dat_son1_hj1")
        .dat_son2_hj1 = myrs.Fields("dat_son2_hj1")
        .dat_son3_hj1 = myrs.Fields("dat_son3_hj1")
        .dat_son4_hj1 = myrs.Fields("dat_son4_hj1")
        .dat_son5_hj1 = myrs.Fields("dat_son5_hj1")
        .dat_son6_hj1 = myrs.Fields("dat_son6_hj1")
        .dat_son7_hj1 = myrs.Fields("dat_son7_hj1")
        .dat_son1_hj2 = myrs.Fields("dat_son1_hj2")
        .dat_son2_hj2 = myrs.Fields("dat_son2_hj2")
        .dat_son3_hj2 = myrs.Fields("dat_son3_hj2")
        .dat_son4_hj2 = myrs.Fields("dat_son4_hj2")
        .dat_son5_hj2 = myrs.Fields("dat_son5_hj2")
        .dat_son6_hj2 = myrs.Fields("dat_son6_hj2")
        .dat_son7_hj2 = myrs.Fields("dat_son7_hj2")
        .gew_son1_hj1 = myrs.Fields("gew_son1_hj1")
        .gew_son2_hj1 = myrs.Fields("gew_son2_hj1")
        .gew_son3_hj1 = myrs.Fields("gew_son3_hj1")
        .gew_son4_hj1 = myrs.Fields("gew_son4_hj1")
        .gew_son5_hj1 = myrs.Fields("gew_son5_hj1")
        .gew_son6_hj1 = myrs.Fields("gew_son6_hj1")
        .gew_son7_hj1 = myrs.Fields("gew_son7_hj1")
        .gew_son1_hj2 = myrs.Fields("gew_son1_hj2")
        .gew_son2_hj2 = myrs.Fields("gew_son2_hj2")
        .gew_son3_hj2 = myrs.Fields("gew_son3_hj2")
        .gew_son4_hj2 = myrs.Fields("gew_son4_hj2")
        .gew_son5_hj2 = myrs.Fields("gew_son5_hj2")
        .gew_son6_hj2 = myrs.Fields("gew_son6_hj2")
        .gew_son7_hj2 = myrs.Fields("gew_son7_hj2")
        .fach_uid = myrs.Fields("uid_fach")
    End With
    
    Set NotenEinstellungenKlasseFachLesen = myclsKlasseNoten
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set myclsKlasseNoten = Nothing
End Function

Public Function KlasseFaecherLesen(lngKlassengruppenUid As Long) As colKlasseF�cher
' diese Funktion liest die F�cher entsprechend der ausgew�hlten Klassengruppe
' zuerst Zugriff auf die Datenbank, anschlie�end sammeln der Daten in der Collection KlasseF�cher, diese wird zur�ckgeliefert
Dim mydb As clsDB
Dim myrs As ADODB.Recordset
Dim intCount As Integer

Dim mycolKlasseF�cher As colKlasseF�cher

    Set mydb = New clsDB
    Set myrs = mydb.KlasseFaecherLesen(lngKlassengruppenUid)
    
    Set mycolKlasseF�cher = New colKlasseF�cher
    
    If myrs.RecordCount > 0 Then
        myrs.MoveFirst
        For intCount = 1 To myrs.RecordCount
            mycolKlasseF�cher.Add myrs.Fields("uid"), myrs.Fields("uid_fach"), myrs.Fields("uid_klassengruppe"), myrs.Fields("bezeichnung_kurz"), myrs.Fields("geloescht")
            myrs.MoveNext
        Next intCount
        Set KlasseFaecherLesen = mycolKlasseF�cher
    Else
        Set KlasseFaecherLesen = Nothing
    End If
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set mycolKlasseF�cher = Nothing
    
End Function

Public Function Sch�lerLesen(lngSchuelerUid As Long) As clsSch�lerdaten
' diese Funktion liest die Sch�lerdaten
' zuerst Zugriff auf die Datenbank, anschlie�end das Objekt clsSch�lerDaten f�llen und dies zur�ckliefern
Dim mydb As clsDB
Dim myrs As ADODB.Recordset

Dim myclsSch�lerdaten As clsSch�lerdaten

    Set mydb = New clsDB
    Set myrs = mydb.Sch�lerdatenLesen(lngSchuelerUid)
    
    Set myclsSch�lerdaten = New clsSch�lerdaten
    
    myrs.MoveFirst
    With myclsSch�lerdaten
        .uid = myrs.Fields("uid")
        .nachname = myrs.Fields("nachname")
        .rufname = myrs.Fields("rufname")
        .vornamen = myrs.Fields("vornamen")
        .geburtsdatum = myrs.Fields("gebdat")
        .geburtsort = myrs.Fields("gebort")
        .geburtsland = myrs.Fields("gebland")
        .bekenntnis = myrs.Fields("bekenntnis")
        .religion = myrs.Fields("religion_unterricht")
        .ausgeschieden = myrs.Fields("ausgeschieden")
        .ausgeschiedenam = Nz(myrs.Fields("ausgeschieden_am"), "")
        .geschlecht = myrs.Fields("uid_geschlecht")
        .klassengruppe = myrs.Fields("uid_klassengruppe")
        .eintritt = myrs.Fields("eintritt_schule")
        If Not IsNull(myrs.Fields("probezeit_bis")) Then
            .probezeit = myrs.Fields("probezeit_bis")
        End If
        .notenschutz = myrs.Fields("notenschutz")
    End With
    
    Set Sch�lerLesen = myclsSch�lerdaten
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set myclsSch�lerdaten = Nothing
    
End Function
Public Function Sch�lerSchreiben(myclsSch�lerdaten As clsSch�lerdaten) As Boolean
'diese Funktion schreibt die Sch�lerdaten
Dim mydb As clsDB

    Set mydb = New clsDB
    mydb.Sch�lerdatenSpeichern myclsSch�lerdaten
    Set mydb = Nothing
    
End Function

Public Function LehrerLesen(lngLehrerUid As Long) As clsLehrerDaten
' diese Funktion liest die Lehrerdaten
' zuerst Zugriff auf die Datenbank, anschlie�end das Objekt clsLehrerDaten f�llen und dies zur�ckliefern
Dim mydb As clsDB
Dim myrs As ADODB.Recordset

Dim myclsLehrerdaten As clsLehrerDaten

    Set mydb = New clsDB
    Set myrs = mydb.LehrerdatenLesen(lngLehrerUid)
    
    Set myclsLehrerdaten = New clsLehrerDaten
    
    myrs.MoveFirst
    With myclsLehrerdaten
        .uid = myrs.Fields("uid")
        .kuerzel = myrs.Fields("kuerzel")
        .nachname = myrs.Fields("familienname")
        .rufname = myrs.Fields("rufname")
        .amt = myrs.Fields("amt")
        .titel = myrs.Fields("titel")
        .geschlecht = myrs.Fields("uid_geschlecht")
        .email = Nz(myrs.Fields("email"), "")
        .schulleitung = myrs.Fields("schulleitung")
    End With
    
    Set LehrerLesen = myclsLehrerdaten
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set myclsLehrerdaten = Nothing
    
End Function
Public Function LehrerSchreiben(myclsLehrerdaten As clsLehrerDaten) As Boolean
'diese Funktion schreibt die Lehrerdaten
Dim mydb As clsDB

    Set mydb = New clsDB
    mydb.LehrerdatenSpeichern myclsLehrerdaten
    Set mydb = Nothing
    
End Function
Public Function KlasseLesen(lngKlasseUid As Long) As clsKlasseDaten
' diese Funktion liest die Klassendaten
' zuerst Zugriff auf die Datenbank, anschlie�end das Objekt clsKlasseDaten f�llen und dies zur�ckliefern
Dim mydb As clsDB
Dim myrs As ADODB.Recordset

Dim myclsKlassedaten As clsKlasseDaten

    Set mydb = New clsDB
    Set myrs = mydb.KlassendatenLesen(lngKlasseUid)
    
    Set myclsKlassedaten = New clsKlasseDaten
    
    myrs.MoveFirst
    With myclsKlassedaten
        .uid = myrs.Fields("uid")
        .bezeichnung = myrs.Fields("bezeichnung")
        .klassleitung = Nz(myrs.Fields("klassleitung"), 0)
        .zeugnisunterzeichner = Nz(myrs.Fields("zeugnisunterzeichnung"), 0)
    End With
    
    Set KlasseLesen = myclsKlassedaten
    
    Set myrs = Nothing
    Set mydb = Nothing
    Set myclsKlassedaten = Nothing
    
End Function
Public Function KlasseSchreiben(myclsKlassedaten As clsKlasseDaten) As Boolean
'diese Funktion schreibt die Klassendaten
Dim mydb As clsDB

    Set mydb = New clsDB
    mydb.KlassendatenSpeichern myclsKlassedaten
    Set mydb = Nothing
    
End Function

Public Function SchuelerNotenSchreiben(mcolSchuelerNoten As colSch�lerNoten) As Boolean
'Funktion zum Schreiben der Sch�lernoten
Dim mydb As clsDB

    Set mydb = New clsDB
    Set mydb.SchuelerNoten = mcolSchuelerNoten
    mydb.SchuelerNotenSpeichern
    Set mydb = Nothing

End Function

Public Function SchuelerNotenSchreibenKurz(mcolSchuelerNotenKurz As colSch�lerNotenKurz) As Boolean
'Funktion zum Schreiben der Sch�lernoten (kurz)
Dim mydb As clsDB

    Set mydb = New clsDB
    Set mydb.SchuelerNotenKurz = mcolSchuelerNotenKurz
    mydb.SchuelerNotenSpeichernKurz
    Set mydb = Nothing

End Function

Public Function SchuelerHalbjahresNotenSchreiben(mcolSchuelerHalbjahr As colSch�lerHalbjahr) As Boolean
'Funktion zum Schreiben der Halbjahresnoten
Dim mydb As clsDB

    Set mydb = New clsDB
    Set mydb.SchuelerHalbjahr = mcolSchuelerHalbjahr
    mydb.SchuelerHalbjahrSpeichern
    Set mydb = Nothing

End Function
Public Function SchuelerHalbjahresNotenSchreibenKurz(mcolSchuelerHalbjahrKurz As colSch�lerHalbjahrKurz) As Boolean
'Funktion zum Schreiben der Halbjahresnoten kurz
Dim mydb As clsDB

    Set mydb = New clsDB
    Set mydb.SchuelerHalbjahrKurz = mcolSchuelerHalbjahrKurz
    mydb.SchuelerHalbjahrSpeichernKurz
    Set mydb = Nothing

End Function
Public Function NotenEinstellungenSchreiben(mclsKlasseNoten As clsKlasseNoten) As Boolean
'Funktion zum Schreiben der Noteneinstellungen der Klasse
Dim mydb As clsDB

    Set mydb = New clsDB
    mydb.KlasseNotenSpeichern mclsKlasseNoten
    Set mydb = Nothing

End Function
Public Function KlasseFaecherSchreiben(mcolKlasseFaecher As colKlasseF�cher) As Boolean
' Funktion zum Schreiben der F�cher einer Klasse
Dim mydb As clsDB

    Set mydb = New clsDB
    Set mydb.KlasseFaecher = mcolKlasseFaecher
    mydb.KlasseFaecherSchreiben
    Set mydb = Nothing
    
End Function
'eigene Funktion zum kaufm�nnischen Runden
Public Function Runden(ByVal Number As Double, ByVal Digits As Integer) As Double
  Runden = Int(Number * 10 ^ Digits + 0.5) / 10 ^ Digits
End Function
'Funktion zum �bertragen der Halbjahresleistung
'Art 1 = Wert in Formularfeld �bertragen
'Art 2 = Formularwert in Klassenmodulwert �bertragen
Public Function HalbjahresleistungUebertragen(lngArt As Long, StrWert As String) As Variant
    If lngArt = 1 Then
        If StrWert = -1 Then
            HalbjahresleistungUebertragen = "-"
        Else
            HalbjahresleistungUebertragen = StrWert
        End If
    Else
        If StrWert = "-" Then
            HalbjahresleistungUebertragen = -1
        Else
            HalbjahresleistungUebertragen = StrWert
        End If
    End If
End Function
'Prozedur zum Ausgeben der Fehlermeldungen
Public Sub FehlermeldungAusgeben(strModul As String, strProzedur As String, strFehlerNr As String, strFehlerBeschreibung As String)
Dim strFehlertext As String

    strFehlertext = "Fehler im Modul " & strModul & " in der Prozedur " & strProzedur & vbCrLf & _
                        "Fehlernummer: " & strFehlerNr & vbCrLf & _
                        "Beschreibung: " & strFehlerBeschreibung
    MsgBox strFehlertext, vbOKOnly, "FOSBOS Notenverwaltung"
    
End Sub
'Prozedur zum Ausgeben einer Meldung
Public Sub MeldungAusgeben(strMeldung As String)

    MsgBox strMeldung, vbOKOnly, "FOSBOS Notenverwaltung"
    
End Sub


