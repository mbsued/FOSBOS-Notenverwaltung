VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchülerNoten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Schüler und seine Noten
Private lngUid As Long
Private strNachname As String
Private strRufname As String
Private lngSchuelerUid As Long
Private lngFachUid As Long
Private lngKlassengruppeUid As Long
Private booIndEinstellung As Boolean
'Variablen für Anzahl der Leistungsnachweise
Private lngAnzahlSaHj1 As Long
Private lngAnzahlSaHj2 As Long
Private lngAnzahlSonHj1 As Long
Private lngAnzahlSonHj2 As Long
'Variablen für Art der sonstigen Leistungsnachweise
Private lngUidArtSon1Hj1 As Long
Private lngUidArtSon2Hj1 As Long
Private lngUidArtSon3Hj1 As Long
Private lngUidArtSon4Hj1 As Long
Private lngUidArtSon5Hj1 As Long
Private lngUidArtSon6Hj1 As Long
Private lngUidArtSon7Hj1 As Long
Private lngUidArtSon1Hj2 As Long
Private lngUidArtSon2Hj2 As Long
Private lngUidArtSon3Hj2 As Long
Private lngUidArtSon4Hj2 As Long
Private lngUidArtSon5Hj2 As Long
Private lngUidArtSon6Hj2 As Long
Private lngUidArtSon7Hj2 As Long
'Variablen für Datum der Leistungsnachweise
Private datSa1Hj1 As Variant
Private datSa2Hj1 As Variant
Private datSa1Hj2 As Variant
Private datSa2Hj2 As Variant
Private datSon1Hj1 As Variant
Private datSon2Hj1 As Variant
Private datSon3Hj1 As Variant
Private datSon4Hj1 As Variant
Private datSon5Hj1 As Variant
Private datSon6Hj1 As Variant
Private datSon7Hj1 As Variant
Private datSon1Hj2 As Variant
Private datSon2Hj2 As Variant
Private datSon3Hj2 As Variant
Private datSon4Hj2 As Variant
Private datSon5Hj2 As Variant
Private datSon6Hj2 As Variant
Private datSon7Hj2 As Variant
'Variablen für Gewichtung der Schülernoten
Private lngGewSa1Hj1 As Variant
Private lngGewSa2Hj1 As Variant
Private lngGewSa1Hj2 As Variant
Private lngGewSa2Hj2 As Variant
Private lngGewSon1Hj1 As Variant
Private lngGewSon2Hj1 As Variant
Private lngGewSon3Hj1 As Variant
Private lngGewSon4Hj1 As Variant
Private lngGewSon5Hj1 As Variant
Private lngGewSon6Hj1 As Variant
Private lngGewSon7Hj1 As Variant
Private lngGewSon1Hj2 As Variant
Private lngGewSon2Hj2 As Variant
Private lngGewSon3Hj2 As Variant
Private lngGewSon4Hj2 As Variant
Private lngGewSon5Hj2 As Variant
Private lngGewSon6Hj2 As Variant
Private lngGewSon7Hj2 As Variant
'Variablen für Schülernoten
Private varNSa1Hj1 As Variant
Private varNSa2Hj1 As Variant
Private varNSa1Hj2 As Variant
Private varNSa2Hj2 As Variant
Private varNSon1Hj1 As Variant
Private varNSon2Hj1 As Variant
Private varNSon3Hj1 As Variant
Private varNSon4Hj1 As Variant
Private varNSon5Hj1 As Variant
Private varNSon6Hj1 As Variant
Private varNSon7Hj1 As Variant
Private varNSon1Hj2 As Variant
Private varNSon2Hj2 As Variant
Private varNSon3Hj2 As Variant
Private varNSon4Hj2 As Variant
Private varNSon5Hj2 As Variant
Private varNSon6Hj2 As Variant
Private varNSon7Hj2 As Variant
'Varibale für gelöscht
Private booGeloescht As Boolean

Public Property Get uid() As Long
    uid = lngUid
End Property

Public Property Let uid(ByVal vNewValue As Long)
    lngUid = vNewValue
End Property

Public Property Get nachname() As String
    nachname = strNachname
End Property

Public Property Let nachname(ByVal vNewValue As String)
    strNachname = vNewValue
End Property

Public Property Get rufname() As String
    rufname = strRufname
End Property

Public Property Let rufname(ByVal vNewValue As String)
    strRufname = vNewValue
End Property

Public Property Get schueler_uid() As Long
    schueler_uid = lngSchuelerUid
End Property

Public Property Let schueler_uid(ByVal vNewValue As Long)
    lngSchuelerUid = vNewValue
End Property

Public Property Get fach_uid() As Long
    fach_uid = lngFachUid
End Property

Public Property Let fach_uid(ByVal vNewValue As Long)
    lngFachUid = vNewValue
End Property

Public Property Get klassengruppe_uid() As Long
    klassengruppe_uid = lngKlassengruppeUid
End Property

Public Property Let klassengruppe_uid(ByVal vNewValue As Long)
    lngKlassengruppeUid = vNewValue
End Property

Public Property Get ind_einstellung() As Boolean
    ind_einstellung = booIndEinstellung
End Property

Public Property Let ind_einstellung(ByVal vNewValue As Boolean)
    booIndEinstellung = vNewValue
End Property

Public Property Get anz_sa_hj1() As Long
    anz_sa_hj1 = lngAnzahlSaHj1
End Property

Public Property Let anz_sa_hj1(ByVal vNewValue As Long)
    lngAnzahlSaHj1 = vNewValue
End Property

Public Property Get anz_sa_hj2() As Long
    anz_sa_hj2 = lngAnzahlSaHj2
End Property

Public Property Let anz_sa_hj2(ByVal vNewValue As Long)
    lngAnzahlSaHj2 = vNewValue
End Property

Public Property Get anz_son_hj1() As Long
    anz_son_hj1 = lngAnzahlSonHj1
End Property

Public Property Let anz_son_hj1(ByVal vNewValue As Long)
    lngAnzahlSonHj1 = vNewValue
End Property

Public Property Get anz_son_hj2() As Long
    anz_son_hj2 = lngAnzahlSonHj2
End Property

Public Property Let anz_son_hj2(ByVal vNewValue As Long)
    lngAnzahlSonHj2 = vNewValue
End Property

Public Property Get uid_art_son1_hj1() As Long
    uid_art_son1_hj1 = lngUidArtSon1Hj1
End Property

Public Property Let uid_art_son1_hj1(ByVal vNewValue As Long)
    lngUidArtSon1Hj1 = vNewValue
End Property

Public Property Get uid_art_son2_hj1() As Long
    uid_art_son2_hj1 = lngUidArtSon2Hj1
End Property

Public Property Let uid_art_son2_hj1(ByVal vNewValue As Long)
    lngUidArtSon2Hj1 = vNewValue
End Property

Public Property Get uid_art_son3_hj1() As Long
    uid_art_son3_hj1 = lngUidArtSon3Hj1
End Property

Public Property Let uid_art_son3_hj1(ByVal vNewValue As Long)
    lngUidArtSon3Hj1 = vNewValue
End Property

Public Property Get uid_art_son4_hj1() As Long
    uid_art_son4_hj1 = lngUidArtSon4Hj1
End Property

Public Property Let uid_art_son4_hj1(ByVal vNewValue As Long)
    lngUidArtSon4Hj1 = vNewValue
End Property

Public Property Get uid_art_son5_hj1() As Long
    uid_art_son5_hj1 = lngUidArtSon5Hj1
End Property

Public Property Let uid_art_son5_hj1(ByVal vNewValue As Long)
    lngUidArtSon5Hj1 = vNewValue
End Property

Public Property Get uid_art_son6_hj1() As Long
    uid_art_son6_hj1 = lngUidArtSon6Hj1
End Property

Public Property Let uid_art_son6_hj1(ByVal vNewValue As Long)
    lngUidArtSon6Hj1 = vNewValue
End Property

Public Property Get uid_art_son7_hj1() As Long
    uid_art_son7_hj1 = lngUidArtSon7Hj1
End Property

Public Property Let uid_art_son7_hj1(ByVal vNewValue As Long)
    lngUidArtSon7Hj1 = vNewValue
End Property

Public Property Get uid_art_son1_hj2() As Long
    uid_art_son1_hj2 = lngUidArtSon1Hj2
End Property

Public Property Let uid_art_son1_hj2(ByVal vNewValue As Long)
    lngUidArtSon1Hj2 = vNewValue
End Property

Public Property Get uid_art_son2_hj2() As Long
    uid_art_son2_hj2 = lngUidArtSon2Hj2
End Property

Public Property Let uid_art_son2_hj2(ByVal vNewValue As Long)
    lngUidArtSon2Hj2 = vNewValue
End Property

Public Property Get uid_art_son3_hj2() As Long
    uid_art_son3_hj2 = lngUidArtSon3Hj2
End Property

Public Property Let uid_art_son3_hj2(ByVal vNewValue As Long)
    lngUidArtSon3Hj2 = vNewValue
End Property

Public Property Get uid_art_son4_hj2() As Long
    uid_art_son4_hj2 = lngUidArtSon4Hj2
End Property

Public Property Let uid_art_son4_hj2(ByVal vNewValue As Long)
    lngUidArtSon4Hj2 = vNewValue
End Property

Public Property Get uid_art_son5_hj2() As Long
    uid_art_son5_hj2 = lngUidArtSon5Hj2
End Property

Public Property Let uid_art_son5_hj2(ByVal vNewValue As Long)
    lngUidArtSon5Hj2 = vNewValue
End Property

Public Property Get uid_art_son6_hj2() As Long
    uid_art_son6_hj2 = lngUidArtSon6Hj2
End Property

Public Property Let uid_art_son6_hj2(ByVal vNewValue As Long)
    lngUidArtSon6Hj2 = vNewValue
End Property

Public Property Get uid_art_son7_hj2() As Long
    uid_art_son7_hj2 = lngUidArtSon7Hj2
End Property

Public Property Let uid_art_son7_hj2(ByVal vNewValue As Long)
    lngUidArtSon7Hj2 = vNewValue
End Property

Public Property Get dat_sa1_hj1() As Variant
    dat_sa1_hj1 = datSa1Hj1
End Property

Public Property Let dat_sa1_hj1(ByVal vNewValue As Variant)
    datSa1Hj1 = vNewValue
End Property

Public Property Get dat_sa2_hj1() As Variant
    dat_sa2_hj1 = datSa2Hj1
End Property

Public Property Let dat_sa2_hj1(ByVal vNewValue As Variant)
    datSa2Hj1 = vNewValue
End Property

Public Property Get dat_sa1_hj2() As Variant
    dat_sa1_hj2 = datSa1Hj2
End Property

Public Property Let dat_sa1_hj2(ByVal vNewValue As Variant)
    datSa1Hj2 = vNewValue
End Property

Public Property Get dat_sa2_hj2() As Variant
    dat_sa2_hj2 = datSa2Hj2
End Property

Public Property Let dat_sa2_hj2(ByVal vNewValue As Variant)
    datSa2Hj2 = vNewValue
End Property

Public Property Get dat_son1_hj1() As Variant
    dat_son1_hj1 = datSon1Hj1
End Property

Public Property Let dat_son1_hj1(ByVal vNewValue As Variant)
    datSon1Hj1 = vNewValue
End Property

Public Property Get dat_son2_hj1() As Variant
    dat_son2_hj1 = datSon2Hj1
End Property

Public Property Let dat_son2_hj1(ByVal vNewValue As Variant)
    datSon2Hj1 = vNewValue
End Property

Public Property Get dat_son3_hj1() As Variant
    dat_son3_hj1 = datSon3Hj1
End Property

Public Property Let dat_son3_hj1(ByVal vNewValue As Variant)
    datSon3Hj1 = vNewValue
End Property

Public Property Get dat_son4_hj1() As Variant
    dat_son4_hj1 = datSon4Hj1
End Property

Public Property Let dat_son4_hj1(ByVal vNewValue As Variant)
    datSon4Hj1 = vNewValue
End Property

Public Property Get dat_son5_hj1() As Variant
    dat_son5_hj1 = datSon5Hj1
End Property

Public Property Let dat_son5_hj1(ByVal vNewValue As Variant)
    datSon5Hj1 = vNewValue
End Property

Public Property Get dat_son6_hj1() As Variant
    dat_son6_hj1 = datSon6Hj1
End Property

Public Property Let dat_son6_hj1(ByVal vNewValue As Variant)
    datSon6Hj1 = vNewValue
End Property

Public Property Get dat_son7_hj1() As Variant
    dat_son7_hj1 = datSon7Hj1
End Property

Public Property Let dat_son7_hj1(ByVal vNewValue As Variant)
    datSon7Hj1 = vNewValue
End Property

Public Property Get dat_son1_hj2() As Variant
    dat_son1_hj2 = datSon1Hj2
End Property

Public Property Let dat_son1_hj2(ByVal vNewValue As Variant)
    datSon1Hj2 = vNewValue
End Property

Public Property Get dat_son2_hj2() As Variant
    dat_son2_hj2 = datSon2Hj2
End Property

Public Property Let dat_son2_hj2(ByVal vNewValue As Variant)
    datSon2Hj2 = vNewValue
End Property

Public Property Get dat_son3_hj2() As Variant
    dat_son3_hj2 = datSon3Hj2
End Property

Public Property Let dat_son3_hj2(ByVal vNewValue As Variant)
    datSon3Hj2 = vNewValue
End Property

Public Property Get dat_son4_hj2() As Variant
    dat_son4_hj2 = datSon4Hj2
End Property

Public Property Let dat_son4_hj2(ByVal vNewValue As Variant)
    datSon4Hj2 = vNewValue
End Property

Public Property Get dat_son5_hj2() As Variant
    dat_son5_hj2 = datSon5Hj2
End Property

Public Property Let dat_son5_hj2(ByVal vNewValue As Variant)
    datSon5Hj2 = vNewValue
End Property

Public Property Get dat_son6_hj2() As Variant
    dat_son6_hj2 = datSon6Hj2
End Property

Public Property Let dat_son6_hj2(ByVal vNewValue As Variant)
    datSon6Hj2 = vNewValue
End Property

Public Property Get dat_son7_hj2() As Variant
    dat_son7_hj2 = datSon7Hj2
End Property

Public Property Let dat_son7_hj2(ByVal vNewValue As Variant)
    datSon7Hj2 = vNewValue
End Property

Public Property Get gew_sa1_hj1() As Variant
    gew_sa1_hj1 = lngGewSa1Hj1
End Property

Public Property Let gew_sa1_hj1(ByVal vNewValue As Variant)
    lngGewSa1Hj1 = vNewValue
End Property

Public Property Get gew_sa2_hj1() As Variant
    gew_sa2_hj1 = lngGewSa2Hj1
End Property

Public Property Let gew_sa2_hj1(ByVal vNewValue As Variant)
    lngGewSa2Hj1 = vNewValue
End Property

Public Property Get gew_sa1_hj2() As Variant
    gew_sa1_hj2 = lngGewSa1Hj2
End Property

Public Property Let gew_sa1_hj2(ByVal vNewValue As Variant)
    lngGewSa1Hj2 = vNewValue
End Property

Public Property Get gew_sa2_hj2() As Variant
    gew_sa2_hj2 = lngGewSa2Hj2
End Property

Public Property Let gew_sa2_hj2(ByVal vNewValue As Variant)
    lngGewSa2Hj2 = vNewValue
End Property

Public Property Get gew_son1_hj1() As Variant
    gew_son1_hj1 = lngGewSon1Hj1
End Property

Public Property Let gew_son1_hj1(ByVal vNewValue As Variant)
    lngGewSon1Hj1 = vNewValue
End Property

Public Property Get gew_son2_hj1() As Variant
    gew_son2_hj1 = lngGewSon2Hj1
End Property

Public Property Let gew_son2_hj1(ByVal vNewValue As Variant)
    lngGewSon2Hj1 = vNewValue
End Property

Public Property Get gew_son3_hj1() As Variant
    gew_son3_hj1 = lngGewSon3Hj1
End Property

Public Property Let gew_son3_hj1(ByVal vNewValue As Variant)
    lngGewSon3Hj1 = vNewValue
End Property

Public Property Get gew_son4_hj1() As Variant
    gew_son4_hj1 = lngGewSon4Hj1
End Property

Public Property Let gew_son4_hj1(ByVal vNewValue As Variant)
    lngGewSon4Hj1 = vNewValue
End Property

Public Property Get gew_son5_hj1() As Variant
    gew_son5_hj1 = lngGewSon5Hj1
End Property

Public Property Let gew_son5_hj1(ByVal vNewValue As Variant)
    lngGewSon5Hj1 = vNewValue
End Property

Public Property Get gew_son6_hj1() As Variant
    gew_son6_hj1 = lngGewSon6Hj1
End Property

Public Property Let gew_son6_hj1(ByVal vNewValue As Variant)
    lngGewSon6Hj1 = vNewValue
End Property

Public Property Get gew_son7_hj1() As Variant
    gew_son7_hj1 = lngGewSon7Hj1
End Property

Public Property Let gew_son7_hj1(ByVal vNewValue As Variant)
    lngGewSon7Hj1 = vNewValue
End Property

Public Property Get gew_son1_hj2() As Variant
    gew_son1_hj2 = lngGewSon1Hj2
End Property

Public Property Let gew_son1_hj2(ByVal vNewValue As Variant)
    lngGewSon1Hj2 = vNewValue
End Property

Public Property Get gew_son2_hj2() As Variant
    gew_son2_hj2 = lngGewSon2Hj2
End Property

Public Property Let gew_son2_hj2(ByVal vNewValue As Variant)
    lngGewSon2Hj2 = vNewValue
End Property

Public Property Get gew_son3_hj2() As Variant
    gew_son3_hj2 = lngGewSon3Hj2
End Property

Public Property Let gew_son3_hj2(ByVal vNewValue As Variant)
    lngGewSon3Hj2 = vNewValue
End Property

Public Property Get gew_son4_hj2() As Variant
    gew_son4_hj2 = lngGewSon4Hj2
End Property

Public Property Let gew_son4_hj2(ByVal vNewValue As Variant)
    lngGewSon4Hj2 = vNewValue
End Property

Public Property Get gew_son5_hj2() As Variant
    gew_son5_hj2 = lngGewSon5Hj2
End Property

Public Property Let gew_son5_hj2(ByVal vNewValue As Variant)
    lngGewSon5Hj2 = vNewValue
End Property

Public Property Get gew_son6_hj2() As Variant
    gew_son6_hj2 = lngGewSon6Hj2
End Property

Public Property Let gew_son6_hj2(ByVal vNewValue As Variant)
    lngGewSon6Hj2 = vNewValue
End Property

Public Property Get gew_son7_hj2() As Variant
    gew_son7_hj2 = lngGewSon7Hj2
End Property

Public Property Let gew_son7_hj2(ByVal vNewValue As Variant)
    lngGewSon7Hj2 = vNewValue
End Property

Public Property Get n_sa1_hj1() As Variant
    n_sa1_hj1 = varNSa1Hj1
End Property

Public Property Let n_sa1_hj1(ByVal vNewValue As Variant)
    varNSa1Hj1 = vNewValue
End Property

Public Property Get n_sa2_hj1() As Variant
    n_sa2_hj1 = varNSa2Hj1
End Property

Public Property Let n_sa2_hj1(ByVal vNewValue As Variant)
    varNSa2Hj1 = vNewValue
End Property

Public Property Get n_sa1_hj2() As Variant
    n_sa1_hj2 = varNSa1Hj2
End Property

Public Property Let n_sa1_hj2(ByVal vNewValue As Variant)
    varNSa1Hj2 = vNewValue
End Property

Public Property Get n_sa2_hj2() As Variant
    n_sa2_hj2 = varNSa2Hj2
End Property

Public Property Let n_sa2_hj2(ByVal vNewValue As Variant)
    varNSa2Hj2 = vNewValue
End Property

Public Property Get n_son1_hj1() As Variant
    n_son1_hj1 = varNSon1Hj1
End Property

Public Property Let n_son1_hj1(ByVal vNewValue As Variant)
    varNSon1Hj1 = vNewValue
End Property

Public Property Get n_son2_hj1() As Variant
    n_son2_hj1 = varNSon2Hj1
End Property

Public Property Let n_son2_hj1(ByVal vNewValue As Variant)
    varNSon2Hj1 = vNewValue
End Property

Public Property Get n_son3_hj1() As Variant
    n_son3_hj1 = varNSon3Hj1
End Property

Public Property Let n_son3_hj1(ByVal vNewValue As Variant)
    varNSon3Hj1 = vNewValue
End Property

Public Property Get n_son4_hj1() As Variant
    n_son4_hj1 = varNSon4Hj1
End Property

Public Property Let n_son4_hj1(ByVal vNewValue As Variant)
    varNSon4Hj1 = vNewValue
End Property

Public Property Get n_son5_hj1() As Variant
    n_son5_hj1 = varNSon5Hj1
End Property

Public Property Let n_son5_hj1(ByVal vNewValue As Variant)
    varNSon5Hj1 = vNewValue
End Property

Public Property Get n_son6_hj1() As Variant
    n_son6_hj1 = varNSon6Hj1
End Property

Public Property Let n_son6_hj1(ByVal vNewValue As Variant)
    varNSon6Hj1 = vNewValue
End Property

Public Property Get n_son7_hj1() As Variant
    n_son7_hj1 = varNSon7Hj1
End Property

Public Property Let n_son7_hj1(ByVal vNewValue As Variant)
    varNSon7Hj1 = vNewValue
End Property

Public Property Get n_son1_hj2() As Variant
    n_son1_hj2 = varNSon1Hj2
End Property

Public Property Let n_son1_hj2(ByVal vNewValue As Variant)
    varNSon1Hj2 = vNewValue
End Property

Public Property Get n_son2_hj2() As Variant
    n_son2_hj2 = varNSon2Hj2
End Property

Public Property Let n_son2_hj2(ByVal vNewValue As Variant)
    varNSon2Hj2 = vNewValue
End Property

Public Property Get n_son3_hj2() As Variant
    n_son3_hj2 = varNSon3Hj2
End Property

Public Property Let n_son3_hj2(ByVal vNewValue As Variant)
    varNSon3Hj2 = vNewValue
End Property

Public Property Get n_son4_hj2() As Variant
    n_son4_hj2 = varNSon4Hj2
End Property

Public Property Let n_son4_hj2(ByVal vNewValue As Variant)
    varNSon4Hj2 = vNewValue
End Property

Public Property Get n_son5_hj2() As Variant
    n_son5_hj2 = varNSon5Hj2
End Property

Public Property Let n_son5_hj2(ByVal vNewValue As Variant)
    varNSon5Hj2 = vNewValue
End Property

Public Property Get n_son6_hj2() As Variant
    n_son6_hj2 = varNSon6Hj2
End Property

Public Property Let n_son6_hj2(ByVal vNewValue As Variant)
    varNSon6Hj2 = vNewValue
End Property

Public Property Get n_son7_hj2() As Variant
    n_son7_hj2 = varNSon7Hj2
End Property

Public Property Let n_son7_hj2(ByVal vNewValue As Variant)
    varNSon7Hj2 = vNewValue
End Property

Public Property Get geloescht() As Boolean
    geloescht = booGeloescht
End Property

Public Property Let geloescht(ByVal vNewValue As Boolean)
    booGeloescht = vNewValue
End Property

