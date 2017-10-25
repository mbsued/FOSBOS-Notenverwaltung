VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchüler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Schüler
Private lngSchuelernummer As Long
Private lngSchulart As Long
Private strKlasse As String
Private strNachname As String
Private strRufname As String
Private strVornamen As String
Private strGebdat As String
Private strGebort As String
Private strGebland As String
Private strBekenntnis As String
Private strReligion As String
Private lngGeschlecht As Long
Private lngJgst As Long
Private lngAusb As Long
Private strAusbKurz As String
Private strEintrittsdatum As String
Private strProbezeitbis As String
Private lngSprache As Long

Public Property Get schuelernummer() As Long
    schuelernummer = lngSchuelernummer
End Property

Public Property Let schuelernummer(ByVal vNewValue As Long)
    lngSchuelernummer = vNewValue
End Property

Public Property Get schulart() As Long
    schulart = lngSchulart
End Property

Public Property Let schulart(ByVal vNewValue As Long)
    lngSchulart = vNewValue
End Property

Public Property Get klasse() As String
    klasse = strKlasse
End Property

Public Property Let klasse(ByVal vNewValue As String)
    strKlasse = vNewValue
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

Public Property Get vornamen() As String
    vornamen = strVornamen
End Property

Public Property Let vornamen(ByVal vNewValue As String)
    strVornamen = vNewValue
End Property

Public Property Get geburtsdatum() As String
    geburtsdatum = strGebdat
End Property

Public Property Let geburtsdatum(ByVal vNewValue As String)
    strGebdat = vNewValue
End Property

Public Property Get geburtsort() As String
    geburtsort = strGebort
End Property

Public Property Let geburtsort(ByVal vNewValue As String)
    strGebort = vNewValue
End Property

Public Property Get geburtsland() As String
    geburtsland = strGebland
End Property

Public Property Let geburtsland(ByVal vNewValue As String)
    strGebland = vNewValue
End Property

Public Property Get bekenntnis() As String
    bekenntnis = strBekenntnis
End Property

Public Property Let bekenntnis(ByVal vNewValue As String)
    strBekenntnis = vNewValue
End Property

Public Property Get religion() As String
    religion = strReligion
End Property

Public Property Let religion(ByVal vNewValue As String)
    strReligion = vNewValue
End Property

Public Property Get geschlecht() As Long
    geschlecht = lngGeschlecht
End Property

Public Property Let geschlecht(ByVal vNewValue As Long)
    lngGeschlecht = vNewValue
End Property

Public Property Get jahrgangsstufe() As Long
    jahrgangsstufe = lngJgst
End Property

Public Property Let jahrgangsstufe(ByVal vNewValue As Long)
    lngJgst = vNewValue
End Property

Public Property Get ausbildungsrichtung() As Long
    ausbildungsrichtung = lngAusb
End Property

Public Property Let ausbildungsrichtung(ByVal vNewValue As Long)
    lngAusb = vNewValue
End Property

Public Property Get ausbildungsrichtung_kurz() As String
    ausbildungsrichtung_kurz = strAusbKurz
End Property

Public Property Let ausbildungsrichtung_kurz(ByVal vNewValue As String)
    strAusbKurz = vNewValue
End Property

Public Property Get eintrittdatum() As String
    eintrittdatum = strEintrittsdatum
End Property

Public Property Let eintrittdatum(ByVal vNewValue As String)
    strEintrittsdatum = vNewValue
End Property

Public Property Get probezeit() As String
    probezeit = strProbezeitbis
End Property

Public Property Let probezeit(ByVal vNewValue As String)
    strProbezeitbis = vNewValue
End Property

Public Property Get sprache() As Long
    sprache = lngSprache
End Property

Public Property Let sprache(ByVal vNewValue As Long)
    lngSprache = vNewValue
End Property

