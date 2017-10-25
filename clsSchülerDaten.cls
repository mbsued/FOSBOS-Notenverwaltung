VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchülerdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für die Formulardaten des Schülers
Private lngUid As Long
Private strNachname As String
Private strRufname As String
Private strVornamen As String
Private strGebdat As String
Private strGebort As String
Private strGebland As String
Private strBekenntnis As String
Private strReligion As String
Private lngGeschlecht As Long
Private lngKlassengruppe As Long
Private booAusgeschieden As Boolean
Private strAusgeschiedendatum As String
Private strEintrittsdatum As String
Private strProbezeit As String
Private booNotenschutz As Boolean

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

Public Property Get klassengruppe() As Long
    klassengruppe = lngKlassengruppe
End Property

Public Property Let klassengruppe(ByVal vNewValue As Long)
    lngKlassengruppe = vNewValue
End Property

Public Property Get ausgeschieden() As Boolean
    ausgeschieden = booAusgeschieden
End Property

Public Property Let ausgeschieden(ByVal vNewValue As Boolean)
    booAusgeschieden = vNewValue
End Property

Public Property Get ausgeschiedenam() As String
    ausgeschiedenam = strAusgeschiedendatum
End Property

Public Property Let ausgeschiedenam(ByVal vNewValue As String)
    strAusgeschiedendatum = vNewValue
End Property

Public Property Get eintritt() As String
    eintritt = strEintrittsdatum
End Property

Public Property Let eintritt(ByVal vNewValue As String)
    strEintrittsdatum = vNewValue
End Property

Public Property Get probezeit() As String
    probezeit = strProbezeit
End Property

Public Property Let probezeit(ByVal vNewValue As String)
    strProbezeit = vNewValue
End Property

Public Property Get notenschutz() As Boolean
    notenschutz = booNotenschutz
End Property

Public Property Let notenschutz(ByVal vNewValue As Boolean)
    booNotenschutz = vNewValue
End Property


