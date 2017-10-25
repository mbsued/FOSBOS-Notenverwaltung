VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHalbjahrNotenKomplett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private lngSchuelernummer As Long
Private strFach As String
Private lngHalbjahr1 As Long
Private lngHalbjahr2 As Long
Private lngPruefung As Long
Private strName As String
Private strRufname As String
Private strSchuljahr As String
Private lngJahrgang As Long
Private strKlasse As String

Public Property Get schuelernummer() As Long
    schuelernummer = lngSchuelernummer
End Property

Public Property Let schuelernummer(ByVal vNewValue As Long)
    lngSchuelernummer = vNewValue
End Property

Public Property Get fach() As String
    fach = strFach
End Property

Public Property Let fach(ByVal vNewValue As String)
    strFach = vNewValue
End Property

Public Property Get halbjahr1() As Long
    halbjahr1 = lngHalbjahr1
End Property

Public Property Let halbjahr1(ByVal vNewValue As Long)
    lngHalbjahr1 = vNewValue
End Property

Public Property Get halbjahr2() As Long
    halbjahr2 = lngHalbjahr2
End Property

Public Property Let halbjahr2(ByVal vNewValue As Long)
    lngHalbjahr2 = vNewValue
End Property

Public Property Get pruefung() As Long
    pruefung = lngPruefung
End Property

Public Property Let pruefung(ByVal vNewValue As Long)
    lngPruefung = vNewValue
End Property

Public Property Get name() As String
    name = strName
End Property

Public Property Let name(ByVal vNewValue As String)
    strName = vNewValue
End Property

Public Property Get rufname() As String
    rufname = strRufname
End Property

Public Property Let rufname(ByVal vNewValue As String)
    strRufname = vNewValue
End Property

Public Property Get schuljahr() As String
    schuljahr = strSchuljahr
End Property

Public Property Let schuljahr(ByVal vNewValue As String)
    strSchuljahr = vNewValue
End Property

Public Property Get jahrgangsstufe() As Long
    jahrgangsstufe = lngJahrgang
End Property

Public Property Let jahrgangsstufe(ByVal vNewValue As Long)
    lngJahrgang = vNewValue
End Property

Public Property Get klasse() As String
    klasse = strKlasse
End Property

Public Property Let klasse(ByVal vNewValue As String)
    strKlasse = vNewValue
End Property



