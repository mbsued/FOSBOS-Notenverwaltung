VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKlasseDaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für die Daten der Klasse zum Ändern
Private lngUid As Long
Private strBezeichnung As String
Private lngKlassleitung As Long
Private lngZeugnisunterzeichner As Long

Public Property Get uid() As Long
    uid = lngUid
End Property

Public Property Let uid(ByVal vNewValue As Long)
    lngUid = vNewValue
End Property

Public Property Get bezeichnung() As String
    bezeichnung = strBezeichnung
End Property

Public Property Let bezeichnung(ByVal vNewValue As String)
    strBezeichnung = vNewValue
End Property

Public Property Get klassleitung() As Long
    klassleitung = lngKlassleitung
End Property

Public Property Let klassleitung(ByVal vNewValue As Long)
    lngKlassleitung = vNewValue
End Property

Public Property Get zeugnisunterzeichner() As Long
    zeugnisunterzeichner = lngZeugnisunterzeichner
End Property

Public Property Let zeugnisunterzeichner(ByVal vNewValue As Long)
    lngZeugnisunterzeichner = vNewValue
End Property

