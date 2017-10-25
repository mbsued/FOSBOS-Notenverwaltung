VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchülerHalbjahrKurz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Schüler und seine Haljahresleistungen Kurzfassung
Private lngUid As Long
Private lngSchuelerUid As Long
Private lngFachUid As Long
Private lngKlassengruppeUid As Long
Private lngJahrgangsstufeUid As Long
Private lngGeloescht As Long

Public Property Get uid() As Long
    uid = lngUid
End Property

Public Property Let uid(ByVal vNewValue As Long)
    lngUid = vNewValue
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

Public Property Let jahrgangsstufe_uid(ByVal vNewValue As Long)
    lngJahrgangsstufeUid = vNewValue
End Property

Public Property Get jahrgangsstufe_uid() As Long
    jahrgangsstufe_uid = lngJahrgangsstufeUid
End Property

Public Property Get geloescht() As Long
    geloescht = lngGeloescht
End Property

Public Property Let geloescht(ByVal vNewValue As Long)
    lngGeloescht = vNewValue
End Property

