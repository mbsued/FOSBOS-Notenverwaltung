VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKlasseFächer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Klasse und seine Fächer
Private lngUid As Long
Private lngFachUid As Long
Private lngKlassengruppeUid As Long
Private booGeloescht As Boolean
Private strFach As String
Public Property Get uid() As Long
    uid = lngUid
End Property

Public Property Let uid(ByVal vNewValue As Long)
    lngUid = vNewValue
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

Public Property Get fachbezeichnung() As String
    fachbezeichnung = strFach
End Property

Public Property Let fachbezeichnung(ByVal vNewValue As String)
    strFach = vNewValue
End Property

Public Property Get geloescht() As Boolean
    geloescht = booGeloescht
End Property

Public Property Let geloescht(ByVal vNewValue As Boolean)
    booGeloescht = vNewValue
End Property

