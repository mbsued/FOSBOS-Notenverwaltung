VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFaecherHalbjahr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Schüler und seine Haljahresleistungen
Private lngUid As Long
Private lngSchuelerUid As Long
Private lngFachUid As Long
Private lngKlassengruppeUid As Long
Private lngJahrgangsstufeUid As Long
Private lngSchulartUid As Long
Private strFach As String
'Variablen für Halbjahresleistungen
Private varVklHj1 As Variant
Private varVklHj2 As Variant
Private var11Hj1 As Variant
Private var11Hj2 As Variant
Private var12Hj1 As Variant
Private var12Hj2 As Variant
Private var13Hj1 As Variant
Private var13Hj2 As Variant
'Variablen für Halbjahresleistungen eingebracht
Private boo11Hj1 As Boolean
Private boo11Hj2 As Boolean
Private boo12Hj1 As Boolean
Private boo12Hj2 As Boolean
Private boo13Hj1 As Boolean
Private boo13Hj2 As Boolean

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

Public Property Let schulart_uid(ByVal vNewValue As Long)
    lngSchulartUid = vNewValue
End Property

Public Property Get schulart_uid() As Long
    schulart_uid = lngSchulartUid
End Property

Public Property Let fach(ByVal vNewValue As String)
    strFach = vNewValue
End Property

Public Property Get fach() As String
    fach = strFach
End Property

Public Property Get n_vkl_hj1() As Variant
    n_vkl_hj1 = varVklHj1
End Property

Public Property Let n_vkl_hj1(ByVal vNewValue As Variant)
    varVklHj1 = vNewValue
End Property

Public Property Get n_vkl_hj2() As Variant
    n_vkl_hj2 = varVklHj2
End Property

Public Property Let n_vkl_hj2(ByVal vNewValue As Variant)
    varVklHj2 = vNewValue
End Property

Public Property Get n_11_hj1() As Variant
    n_11_hj1 = var11Hj1
End Property

Public Property Let n_11_hj1(ByVal vNewValue As Variant)
    var11Hj1 = vNewValue
End Property

Public Property Get n_11_hj2() As Variant
    n_11_hj2 = var11Hj2
End Property

Public Property Let n_11_hj2(ByVal vNewValue As Variant)
    var11Hj2 = vNewValue
End Property

Public Property Get n_12_hj1() As Variant
    n_12_hj1 = var12Hj1
End Property

Public Property Let n_12_hj1(ByVal vNewValue As Variant)
    var12Hj1 = vNewValue
End Property

Public Property Get n_12_hj2() As Variant
    n_12_hj2 = var12Hj2
End Property

Public Property Let n_12_hj2(ByVal vNewValue As Variant)
    var12Hj2 = vNewValue
End Property

Public Property Get n_13_hj1() As Variant
    n_13_hj1 = var13Hj1
End Property

Public Property Let n_13_hj1(ByVal vNewValue As Variant)
    var13Hj1 = vNewValue
End Property

Public Property Get n_13_hj2() As Variant
    n_13_hj2 = var13Hj2
End Property

Public Property Let n_13_hj2(ByVal vNewValue As Variant)
    var13Hj2 = vNewValue
End Property

Public Property Get ein_11_hj1() As Boolean
    ein_11_hj1 = boo11Hj1
End Property

Public Property Let ein_11_hj1(ByVal vNewValue As Boolean)
    boo11Hj1 = vNewValue
End Property

Public Property Get ein_11_hj2() As Boolean
    ein_11_hj2 = boo11Hj2
End Property

Public Property Let ein_11_hj2(ByVal vNewValue As Boolean)
    boo11Hj2 = vNewValue
End Property

Public Property Get ein_12_hj1() As Boolean
    ein_12_hj1 = boo12Hj1
End Property

Public Property Let ein_12_hj1(ByVal vNewValue As Boolean)
    boo12Hj1 = vNewValue
End Property

Public Property Get ein_12_hj2() As Boolean
    ein_12_hj2 = boo12Hj2
End Property

Public Property Let ein_12_hj2(ByVal vNewValue As Boolean)
    boo12Hj2 = vNewValue
End Property

Public Property Get ein_13_hj1() As Boolean
    ein_13_hj1 = boo13Hj1
End Property

Public Property Let ein_13_hj1(ByVal vNewValue As Boolean)
    boo13Hj1 = vNewValue
End Property

Public Property Get ein_13_hj2() As Boolean
    ein_13_hj2 = boo13Hj2
End Property

Public Property Let ein_13_hj2(ByVal vNewValue As Boolean)
    boo13Hj2 = vNewValue
End Property


