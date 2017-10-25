VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHalbjahrNoten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private lngSchuelernummer As Long
Private lngHalbjahr1 As Long
Private lngHalbjahr2 As Long

Public Property Get schuelernummer() As Long
    schuelernummer = lngSchuelernummer
End Property

Public Property Let schuelernummer(ByVal vNewValue As Long)
    lngSchuelernummer = vNewValue
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


