VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Schule
Private strSchulnummer As String
Private strSchulart As String
Private strSchulname As String
Private strSchulnamekurz As String

Public Property Get schulnummer() As String
    schulnummer = strSchulnummer
End Property

Public Property Let schulnummer(ByVal vNewValue As String)
    strSchulnummer = vNewValue
End Property

Public Property Get schulart() As String
    schulart = strSchulart
End Property

Public Property Let schulart(ByVal vNewValue As String)
    strSchulart = vNewValue
End Property

Public Property Get schulname() As String
    schulname = strSchulname
End Property

Public Property Let schulname(ByVal vNewValue As String)
    strSchulname = vNewValue
End Property

Public Property Get schulnamekurz() As String
    schulnamekurz = strSchulnamekurz
End Property

Public Property Let schulnamekurz(ByVal vNewValue As String)
    strSchulnamekurz = vNewValue
End Property

