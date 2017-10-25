VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLehrer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für das Objekt Lehrer
Private strNachname As String
Private strRufname As String
Private strKuerzel As String
Private strAmt As String
Private strTitel As String
Private lngGeschlecht As Long

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

Public Property Get kuerzel() As String
    kuerzel = strKuerzel
End Property

Public Property Let kuerzel(ByVal vNewValue As String)
    strKuerzel = vNewValue
End Property

Public Property Get amt() As String
    amt = strAmt
End Property

Public Property Let amt(ByVal vNewValue As String)
    strAmt = vNewValue
End Property

Public Property Get titel() As String
    titel = strTitel
End Property

Public Property Let titel(ByVal vNewValue As String)
    strTitel = vNewValue
End Property

Public Property Get geschlecht() As Long
    geschlecht = lngGeschlecht
End Property

Public Property Let geschlecht(ByVal vNewValue As Long)
    lngGeschlecht = vNewValue
End Property


