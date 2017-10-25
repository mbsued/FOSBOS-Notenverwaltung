VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLehrerDaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Klassenmodul für die Daten des Lehrers zum Ändern
Private lngUid As Long
Private strKuerzel As String
Private strNachname As String
Private strRufname As String
Private strAmt As String
Private strTitel As String
Private lngGeschlecht As Long
Private strEmail As String
Private booSchulleitung As Boolean

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

Public Property Get email() As String
    email = strEmail
End Property

Public Property Let email(ByVal vNewValue As String)
    strEmail = vNewValue
End Property

Public Property Get schulleitung() As Boolean
    schulleitung = booSchulleitung
End Property

Public Property Let schulleitung(ByVal vNewValue As Boolean)
    booSchulleitung = vNewValue
End Property

