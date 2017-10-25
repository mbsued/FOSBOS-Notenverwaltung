VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colFaecherHalbjahr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle Fächer eines Schülers mit ihren Halbjahresleistungen
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, _
                    ByVal lngSchuelerUid As Long, ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal lngJahrgangsstufeUid As Long, ByVal lngSchulartUid, _
                    ByVal strFach As String, _
                    ByVal varVklHj1 As Variant, _
                    ByVal varVklHj2 As Variant, _
                    ByVal var11Hj1 As Variant, _
                    ByVal var11Hj2 As Variant, _
                    ByVal var12Hj1 As Variant, _
                    ByVal var12Hj2 As Variant, _
                    ByVal var13Hj1 As Variant, _
                    ByVal var13Hj2 As Variant, _
                    ByVal boo11Hj1 As Boolean, _
                    ByVal boo11Hj2 As Boolean, _
                    ByVal boo12Hj1 As Boolean, _
                    ByVal boo12Hj2 As Boolean, _
                    ByVal boo13Hj1 As Boolean, _
                    ByVal boo13Hj2 As Boolean) As clsFaecherHalbjahr
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsFaecherHalbjahr
    Set objNewMember = New clsFaecherHalbjahr
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .schueler_uid = lngSchuelerUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .jahrgangsstufe_uid = lngJahrgangsstufeUid
        .schulart_uid = lngSchulartUid
        .fach = strFach
        .n_vkl_hj1 = varVklHj1
        .n_vkl_hj2 = varVklHj2
        .n_11_hj1 = var11Hj1
        .n_11_hj2 = var11Hj2
        .n_12_hj1 = var12Hj1
        .n_12_hj2 = var12Hj2
        .n_13_hj1 = var13Hj1
        .n_13_hj2 = var13Hj2
        .ein_11_hj1 = boo11Hj1
        .ein_11_hj2 = boo11Hj2
        .ein_12_hj1 = boo12Hj1
        .ein_12_hj2 = boo12Hj2
        .ein_13_hj1 = boo13Hj1
        .ein_13_hj2 = boo13Hj2
   End With

 ' das Objekt schließlich zur Collection hinzufügen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
   Set Add = objNewMember

 ' Objekt löschen
   Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsFaecherHalbjahr
  ' Bestimmten Datensatz der Collection zurückgeben.
  ' Der Datensatz wird über die Position (Index) "angesprochen".
  Set Item = mCol(index)
End Property
Public Property Get Count() As Long
  ' Anzahl Collections-Einträge
  Count = mCol.Count
End Property
Public Sub Remove(ByVal index As Integer)
  ' Löschen eines Eintrags aus dem Collection-Objekt
  mCol.Remove index
End Sub

Private Sub Class_Initialize()
' Erstellt das Collection-Objekt beim ersten Zugriff
  ' auf die Klasse
  Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
  ' Zerstört das Collection-Objekt,
  ' wenn die Klasse beendet wird
  Set mCol = Nothing
End Sub


