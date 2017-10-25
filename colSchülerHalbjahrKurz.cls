VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSchülerHalbjahrKurz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle Halbjahresleistungen eines Schülers in der Kurzfassung
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, _
                    ByVal lngSchuelerUid As Long, ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal lngJahrgangsstufeUid As Long, _
                    ByVal lngGeloescht As Long) As clsSchülerHalbjahrKurz
  ' Neues Objekt erstellen
Dim objNewMember As clsSchülerHalbjahrKurz
    Set objNewMember = New clsSchülerHalbjahrKurz
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .schueler_uid = lngSchuelerUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .jahrgangsstufe_uid = lngJahrgangsstufeUid
        .geloescht = lngGeloescht
   End With

 ' das Objekt schließlich zur Collection hinzufügen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
   Set Add = objNewMember

 ' Objekt löschen
   Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsSchülerHalbjahrKurz
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



