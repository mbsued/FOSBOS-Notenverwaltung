VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colKlasseFächer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle eingelesenen Fächer einer Klasse
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, ByVal lngFachUid, ByVal lngKlassengruppeUid As Long, _
                    ByVal strFach As String, ByVal booGeloescht As Boolean) As clsKlasseFächer
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsKlasseFächer
    Set objNewMember = New clsKlasseFächer
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .fachbezeichnung = strFach
        .geloescht = booGeloescht
    End With

 ' das Objekt schließlich zur Collection hinzufügen
    mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
    Set Add = objNewMember

 ' Objekt löschen
    Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsKlasseFächer
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


