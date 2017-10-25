VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colHalbjahrNotenKomplett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle Halbjahresleistungen Komplett zum Importieren
Private mCol As Collection
Public Function Add(ByVal lngSchuelernummer As Long, ByVal strFach As String, _
                    ByVal lngHalbjahr1 As Long, ByVal lngHalbjahr2 As Long, _
                    ByVal lngPruefung As Long, ByVal strName As String, ByVal strRufname As String, _
                    ByVal strSchuljahr As String, ByVal strJahrgang As String, ByVal strKlasse As String) As clsHalbjahrNotenKomplett
  ' Neues Halbjahr-Objekt erstellen
Dim objNewMember As clsHalbjahrNotenKomplett
    Set objNewMember = New clsHalbjahrNotenKomplett
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .schuelernummer = lngSchuelernummer
        .fach = strFach
        .halbjahr1 = lngHalbjahr1
        .halbjahr2 = lngHalbjahr2
        .pruefung = lngPruefung
        .name = strName
        .rufname = strRufname
        .schuljahr = strSchuljahr
        .jahrgangsstufe = strJahrgang
        .klasse = strKlasse
    End With

    ' das Objekt schließlich zur Collection hinzufügen
    mCol.Add objNewMember

    ' das neu erstellte Objekt zurückgeben
    Set Add = objNewMember

    ' Objekt löschen
    Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsHalbjahrNotenKomplett
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

