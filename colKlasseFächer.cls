VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colKlasseF�cher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection f�r alle eingelesenen F�cher einer Klasse
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, ByVal lngFachUid, ByVal lngKlassengruppeUid As Long, _
                    ByVal strFach As String, ByVal booGeloescht As Boolean) As clsKlasseF�cher
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsKlasseF�cher
    Set objNewMember = New clsKlasseF�cher
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .fachbezeichnung = strFach
        .geloescht = booGeloescht
    End With

 ' das Objekt schlie�lich zur Collection hinzuf�gen
    mCol.Add objNewMember

 ' das neu erstellte Objekt zur�ckgeben
    Set Add = objNewMember

 ' Objekt l�schen
    Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsKlasseF�cher
  ' Bestimmten Datensatz der Collection zur�ckgeben.
  ' Der Datensatz wird �ber die Position (Index) "angesprochen".
  Set Item = mCol(index)
End Property
Public Property Get Count() As Long
  ' Anzahl Collections-Eintr�ge
  Count = mCol.Count
End Property
Public Sub Remove(ByVal index As Integer)
  ' L�schen eines Eintrags aus dem Collection-Objekt
  mCol.Remove index
End Sub

Private Sub Class_Initialize()
' Erstellt das Collection-Objekt beim ersten Zugriff
  ' auf die Klasse
  Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
  ' Zerst�rt das Collection-Objekt,
  ' wenn die Klasse beendet wird
  Set mCol = Nothing
End Sub


