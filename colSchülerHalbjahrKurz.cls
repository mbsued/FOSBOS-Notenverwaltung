VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSch�lerHalbjahrKurz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection f�r alle Halbjahresleistungen eines Sch�lers in der Kurzfassung
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, _
                    ByVal lngSchuelerUid As Long, ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal lngJahrgangsstufeUid As Long, _
                    ByVal lngGeloescht As Long) As clsSch�lerHalbjahrKurz
  ' Neues Objekt erstellen
Dim objNewMember As clsSch�lerHalbjahrKurz
    Set objNewMember = New clsSch�lerHalbjahrKurz
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .schueler_uid = lngSchuelerUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .jahrgangsstufe_uid = lngJahrgangsstufeUid
        .geloescht = lngGeloescht
   End With

 ' das Objekt schlie�lich zur Collection hinzuf�gen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zur�ckgeben
   Set Add = objNewMember

 ' Objekt l�schen
   Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsSch�lerHalbjahrKurz
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



