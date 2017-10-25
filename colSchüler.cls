VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSchüler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle eingelesenen Schüler
Private mCol As Collection
Public Function Add(ByVal lngSchuelernummer As Long, ByVal lngSchulart As Long, ByVal strKlasse, ByVal strNachname As String, ByVal strRufname As String, _
                    ByVal strVornamen As String, ByVal strGebdat As String, ByVal strGebort As String, _
                    ByVal strGebland As String, ByVal strBekenntnis As String, _
                    ByVal strReligion As String, ByVal lngGeschlecht As Long, _
                    ByVal lngJgst As Long, ByVal lngAusb As Long, ByVal strAusbKurz As String, _
                    ByVal strEintrittsdatum As String, ByVal strProbezeit As String, ByVal strSprache As String) As clsSchüler
  ' Neues Schüler-Objekt erstellen
Dim objNewMember As clsSchüler

    On Error GoTo Err_Schueler_Add
    
    Set objNewMember = New clsSchüler
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .schuelernummer = lngSchuelernummer
        .schulart = lngSchulart
        .klasse = strKlasse
        .nachname = strNachname
        .rufname = strRufname
        .vornamen = strVornamen
        .geburtsdatum = strGebdat
        .geburtsort = strGebort
        .geburtsland = strGebland
        .bekenntnis = strBekenntnis
        .religion = strReligion
        .geschlecht = lngGeschlecht
        .jahrgangsstufe = lngJgst
        .ausbildungsrichtung = lngAusb
        .ausbildungsrichtung_kurz = strAusbKurz
        .eintrittdatum = strEintrittsdatum
        .probezeit = strProbezeit
        .sprache = strSprache
   End With

 ' das Objekt schließlich zur Collection hinzufügen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
   Set Add = objNewMember

 ' Objekt löschen
   Set objNewMember = Nothing
   
Exit_Schueler_Add:
    Exit Function
    
Err_Schueler_Add:
    FehlermeldungAusgeben "Klasse Schüler", "Hinzufügen des Schülers", Err.Number, Err.Description
    Resume Exit_Schueler_Add
    
End Function
Public Property Get Item(ByVal index As Long) As clsSchüler
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

