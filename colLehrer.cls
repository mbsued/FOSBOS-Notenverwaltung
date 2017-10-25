VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colLehrer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle eingelesenen Lehrer
Private mCol As Collection
Public Function Add(ByVal strNachname As String, ByVal strRufname As String, _
                    ByVal strKuerzel As String, ByVal strAmt As String, ByVal strTitel As String, _
                    ByVal lngGeschlecht As Long) As clsLehrer
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsLehrer

    On Error GoTo Err_Lehrer_Add
    
    Set objNewMember = New clsLehrer
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .nachname = strNachname
        .rufname = strRufname
        .kuerzel = strKuerzel
        .amt = strAmt
        .titel = strTitel
        .geschlecht = lngGeschlecht
  End With

 ' das Objekt schließlich zur Collection hinzufügen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
   Set Add = objNewMember

 ' Objekt löschen
   Set objNewMember = Nothing
   
Exit_Lehrer_Add:
    Exit Function
    
Err_Lehrer_Add:
    FehlermeldungAusgeben "Klasse Lehrer", "Hinzufügen der Lehrkraft", Err.Number, Err.Description
    Resume Exit_Lehrer_Add
    
End Function
Public Property Get Item(ByVal index As Long) As clsLehrer
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


