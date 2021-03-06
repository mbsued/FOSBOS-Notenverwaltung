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
' Collection f�r alle eingelesenen Lehrer
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

 ' das Objekt schlie�lich zur Collection hinzuf�gen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zur�ckgeben
   Set Add = objNewMember

 ' Objekt l�schen
   Set objNewMember = Nothing
   
Exit_Lehrer_Add:
    Exit Function
    
Err_Lehrer_Add:
    FehlermeldungAusgeben "Klasse Lehrer", "Hinzuf�gen der Lehrkraft", Err.Number, Err.Description
    Resume Exit_Lehrer_Add
    
End Function
Public Property Get Item(ByVal index As Long) As clsLehrer
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


