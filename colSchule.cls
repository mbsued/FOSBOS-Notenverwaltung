VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSchule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection f�r alle eingelesenen Schulen
Private mCol As Collection
Public Function Add(ByVal strSchulnummer As String, ByVal strSchulart As String, ByVal strSchulname As String, ByVal strSchulnamekurz As String) As clsSchule
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsSchule
    On Error GoTo Err_Schule_Add
    
    Set objNewMember = New clsSchule
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .schulnummer = strSchulnummer
        .schulart = strSchulart
        .schulname = strSchulname
        .schulnamekurz = strSchulnamekurz
    End With

 ' das Objekt schlie�lich zur Collection hinzuf�gen
    mCol.Add objNewMember

 ' das neu erstellte Objekt zur�ckgeben
    Set Add = objNewMember

 ' Objekt l�schen
    Set objNewMember = Nothing
   
Exit_Schule_Add:
    Exit Function
    
Err_Schule_Add:
    FehlermeldungAusgeben "Klasse Schule", "Hinzuf�gen der Schule", Err.Number, Err.Description
    Resume Exit_Schule_Add
    
End Function
Public Property Get Item(ByVal index As Long) As clsSchule
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


