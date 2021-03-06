VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSch�lerHalbjahr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection f�r alle Sch�ler eines Faches mit ihren Halbjahresleistungen
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, _
                    ByVal lngSchuelerUid As Long, ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal lngJahrgangsstufeUid As Long, _
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
                    ByVal boo13Hj2 As Boolean) As clsSch�lerHalbjahr
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsSch�lerHalbjahr
    Set objNewMember = New clsSch�lerHalbjahr
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .schueler_uid = lngSchuelerUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .jahrgangsstufe_uid = lngJahrgangsstufeUid
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

 ' das Objekt schlie�lich zur Collection hinzuf�gen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zur�ckgeben
   Set Add = objNewMember

 ' Objekt l�schen
   Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsSch�lerHalbjahr
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


