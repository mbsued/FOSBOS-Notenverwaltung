VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "colSchülerNoten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Collection für alle Schüler einer Klassengruppe und eines Faches mit ihren Noten
Private mCol As Collection
Public Function Add(ByVal lngUid As Long, ByVal strNachname As String, ByVal strRufname As String, _
                    ByVal lngSchuelerUid As Long, ByVal lngFachUid As Long, ByVal lngKlassengruppeUid As Long, ByVal booIndEinstellung As Boolean, _
                    ByVal varAnzLeistungsnachweise As Variant, _
                    ByVal varArtLeistungsnachweise As Variant, _
                    ByVal varDatumLeistungsnachweise As Variant, _
                    ByVal varGewLeistungsnachweise As Variant, _
                    ByVal varNotenLeistungsnachweise As Variant, _
                    ByVal booGeloescht As Boolean) As clsSchülerNoten
  ' Neues Termin-Objekt erstellen
Dim objNewMember As clsSchülerNoten
    Set objNewMember = New clsSchülerNoten
   
    ' Daten an das erstellte Objekt zuweisen
    With objNewMember
        .uid = lngUid
        .nachname = strNachname
        .rufname = strRufname
        .schueler_uid = lngSchuelerUid
        .fach_uid = lngFachUid
        .klassengruppe_uid = lngKlassengruppeUid
        .ind_einstellung = booIndEinstellung
        .anz_sa_hj1 = varAnzLeistungsnachweise(0)
        .anz_sa_hj2 = varAnzLeistungsnachweise(1)
        .anz_son_hj1 = varAnzLeistungsnachweise(2)
        .anz_son_hj2 = varAnzLeistungsnachweise(3)
        .uid_art_son1_hj1 = varArtLeistungsnachweise(0)
        .uid_art_son2_hj1 = varArtLeistungsnachweise(1)
        .uid_art_son3_hj1 = varArtLeistungsnachweise(2)
        .uid_art_son4_hj1 = varArtLeistungsnachweise(3)
        .uid_art_son5_hj1 = varArtLeistungsnachweise(4)
        .uid_art_son6_hj1 = varArtLeistungsnachweise(5)
        .uid_art_son7_hj1 = varArtLeistungsnachweise(6)
        .uid_art_son1_hj2 = varArtLeistungsnachweise(7)
        .uid_art_son2_hj2 = varArtLeistungsnachweise(8)
        .uid_art_son3_hj2 = varArtLeistungsnachweise(9)
        .uid_art_son4_hj2 = varArtLeistungsnachweise(10)
        .uid_art_son5_hj2 = varArtLeistungsnachweise(11)
        .uid_art_son6_hj2 = varArtLeistungsnachweise(12)
        .uid_art_son7_hj2 = varArtLeistungsnachweise(13)
        .dat_sa1_hj1 = varDatumLeistungsnachweise(0)
        .dat_sa2_hj1 = varDatumLeistungsnachweise(1)
        .dat_sa1_hj2 = varDatumLeistungsnachweise(2)
        .dat_sa2_hj2 = varDatumLeistungsnachweise(3)
        .dat_son1_hj1 = varDatumLeistungsnachweise(4)
        .dat_son2_hj1 = varDatumLeistungsnachweise(5)
        .dat_son3_hj1 = varDatumLeistungsnachweise(6)
        .dat_son4_hj1 = varDatumLeistungsnachweise(7)
        .dat_son5_hj1 = varDatumLeistungsnachweise(8)
        .dat_son6_hj1 = varDatumLeistungsnachweise(9)
        .dat_son7_hj1 = varDatumLeistungsnachweise(10)
        .dat_son1_hj2 = varDatumLeistungsnachweise(11)
        .dat_son2_hj2 = varDatumLeistungsnachweise(12)
        .dat_son3_hj2 = varDatumLeistungsnachweise(13)
        .dat_son4_hj2 = varDatumLeistungsnachweise(14)
        .dat_son5_hj2 = varDatumLeistungsnachweise(15)
        .dat_son6_hj2 = varDatumLeistungsnachweise(16)
        .dat_son7_hj1 = varDatumLeistungsnachweise(17)
        .gew_sa1_hj1 = varGewLeistungsnachweise(0)
        .gew_sa2_hj1 = varGewLeistungsnachweise(1)
        .gew_sa1_hj2 = varGewLeistungsnachweise(2)
        .gew_sa2_hj2 = varGewLeistungsnachweise(3)
        .gew_son1_hj1 = varGewLeistungsnachweise(4)
        .gew_son2_hj1 = varGewLeistungsnachweise(5)
        .gew_son3_hj1 = varGewLeistungsnachweise(6)
        .gew_son4_hj1 = varGewLeistungsnachweise(7)
        .gew_son5_hj1 = varGewLeistungsnachweise(8)
        .gew_son6_hj1 = varGewLeistungsnachweise(9)
        .gew_son7_hj1 = varGewLeistungsnachweise(10)
        .gew_son1_hj2 = varGewLeistungsnachweise(11)
        .gew_son2_hj2 = varGewLeistungsnachweise(12)
        .gew_son3_hj2 = varGewLeistungsnachweise(13)
        .gew_son4_hj2 = varGewLeistungsnachweise(14)
        .gew_son5_hj2 = varGewLeistungsnachweise(15)
        .gew_son6_hj2 = varGewLeistungsnachweise(16)
        .gew_son7_hj1 = varGewLeistungsnachweise(17)
        .n_sa1_hj1 = varNotenLeistungsnachweise(0)
        .n_sa2_hj1 = varNotenLeistungsnachweise(1)
        .n_sa1_hj2 = varNotenLeistungsnachweise(2)
        .n_sa2_hj2 = varNotenLeistungsnachweise(3)
        .n_son1_hj1 = varNotenLeistungsnachweise(4)
        .n_son2_hj1 = varNotenLeistungsnachweise(5)
        .n_son3_hj1 = varNotenLeistungsnachweise(6)
        .n_son4_hj1 = varNotenLeistungsnachweise(7)
        .n_son5_hj1 = varNotenLeistungsnachweise(8)
        .n_son6_hj1 = varNotenLeistungsnachweise(9)
        .n_son7_hj1 = varNotenLeistungsnachweise(10)
        .n_son1_hj2 = varNotenLeistungsnachweise(11)
        .n_son2_hj2 = varNotenLeistungsnachweise(12)
        .n_son3_hj2 = varNotenLeistungsnachweise(13)
        .n_son4_hj2 = varNotenLeistungsnachweise(14)
        .n_son5_hj2 = varNotenLeistungsnachweise(15)
        .n_son6_hj2 = varNotenLeistungsnachweise(16)
        .n_son7_hj2 = varNotenLeistungsnachweise(17)
        .geloescht = booGeloescht
   End With

 ' das Objekt schließlich zur Collection hinzufügen
   mCol.Add objNewMember

 ' das neu erstellte Objekt zurückgeben
   Set Add = objNewMember

 ' Objekt löschen
   Set objNewMember = Nothing
End Function
Public Property Get Item(ByVal index As Long) As clsSchülerNoten
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


