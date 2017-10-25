Attribute VB_Name = "ModBerichtFunktionen"
Option Compare Database
Option Explicit
Public Function rptArtderLeistungserhebung(lngArt As Long) As String
' Funktion liefert anhand der übergebenen Zahl, den Klartext der Art der Leistungserhebung zurück

    On Error GoTo Err_ArtderLeistungserhebung
    
    Select Case lngArt
            Case 0  'kein Eintrag
                rptArtderLeistungserhebung = ""
                Exit Function
            Case 1  'Kurzarbeit
                rptArtderLeistungserhebung = "KA"
                Exit Function
            Case 2  'Stegreifaufgabe
                rptArtderLeistungserhebung = "Ex"
                Exit Function
            Case 3  'mündliche Leistung
                rptArtderLeistungserhebung = "mdl"
                Exit Function
            Case 4  'Ersatzprüfung
                rptArtderLeistungserhebung = "ErPr"
                Exit Function
            Case 5  'fachpraktische Tätigkeit
                rptArtderLeistungserhebung = "fpT"
                Exit Function
            Case 6  'fachpraktische Anleitung
                rptArtderLeistungserhebung = "fpAn"
                Exit Function
            Case 7  'fachpraktische Vertiefung
                rptArtderLeistungserhebung = "fpV"
                Exit Function
            Case Else
                rptArtderLeistungserhebung = "ubk"
    End Select
    
    
Exit_ArtderLeistungserhebung:
    Exit Function
    
Err_ArtderLeistungserhebung:
    FehlermeldungAusgeben "Art der Leistungserhebung", Err.Source, Err.Number, Err.Description
    Resume Exit_ArtderLeistungserhebung
    
End Function
Public Function rptSchuljahr() As String

Dim lngAktuellesJahr As Long
Dim lngVorherigesJahr As Long
Dim lngNaechstesJahr As Long
    
    lngAktuellesJahr = Format(Date, "YYYY")
    lngNaechstesJahr = lngAktuellesJahr + 1
    lngVorherigesJahr = lngAktuellesJahr - 1
    
    If Format(Date, "M") > 7 Then          'wir sind im Monat 8 - 12
        rptSchuljahr = "Schuljahr " & lngAktuellesJahr & "/" & lngNaechstesJahr
    Else
        rptSchuljahr = "Schuljahr " & lngVorherigesJahr & "/" & lngAktuellesJahr
    End If
    

End Function
Public Function rptSchueler(lngGeschlechtUid As Long, strNachname As String, strRufname As String) As String
' Schüler aufbauen

Dim StrSchueler As String

    If lngGeschlechtUid = 1 Then
        StrSchueler = "Frau "
    Else
        StrSchueler = "Herr "
    End If
    
    StrSchueler = StrSchueler & strNachname & ", " & strRufname
    
    rptSchueler = StrSchueler
End Function
Public Function rptJahresergebnis(lngHalbjahr1 As Long, lngHalbjahr2 As Long) As String
' Jahresergebnis berechnen

Dim dblsumme As Double
Dim dblErgebnis As Double

    If Not lngHalbjahr1 = -1 And Not lngHalbjahr2 = -1 Then
        dblsumme = CDbl(lngHalbjahr1) + CDbl(lngHalbjahr2)
        dblErgebnis = dblsumme / 2
        If dblErgebnis < 1 Then
            rptJahresergebnis = 0
        Else
            rptJahresergebnis = Runden(dblErgebnis, 0)
        End If
    Else
        rptJahresergebnis = ""
    End If
    
End Function
