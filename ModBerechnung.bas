Attribute VB_Name = "ModBerechnung"
Option Compare Database
Option Explicit

Public Function DurchschnittSonstigeLeistungBerechnen(lngGew1 As Variant, lngNote1 As Variant, _
                                                        lngGew2 As Variant, lngNote2 As Variant, _
                                                        lngGew3 As Variant, lngNote3 As Variant, _
                                                        lngGew4 As Variant, lngNote4 As Variant, _
                                                        lngGew5 As Variant, lngNote5 As Variant, _
                                                        lngGew6 As Variant, lngNote6 As Variant, _
                                                        lngGew7 As Variant, lngNote7 As Variant) As Double
' Berechnen des Durchschnitts der SonstigenLeistungen
' Übergeben werden alle 7 Gewichtungen und Noten

Dim dblSummeSonstigeLeistungen      'Summe der sonstigen Leistungen
Dim dblAnzahlSonstigeLeistungen     'Anzahl der Sonstigen Leistungen ermittelt aus den Gewichtungen

    dblSummeSonstigeLeistungen = 0
    dblAnzahlSonstigeLeistungen = 0
    
    If lngGew1 > 0 And (lngNote1 >= 0 And lngNote1 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew1 * lngNote1)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew1
    End If
    
    If lngGew2 > 0 And (lngNote2 >= 0 And lngNote2 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew2 * lngNote2)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew2
    End If
    
    If lngGew3 > 0 And (lngNote3 >= 0 And lngNote3 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew3 * lngNote3)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew3
    End If
    
    If lngGew4 > 0 And (lngNote4 >= 0 And lngNote4 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew4 * lngNote4)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew4
    End If
    
    If lngGew5 > 0 And (lngNote5 >= 0 And lngNote5 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew5 * lngNote5)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew5
    End If
    
    If lngGew6 > 0 And (lngNote6 >= 0 And lngNote6 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew6 * lngNote6)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew6
    End If
    
    If lngGew7 > 0 And (lngNote7 >= 0 And lngNote7 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew7 * lngNote7)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew7
    End If
    
    If dblAnzahlSonstigeLeistungen > 0 And dblSummeSonstigeLeistungen > 0 Then
        DurchschnittSonstigeLeistungBerechnen = dblSummeSonstigeLeistungen / dblAnzahlSonstigeLeistungen
    Else
        If dblAnzahlSonstigeLeistungen > 0 And dblSummeSonstigeLeistungen = 0 Then
            DurchschnittSonstigeLeistungBerechnen = 0
        Else
            DurchschnittSonstigeLeistungBerechnen = -1
        End If
    End If
End Function
Public Function DurchschnittFpaLeistungBerechnen(lngGew1 As Variant, lngNote1 As Variant, _
                                                        lngGew2 As Variant, lngNote2 As Variant, _
                                                        lngGew3 As Variant, lngNote3 As Variant) As Double
' Berechnen des Durchschnitts der SonstigenLeistungen in der FPA
' Übergeben werden 3 Gewichtungen und Noten

Dim dblSummeSonstigeLeistungen      'Summe der sonstigen Leistungen
Dim dblAnzahlSonstigeLeistungen     'Anzahl der Sonstigen Leistungen ermittelt aus den Gewichtungen

    dblSummeSonstigeLeistungen = 0
    dblAnzahlSonstigeLeistungen = 0
    
    If lngNote1 = 0 Or lngNote2 = 0 Or lngNote3 = 0 Then  'wenn in einer Leistung 0, dann Gesamtergebnis 0
        DurchschnittFpaLeistungBerechnen = 0
        Exit Function
    End If
    
    If lngGew1 > 0 And (lngNote1 >= 0 And lngNote1 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew1 * lngNote1)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew1
    End If
    
    If lngGew2 > 0 And (lngNote2 >= 0 And lngNote2 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew2 * lngNote2)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew2
    End If
    
    If lngGew3 > 0 And (lngNote3 >= 0 And lngNote3 < 16) Then
        dblSummeSonstigeLeistungen = dblSummeSonstigeLeistungen + (lngGew3 * lngNote3)
        dblAnzahlSonstigeLeistungen = dblAnzahlSonstigeLeistungen + lngGew3
    End If
    
    If dblAnzahlSonstigeLeistungen > 0 And dblSummeSonstigeLeistungen > 0 Then
        DurchschnittFpaLeistungBerechnen = dblSummeSonstigeLeistungen / dblAnzahlSonstigeLeistungen
    Else
        If dblAnzahlSonstigeLeistungen > 0 And dblSummeSonstigeLeistungen = 0 Then
            DurchschnittFpaLeistungBerechnen = 0
        Else
            DurchschnittFpaLeistungBerechnen = -1
        End If
    End If
End Function
Public Function SchulaufgabenBerechnen(lngNote1 As Variant, lngNote2 As Variant, lngAnzahlSa As Long) As Double
' Berechnen der Schulaufgaben
' Übergabe der 2 Noten
Dim lngSummeSchulaufgaben As Long      'Summe der Schulaufgaben
Dim lngAnzahlSchulaufgaben As Long      'Anzahl der Schulaufgaben
Dim booSchulaufgabenNoteDa As Boolean

    lngSummeSchulaufgaben = 0
    lngAnzahlSchulaufgaben = 0
    booSchulaufgabenNoteDa = False
    
    If (Not IsNull(lngNote1) And Not lngNote1 = "" And Not lngNote1 = "-") Then
        lngSummeSchulaufgaben = lngSummeSchulaufgaben + lngNote1
        lngAnzahlSchulaufgaben = lngAnzahlSchulaufgaben + 1
        booSchulaufgabenNoteDa = True
    End If
    
    If (Not IsNull(lngNote2) And Not lngNote2 = "" And Not lngNote2 = "-") Then
        lngSummeSchulaufgaben = lngSummeSchulaufgaben + lngNote2
        lngAnzahlSchulaufgaben = lngAnzahlSchulaufgaben + 1
        booSchulaufgabenNoteDa = True
    End If
    
    If booSchulaufgabenNoteDa = True Then
        If lngAnzahlSchulaufgaben = lngAnzahlSa Then    'nur sinnvoll wenn alle Schulaufgaben gefüllt sind
            SchulaufgabenBerechnen = lngSummeSchulaufgaben
        Else
            SchulaufgabenBerechnen = -1
        End If
    Else
        SchulaufgabenBerechnen = -1
    End If
End Function


