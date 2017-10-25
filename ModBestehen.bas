Attribute VB_Name = "ModBestehen"
Option Compare Database
Option Explicit
Public Function BestehenPruefen(lngArt As Long, lngJahrgangsstufe As Long, mycolFaecherhalbjahr As colFaecherHalbjahr) As Long
'Überprüfen des Bestehens Art=1: Probezeit, Art=2: Schuljahr
'Rückgabewerte: 1 = Bestanden, 0 = nicht bestanden, -1 = Noten nicht vollständig
Dim lngRueckgabeWert As Long

    On Error GoTo Err_Bestehen
    
    Select Case lngArt
        Case 1  'Probezeit
            If NotenVollstaendig(1, lngJahrgangsstufe, mycolFaecherhalbjahr) Then
                If FpaBestanden(1, mycolFaecherhalbjahr) Then
                    If AlleJahresnotenGroesserDrei(1, lngJahrgangsstufe, mycolFaecherhalbjahr) Then
                        lngRueckgabeWert = 1
                    Else
                        lngRueckgabeWert = 0
                    End If
                Else
                    lngRueckgabeWert = 0
                End If
            Else
                lngRueckgabeWert = -1
            End If
        Case 2  'Schuljahr
            If NotenVollstaendig(2, lngJahrgangsstufe, mycolFaecherhalbjahr) Then
                If FpaBestanden(2, mycolFaecherhalbjahr) Then
                    If AlleJahresnotenGroesserDrei(2, lngJahrgangsstufe, mycolFaecherhalbjahr) Then
                        lngRueckgabeWert = 1
                    Else
                        lngRueckgabeWert = 0
                    End If
                Else
                    lngRueckgabeWert = 0
                End If
            Else
                lngRueckgabeWert = -1
            End If
    End Select
    
    BestehenPruefen = lngRueckgabeWert
    
Exit_Bestehen:
    Exit Function
    
Err_Bestehen:
    FehlermeldungAusgeben "Bestehen überprüfen", Err.Source, Err.Number, Err.Description
    Resume Exit_Bestehen
    
End Function
Private Function NotenVollstaendig(lngArt As Long, lngJahrgangsstufe As Long, mycolFaecherhalbjahr As colFaecherHalbjahr) As Boolean
' Prüft ob die Halbjahresnoten vollständig sind
' Art 1: Probezeit, Art 2: Schuljahr
Dim booVollstaendig As Boolean
Dim intCount As Integer

    Select Case lngArt
        Case 1
            Select Case lngJahrgangsstufe
                Case 1 'Vorklasse
                    booVollstaendig = True
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                    Next intCount
                Case 2 '11.Klasse
                    booVollstaendig = True
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                    Next intCount
                Case 3 '12.Klasse
                    booVollstaendig = False
                Case 4 '13.Klasse
                    booVollstaendig = False
            End Select
        Case 2
            Select Case lngJahrgangsstufe
                Case 1 'Vorklasse
                    booVollstaendig = True
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                        If mycolFaecherhalbjahr.Item(intCount).n_vkl_hj2 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                    Next intCount
                Case 2 '11.Klasse
                    booVollstaendig = True
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                        If mycolFaecherhalbjahr.Item(intCount).n_11_hj2 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                    Next intCount
                Case 3 '12.Klasse
                    booVollstaendig = True
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).schulart_uid = 1 Then
                            If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 = -1 Then
                                booVollstaendig = False
                                GoTo Exit_NotenVollstaendig
                            End If
                            If mycolFaecherhalbjahr.Item(intCount).n_11_hj2 = -1 Then
                                booVollstaendig = False
                                GoTo Exit_NotenVollstaendig
                            End If
                        End If
                        If mycolFaecherhalbjahr.Item(intCount).n_12_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                        If mycolFaecherhalbjahr.Item(intCount).n_12_hj2 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                    Next intCount
                Case 4 '13.Klasse
                    booVollstaendig = True
                        If mycolFaecherhalbjahr.Item(intCount).n_13_hj1 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
                        If mycolFaecherhalbjahr.Item(intCount).n_13_hj2 = -1 Then
                            booVollstaendig = False
                            GoTo Exit_NotenVollstaendig
                        End If
            End Select
    End Select
    
Exit_NotenVollstaendig:
    NotenVollstaendig = booVollstaendig
    
End Function
Private Function FpaBestanden(lngArt As Long, mycolFaecherhalbjahr As colFaecherHalbjahr) As Boolean
'ist die fachpraktische Ausbildung bestanden?
Dim booRueckgabe As Boolean
Dim intCount As Integer

    Select Case lngArt
        Case 1  ' Probezeit
            booRueckgabe = True
            For intCount = 1 To mycolFaecherhalbjahr.Count
                If mycolFaecherhalbjahr.Item(intCount).fach_uid = 35 Then    'fachpraktische Ausbildung
                    If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 < 4 Then
                        booRueckgabe = False
                        GoTo Exit_FpaBestanden
                    End If
                End If
            Next intCount
        Case 2  ' Schuljahr
            booRueckgabe = True
            For intCount = 1 To mycolFaecherhalbjahr.Count
                If mycolFaecherhalbjahr.Item(intCount).fach_uid = 35 Then    'fachpraktische Ausbildung
                    If (Not mycolFaecherhalbjahr.Item(intCount).n_11_hj1 > 4 And Not mycolFaecherhalbjahr.Item(intCount).n_11_hj2 > 4) Or _
                        ((CDbl(mycolFaecherhalbjahr.Item(intCount).n_11_hj1) + CDbl(mycolFaecherhalbjahr.Item(intCount).n_11_hj2)) < 10) Then
                        booRueckgabe = False
                        GoTo Exit_FpaBestanden
                    End If
                End If
            Next intCount
    End Select
    
Exit_FpaBestanden:
    FpaBestanden = booRueckgabe
    
End Function
Private Function AlleJahresnotenGroesserDrei(lngArt As Long, lngJahrgangsstufe As Long, mycolFaecherhalbjahr As colFaecherHalbjahr) As Boolean
' alle Leistungen >= 4?
Dim booRueckgabe As Boolean

Dim intCount As Integer
Dim intAnzahlNull As Integer
Dim intAnzahlKleinerVier As Integer
Dim intObergrenze1 As Integer
Dim intObergrenze2 As Integer
Dim dblsumme As Double
Dim dblJahresnote As Double
Dim intSummeJahresnoten As Integer

    On Error GoTo Err_JahresnotenPruefen
    
    intObergrenze1 = 5 * CDbl(mycolFaecherhalbjahr.Count - 1)
    intObergrenze2 = 6 * CDbl(mycolFaecherhalbjahr.Count - 1)
    
    Select Case lngArt
        Case 1  'Probezeit
            Select Case lngJahrgangsstufe
                Case 1  'Vorklasse
                    intAnzahlNull = 0
                    intAnzahlKleinerVier = 0
                    intSummeJahresnoten = 0
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1 = 0 Then
                            intAnzahlNull = intAnzahlNull + 1
                        Else
                            If mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1 < 4 Then
                                intAnzahlKleinerVier = intAnzahlKleinerVier + 1
                            End If
                        End If
                        intSummeJahresnoten = intSummeJahresnoten + CInt(mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1)
                    Next intCount
                Case 2  '11. Klasse
                    intAnzahlNull = 0
                    intAnzahlKleinerVier = 0
                    intSummeJahresnoten = 0
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).fach_uid <> 35 Then
                            If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 = 0 Then
                                intAnzahlNull = intAnzahlNull + 1
                            Else
                                If mycolFaecherhalbjahr.Item(intCount).n_11_hj1 < 4 Then
                                    intAnzahlKleinerVier = intAnzahlKleinerVier + 1
                                End If
                            End If
                            intSummeJahresnoten = intSummeJahresnoten + CInt(mycolFaecherhalbjahr.Item(intCount).n_11_hj1)
                        End If
                    Next intCount
                Case 3  '12. Klasse
                Case 4  '13. Klasse
            End Select
        Case 2  'Schuljahr
            Select Case lngJahrgangsstufe
                Case 1  'Vorklasse
                    intAnzahlNull = 0
                    intAnzahlKleinerVier = 0
                    intSummeJahresnoten = 0
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        If mycolFaecherhalbjahr.Item(intCount).fach_uid <> 35 Then
                            dblsumme = CDbl(mycolFaecherhalbjahr.Item(intCount).n_vkl_hj1) + CDbl(mycolFaecherhalbjahr.Item(intCount).n_vkl_hj2)
                            dblJahresnote = Runden(dblsumme, 0)
                            If dblJahresnote = 0 Then
                                intAnzahlNull = intAnzahlNull + 1
                            Else
                                If dblJahresnote < 4 Then
                                    intAnzahlKleinerVier = intAnzahlKleinerVier + 1
                                End If
                            End If
                            intSummeJahresnoten = intSummeJahresnoten + dblJahresnote
                        End If
                    Next intCount
                Case 2  '11. Klasse
                    intAnzahlNull = 0
                    intAnzahlKleinerVier = 0
                    intSummeJahresnoten = 0
                    For intCount = 1 To mycolFaecherhalbjahr.Count
                        dblsumme = CDbl(mycolFaecherhalbjahr.Item(intCount).n_11_hj1) + CDbl(mycolFaecherhalbjahr.Item(intCount).n_11_hj2)
                        dblJahresnote = Runden(dblsumme, 0)
                        If dblJahresnote = 0 Then
                            intAnzahlNull = intAnzahlNull + 1
                        Else
                            If dblJahresnote < 4 Then
                                intAnzahlKleinerVier = intAnzahlKleinerVier + 1
                            End If
                        End If
                        intSummeJahresnoten = intSummeJahresnoten + dblJahresnote
                    Next intCount
                Case 3  '12. Klasse
                Case 4  '13. Klasse
            End Select
    End Select
    
    If (intAnzahlNull = 0 And intAnzahlKleinerVier = 0) Then 'bestanden
        booRueckgabe = True
        GoTo Rueckgabe_JahresnotenPruefen
    Else
        If (intAnzahlNull = 0 And intAnzahlKleinerVier = 1 And intSummeJahresnoten >= intObergrenze1) Then
            booRueckgabe = True
            GoTo Rueckgabe_JahresnotenPruefen
        Else
            If (intAnzahlNull = 1 And intSummeJahresnoten >= intObergrenze2) Or (intAnzahlKleinerVier = 2 And intSummeJahresnoten >= intObergrenze2) Then
                booRueckgabe = True
                GoTo Rueckgabe_JahresnotenPruefen
            Else
                booRueckgabe = False
                GoTo Rueckgabe_JahresnotenPruefen
            End If
        End If
    End If
    
Rueckgabe_JahresnotenPruefen:
    AlleJahresnotenGroesserDrei = booRueckgabe
    
Exit_JahresnotenPruefen:
    Exit Function

Err_JahresnotenPruefen:
    FehlermeldungAusgeben "Jahresnoten überprüfen", Err.Source, Err.Number, Err.Description
    Resume Exit_JahresnotenPruefen

End Function


