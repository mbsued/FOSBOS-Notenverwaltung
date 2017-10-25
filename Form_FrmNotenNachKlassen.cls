VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FrmNotenNachKlassen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private myColSchuelerNoten As colSchülerNoten       'Sammlung der Schüler
Private myColSchuelerHalbjahr As colSchülerHalbjahr 'Sammlung der Halbjahresleistungen der Schüler
Private myclsKlasseNoten As clsKlasseNoten          'Noteneinstellungen für komplette Klasse
Private booFormularAenderung As Boolean             'Wurde eine Änderung durchgeführt?
Private booHalbjahrAenderung As Boolean             'Wurde die Halbjahresleistung geändert?
Private booRueckgabeAenderungUebernehmen As Boolean 'Sind die Änderungen beim Schüler zu übernehmen?
Private booBeenden As Boolean

Private strRueckgabeSchuelerNotenSchulaufgabenHj1(0 To 1) As String   'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Schulaufgaben 1. Halbjahr Datum)
Private strRueckgabeSchuelerNotenSchulaufgabenHj2(0 To 1) As String   'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Schulaufgaben 2. Halbjahr Datum)
Private strRueckgabeSchuelerNotenSonstigeDatHj1(0 To 6) As String        'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Sonstige Leistungen 1. Halbjahr Datum)
Private strRueckgabeSchuelerNotenSonstigeDatHj2(0 To 6) As String        'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Sonstige Leistungen 2. Halbjahr Datum)
Private strRueckgabeSchuelerNotenSonstigeGewHj1(0 To 6) As String        'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Sonstige Leistungen 1. Halbjahr Gewichtung)
Private strRueckgabeSchuelerNotenSonstigeGewHj2(0 To 6) As String        'Rückgabe der Einstellungen aus dem Formular EinstellungenNotenSchueler (Sonstige Leistungen 2. Halbjahr Gewichtung)
Private intRueckgabeKlasseUidArt As Integer         ' Rückgabe der Einstellungen aus dem Formular EinstellungenNotenKlassen (Art der Leistungserhebung)
Private strRueckgabeKlasseDatum As String           ' Rückgabe der Einstellungen aus dem Formular EinstellungenNotenKlassen (Leistungsdatum)
Private strRueckgabeKlasseGewichtung As String      ' Rückgabe der Einstellungen aus dem Formular EinstellungenNotenKlassen (Gewichtung)

Private Sub cmd_schueler_1_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 1
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 1
    End If

End Sub

Private Sub cmd_schueler_2_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 2
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 2
    End If

End Sub

Private Sub cmd_schueler_3_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 3
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 3
    End If

End Sub

Private Sub cmd_schueler_4_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 4
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 4
    End If

End Sub

Private Sub cmd_schueler_5_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 5
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 5
    End If

End Sub

Private Sub cmd_schueler_6_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 6
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 6
    End If

End Sub

Private Sub cmd_schueler_7_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 7
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 7
    End If

End Sub

Private Sub cmd_schueler_8_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 8
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 8
    End If

End Sub

Private Sub cmd_schueler_9_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 9
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 9
    End If

End Sub

Private Sub cmd_schueler_10_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 10
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 10
    End If

End Sub

Private Sub cmd_schueler_11_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 11
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 11
    End If

End Sub

Private Sub cmd_schueler_12_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 12
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 12
    End If

End Sub

Private Sub cmd_schueler_13_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 13
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 13
    End If

End Sub

Private Sub cmd_schueler_14_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 14
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 14
    End If

End Sub

Private Sub cmd_schueler_15_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 15
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 15
    End If

End Sub

Private Sub cmd_schueler_16_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 16
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 16
    End If

End Sub

Private Sub cmd_schueler_17_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 17
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 17
    End If

End Sub

Private Sub cmd_schueler_18_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 18
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 18
    End If

End Sub

Private Sub cmd_schueler_19_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 19
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 19
    End If

End Sub

Private Sub cmd_schueler_20_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 20
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 20
    End If

End Sub

Private Sub cmd_schueler_21_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 21
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 21
    End If

End Sub

Private Sub cmd_schueler_22_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 22
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 22
    End If

End Sub

Private Sub cmd_schueler_23_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 23
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 23
    End If

End Sub

Private Sub cmd_schueler_24_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 24
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 24
    End If

End Sub

Private Sub cmd_schueler_25_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 25
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 25
    End If

End Sub

Private Sub cmd_schueler_26_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 26
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 26
    End If

End Sub

Private Sub cmd_schueler_27_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 27
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 27
    End If

End Sub

Private Sub cmd_schueler_28_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 28
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 28
    End If

End Sub

Private Sub cmd_schueler_29_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 29
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 29
    End If

End Sub

Private Sub cmd_schueler_30_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 30
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 30
    End If

End Sub

Private Sub cmd_schueler_31_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 31
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 31
    End If

End Sub

Private Sub cmd_schueler_32_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 32
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 32
    End If

End Sub

Private Sub cmd_schueler_33_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 33
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 33
    End If

End Sub

Private Sub cmd_schueler_34_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 34
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 34
    End If

End Sub

Private Sub cmd_schueler_35_Click()
'PopUp - Formular für Einstellungen öffnen
'Werte übergeben
'sind Einstellungen zu übernehmen, dann machs auch

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSchuelerOeffnen 35
                                        
    If booRueckgabeAenderungUebernehmen Then
        IndividuelleDatenSchreiben 35
    End If

End Sub

Private Sub cmd_son1_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
' sind die Einstellungen zu übernehmen, dann machs auch
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 1, 1, Nz(myclsKlasseNoten.uid_art_son1_hj1, 0), Nz(myclsKlasseNoten.gew_son1_hj1, ""), Nz(myclsKlasseNoten.dat_son1_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son1_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son1_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son1_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son1_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son1_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son1_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son1_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son2_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 2, 1, Nz(myclsKlasseNoten.uid_art_son2_hj1, 0), Nz(myclsKlasseNoten.gew_son2_hj1, ""), Nz(myclsKlasseNoten.dat_son2_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son2_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son2_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son2_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son2_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son2_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son2_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son2_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son3_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 3, 1, Nz(myclsKlasseNoten.uid_art_son3_hj1, 0), Nz(myclsKlasseNoten.gew_son3_hj1, ""), Nz(myclsKlasseNoten.dat_son3_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son3_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son3_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son3_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son3_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son3_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son3_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son3_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son4_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 4, 1, Nz(myclsKlasseNoten.uid_art_son4_hj1, 0), Nz(myclsKlasseNoten.gew_son4_hj1, ""), Nz(myclsKlasseNoten.dat_son4_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son4_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son4_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son4_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son4_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son4_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son4_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son4_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son5_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 5, 1, Nz(myclsKlasseNoten.uid_art_son5_hj1, 0), Nz(myclsKlasseNoten.gew_son5_hj1, ""), Nz(myclsKlasseNoten.dat_son5_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son5_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son5_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son5_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son5_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son5_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son5_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son5_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son6_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 6, 1, Nz(myclsKlasseNoten.uid_art_son6_hj1, 0), Nz(myclsKlasseNoten.gew_son6_hj1, ""), Nz(myclsKlasseNoten.dat_son6_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son6_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son6_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son6_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son6_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son6_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son6_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son6_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son7_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 7, 1, Nz(myclsKlasseNoten.uid_art_son7_hj1, 0), Nz(myclsKlasseNoten.gew_son7_hj1, ""), Nz(myclsKlasseNoten.dat_son7_hj1, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son7_hj1 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son7_hj1 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son7_hj1 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son7_hj1 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son7_hj1 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son7_hj1 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son7_hj1 = strRueckgabeKlasseDatum
        End If
        booFormularAenderung = True

    End If
    
End Sub

Private Sub cmd_son1_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 1, 2, Nz(myclsKlasseNoten.uid_art_son1_hj2, 0), Nz(myclsKlasseNoten.gew_son1_hj2, ""), Nz(myclsKlasseNoten.dat_son1_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son1_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son1_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son1_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son1_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son1_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son1_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son1_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_son2_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 2, 2, Nz(myclsKlasseNoten.uid_art_son2_hj2, 0), Nz(myclsKlasseNoten.gew_son2_hj2, ""), Nz(myclsKlasseNoten.dat_son2_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son2_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son2_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son2_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son2_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son2_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son2_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son2_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
    
End Sub

Private Sub cmd_son3_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 3, 2, Nz(myclsKlasseNoten.uid_art_son3_hj2, 0), Nz(myclsKlasseNoten.gew_son3_hj2, ""), Nz(myclsKlasseNoten.dat_son3_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son3_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son3_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son3_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son3_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son3_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son3_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son3_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_son4_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 4, 2, Nz(myclsKlasseNoten.uid_art_son4_hj2, 0), Nz(myclsKlasseNoten.gew_son4_hj2, ""), Nz(myclsKlasseNoten.dat_son4_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son4_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son4_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son4_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son4_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son4_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son4_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son4_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_son5_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 5, 2, Nz(myclsKlasseNoten.uid_art_son5_hj2, 0), Nz(myclsKlasseNoten.gew_son5_hj2, ""), Nz(myclsKlasseNoten.dat_son5_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son5_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son5_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son5_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son5_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son5_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son5_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son5_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_son6_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 6, 2, Nz(myclsKlasseNoten.uid_art_son6_hj2, 0), Nz(myclsKlasseNoten.gew_son6_hj2, ""), Nz(myclsKlasseNoten.dat_son6_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son6_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son6_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son6_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son6_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son6_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son6_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son6_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_son7_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenOeffnen 7, 2, Nz(myclsKlasseNoten.uid_art_son7_hj2, 0), Nz(myclsKlasseNoten.gew_son7_hj2, ""), Nz(myclsKlasseNoten.dat_son7_hj2, ""), Nz(myclsKlasseNoten.fach_uid, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                myColSchuelerNoten.Item(intCount).uid_art_son7_hj2 = intRueckgabeKlasseUidArt
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_son7_hj2 = strRueckgabeKlasseDatum
                End If
                myColSchuelerNoten.Item(intCount).gew_son7_hj2 = strRueckgabeKlasseGewichtung
            Else
                myColSchuelerNoten.Item(intCount).uid_art_son7_hj2 = intRueckgabeKlasseUidArt
            End If
        Next intCount
        
        myclsKlasseNoten.uid_art_son7_hj2 = intRueckgabeKlasseUidArt
        myclsKlasseNoten.gew_son7_hj2 = strRueckgabeKlasseGewichtung
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_son7_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub
Private Sub cmd_sa1_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
' sind die Einstellungen zu übernehmen, dann machs auch
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSAOeffnen 1, 1, Nz(myclsKlasseNoten.dat_sa1_hj1, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_sa1_hj1 = strRueckgabeKlasseDatum
                End If
            End If
        Next intCount
        
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_sa1_hj1 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub


Private Sub cmd_sa2_1_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSAOeffnen 2, 1, Nz(myclsKlasseNoten.dat_sa2_hj1, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_sa2_hj1 = strRueckgabeKlasseDatum
                End If
            End If
        Next intCount
        
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_sa2_hj1 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_sa1_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSAOeffnen 1, 2, Nz(myclsKlasseNoten.dat_sa1_hj2, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_sa1_hj2 = strRueckgabeKlasseDatum
                End If
            End If
        Next intCount
        
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_sa1_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_sa2_2_Click()
' PopUp-Formular für Einstellungen öffnen
' Spaltenzahl und Halbjahr übergeben
Dim intCount As Integer

    booRueckgabeAenderungUebernehmen = False
    FormEinstellungenSAOeffnen 2, 2, Nz(myclsKlasseNoten.dat_sa2_hj2, "")
    If booRueckgabeAenderungUebernehmen Then
        For intCount = 1 To myColSchuelerNoten.Count
            If myColSchuelerNoten.Item(intCount).ind_einstellung = False Then
                If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
                    myColSchuelerNoten.Item(intCount).dat_sa2_hj2 = strRueckgabeKlasseDatum
                End If
            End If
        Next intCount
        
        If Not IsNull(strRueckgabeKlasseDatum) And Not strRueckgabeKlasseDatum = "" Then
            myclsKlasseNoten.dat_sa2_hj2 = strRueckgabeKlasseDatum
        End If
        
        booFormularAenderung = True
        
    End If
    
End Sub

Private Sub cmd_sone3_1_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 3, 1
End Sub

Private Sub cmd_sone4_1_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 4, 1
End Sub

Private Sub cmd_sone5_1_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 5, 1
End Sub

Private Sub cmd_sone6_1_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 6, 1
End Sub

Private Sub cmd_sone7_1_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 7, 1
End Sub

Private Sub cmd_sone3_2_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 3, 2
End Sub

Private Sub cmd_sone4_2_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 4, 2
End Sub

Private Sub cmd_sone5_2_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 5, 2
End Sub

Private Sub cmd_sone6_2_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 6, 2
End Sub

Private Sub cmd_sone7_2_Click()
' eine weitere Spalte einblenden
' diesen Button ausblenden, dafür den nächsten einblenden
' Eingabefelder sichtbar machen
' Änderungsbutton einblenden
' Anzahl der Spalten für ClsSchülerNoten erhöhen
' Anzahl der Spalten für ColSchülerNoten erhöhen
' Formular PUEinstellungenNotenKlasse öffnen
    SpalteHinzufuegen 7, 2
End Sub

Private Sub cmdAbbrechen_Click()
' Formular schließen vorher nachfragen ob Änderungen verworfen werden sollen

    On Error GoTo Err_cmdclose_Click
    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen verworfen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            booBeenden = True
            DOCMD.Close
        End If
    Else
        If booHalbjahrAenderung Then
            If MsgBox("Sollen die Änderungen an den Halbjahresleistungen verworfen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
                booBeenden = True
                DOCMD.Close
            End If
        Else
            booBeenden = True
            DOCMD.Close
        End If
    End If

Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click
End Sub

Private Sub cmdBerechnen_Click()
' Berechne die Noten der Schüler
Dim dblDurchschnittSonstigeLeistungen As Double     'Durchschnitt der sonstigen Leistungen
Dim dblSummeSchulaufgaben As Double                 'Summe der Schulaufgaben
Dim dblErgebnis As Double                           'Ergebnis

Dim intCountSchueler As Integer     'Zähler für Schüler
Dim intCountHalbjahr As Integer     'Zähler für Halbjahr
Dim strControl1 As String           'Steuerelement Durchschnitt Sonstige
Dim strControl2 As String           'Steuerelement Halbjahresleistung berechnet
Dim StrControl3 As String           'Steuerelement Halbjahresleistung überschreibbar

    On Error GoTo Err_cmdBerechnen_Click

    For intCountSchueler = 1 To myColSchuelerNoten.Count
        For intCountHalbjahr = 1 To 2
            strControl1 = "txt_mdg_" & intCountHalbjahr & "_" & intCountSchueler
            strControl2 = "txt_hjg_" & intCountHalbjahr & "_" & intCountSchueler
            StrControl3 = "txt_hjgz_" & intCountHalbjahr & "_" & intCountSchueler
            Select Case intCountHalbjahr
                Case 1
                    If myColSchuelerNoten.Item(intCountSchueler).fach_uid = 35 Then 'fachpraktische Ausbildung
                        dblDurchschnittSonstigeLeistungen = DurchschnittFpaLeistungBerechnen(Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son1_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son1_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son2_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son2_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son3_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son3_hj1)
                    Else
                        dblDurchschnittSonstigeLeistungen = DurchschnittSonstigeLeistungBerechnen(Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son1_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son1_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son2_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son2_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son3_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son3_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son4_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son4_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son5_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son5_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son6_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son6_hj1, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son7_hj1, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son7_hj1)
                    End If
                    If dblDurchschnittSonstigeLeistungen > -1 Then
                        Me.Controls(strControl1).Value = dblDurchschnittSonstigeLeistungen
                    Else
                        Me.Controls(strControl1).Value = "-"
                    End If
                    
                    dblSummeSchulaufgaben = SchulaufgabenBerechnen(myColSchuelerNoten.Item(intCountSchueler).n_sa1_hj1, myColSchuelerNoten.Item(intCountSchueler).n_sa2_hj1, myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj1)
                    
                    If myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj1 > 0 Then
                        If dblSummeSchulaufgaben > -1 And dblDurchschnittSonstigeLeistungen > -1 Then
                            If myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj1 = 1 Then
                                dblErgebnis = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2
                                If dblErgebnis < 1 Then
                                    Me.Controls(strControl2).Value = 0
                                    Me.Controls(StrControl3).Value = 0
                                Else
                                    Me.Controls(strControl2).Value = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2
                                    Me.Controls(StrControl3).Value = Runden((dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2, 0)
                                End If
                            Else
                                dblErgebnis = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3
                                If dblErgebnis < 1 Then
                                    Me.Controls(strControl2).Value = 0
                                    Me.Controls(StrControl3).Value = 0
                                Else
                                    Me.Controls(strControl2).Value = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3
                                    Me.Controls(StrControl3).Value = Runden((dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3, 0)
                                End If
                            End If
                        Else
                            If dblDurchschnittSonstigeLeistungen > -1 And _
                                (myColSchuelerNoten.Item(intCountSchueler).uid_art_son1_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son2_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son3_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son4_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son5_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son6_hj1 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son7_hj1 = 4) Then    'keine Schulaufgaben aber Sonstige Leistungen da und Ersatzprüfung
                                Me.Controls(strControl2).Value = dblDurchschnittSonstigeLeistungen
                                Me.Controls(StrControl3).Value = Runden(dblDurchschnittSonstigeLeistungen, 0)
                            Else
                                Me.Controls(strControl2).Value = "-"
                                Me.Controls(StrControl3).Value = "-"
                            End If
                        End If
                    Else
                        If dblDurchschnittSonstigeLeistungen > -1 Then
                            If dblDurchschnittSonstigeLeistungen < 1 Then
                                Me.Controls(strControl2).Value = 0
                                Me.Controls(StrControl3).Value = 0
                            Else
                                Me.Controls(strControl2).Value = dblDurchschnittSonstigeLeistungen
                                Me.Controls(StrControl3).Value = Runden(dblDurchschnittSonstigeLeistungen, 0)
                            End If
                        Else
                            Me.Controls(strControl2).Value = "-"
                            Me.Controls(StrControl3).Value = "-"
                        End If
                    End If
                    Select Case myColSchuelerHalbjahr.Item(intCountSchueler).jahrgangsstufe_uid
                        Case 1, 5, 6 'VKL und Vorkurs
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_vkl_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 2  '11. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_11_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 3  '12. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_12_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 4  '13. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_13_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                    End Select
                Case 2
                    If myColSchuelerNoten.Item(intCountSchueler).fach_uid = 35 Then 'fachpraktische Ausbildung
                        dblDurchschnittSonstigeLeistungen = DurchschnittFpaLeistungBerechnen(Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son1_hj2, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son1_hj2, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son2_hj2, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son2_hj2, _
                                                                                                Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son3_hj2, 0), _
                                                                                                myColSchuelerNoten.Item(intCountSchueler).n_son3_hj2)
                        Else

                            dblDurchschnittSonstigeLeistungen = DurchschnittSonstigeLeistungBerechnen(Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son1_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son1_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son2_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son2_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son3_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son3_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son4_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son4_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son5_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son5_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son6_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son6_hj2, _
                                                                                                    Nz(myColSchuelerNoten.Item(intCountSchueler).gew_son7_hj2, 0), _
                                                                                                    myColSchuelerNoten.Item(intCountSchueler).n_son7_hj2)
                    End If
                    If dblDurchschnittSonstigeLeistungen > -1 Then
                        Me.Controls(strControl1).Value = dblDurchschnittSonstigeLeistungen
                    Else
                        Me.Controls(strControl1).Value = "-"
                    End If
                    
                    dblSummeSchulaufgaben = SchulaufgabenBerechnen(myColSchuelerNoten.Item(intCountSchueler).n_sa1_hj2, myColSchuelerNoten.Item(intCountSchueler).n_sa2_hj2, myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj2)
                    
                    If myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj2 > 0 Then
                        If dblSummeSchulaufgaben > -1 And dblDurchschnittSonstigeLeistungen > -1 Then
                            If myColSchuelerNoten.Item(intCountSchueler).anz_sa_hj2 = 1 Then
                                dblErgebnis = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2
                                If dblErgebnis < 1 Then
                                    Me.Controls(strControl2).Value = 0
                                    Me.Controls(StrControl3).Value = 0
                                Else
                                    Me.Controls(strControl2).Value = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2
                                    Me.Controls(StrControl3).Value = Runden((dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 2, 0)
                                End If
                            Else
                                dblErgebnis = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3
                                If dblErgebnis < 1 Then
                                    Me.Controls(strControl2).Value = 0
                                    Me.Controls(StrControl3).Value = 0
                                Else
                                    Me.Controls(strControl2).Value = (dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3
                                    Me.Controls(StrControl3).Value = Runden((dblSummeSchulaufgaben + dblDurchschnittSonstigeLeistungen) / 3, 0)
                                End If
                            End If
                        Else
                            If dblDurchschnittSonstigeLeistungen > -1 And _
                                (myColSchuelerNoten.Item(intCountSchueler).uid_art_son1_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son2_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son3_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son4_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son5_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son6_hj2 = 4 Or _
                                myColSchuelerNoten.Item(intCountSchueler).uid_art_son7_hj2 = 4) Then    'keine Schulaufgaben aber Sonstige Leistungen da und Ersatzprüfung
                                Me.Controls(strControl2).Value = dblDurchschnittSonstigeLeistungen
                                Me.Controls(StrControl3).Value = Runden(dblDurchschnittSonstigeLeistungen, 0)
                            Else
                                Me.Controls(strControl2).Value = "-"
                                Me.Controls(StrControl3).Value = "-"
                            End If
                        End If
                    Else
                        If dblDurchschnittSonstigeLeistungen > -1 Then
                            If dblDurchschnittSonstigeLeistungen < 1 Then
                                Me.Controls(strControl2).Value = 0
                                Me.Controls(StrControl3).Value = 0
                            Else
                                Me.Controls(strControl2).Value = dblDurchschnittSonstigeLeistungen
                                Me.Controls(StrControl3).Value = Runden(dblDurchschnittSonstigeLeistungen, 0)
                            End If
                        Else
                            Me.Controls(strControl2).Value = "-"
                            Me.Controls(StrControl3).Value = "-"
                        End If
                    End If
                    Select Case myColSchuelerHalbjahr.Item(intCountSchueler).jahrgangsstufe_uid
                        Case 1, 5, 6 'VKL und Vorkurs
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_vkl_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 2  '11. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_11_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 3  '12. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_12_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                        Case 4  '13. Jahrgangsstufe
                            myColSchuelerHalbjahr.Item(intCountSchueler).n_13_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(strControl2).Value)
                    End Select
            End Select
        Next intCountHalbjahr
    Next intCountSchueler
    booHalbjahrAenderung = True
    
Exit_cmdBerechnen_Click:
    Exit Sub
    
Err_cmdBerechnen_Click:

    FehlermeldungAusgeben "Formular Halbjahr", Err.Source, Err.Number, Err.Description
    Resume Exit_cmdBerechnen_Click

End Sub

Private Sub cmdClose_Click()
'Schließen vorher nachfragen ob gespeichert werden soll
    On Error GoTo Err_cmdclose_Click
    
    If booFormularAenderung Then
        If MsgBox("Sollen die Änderungen an den Einzelnoten gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If
    If booHalbjahrAenderung Then
        If MsgBox("Sollen die Änderungen an den Halbjahresleistungen gespeichert werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdUebernehmen_Click
        End If
    End If
    
    booBeenden = True
    DOCMD.Close
    
Exit_cmdclose_Click:
    Exit Sub

Err_cmdclose_Click:
    MsgBox Err.Description
    Resume Exit_cmdclose_Click

End Sub

Private Sub cmdSpeichern_Click()
' die vorgegebenen Daten speichern (Schülerdaten und Noteneinstellungen der Klasse)
    SchuelerNotenSchreiben myColSchuelerNoten
    NotenEinstellungenSchreiben myclsKlasseNoten
    ' Formular geändert zurücksetzen
    booFormularAenderung = False

End Sub

Private Sub cmdUebernehmen_Click()
'Übernahme der Halbjahresleistungen
    SchuelerHalbjahresNotenSchreiben myColSchuelerHalbjahr
    booHalbjahrAenderung = False
End Sub

Private Sub Form_Load()
'Höhe und Breite einstellen
'Formular Änderung auf falsch stellen
'1. Klassengruppe und 1.Fach einstellen

    booBeenden = False
    Me.InsideHeight = Me.Section(acDetail).Height
    Me.InsideWidth = Me.Width
            
    booFormularAenderung = False
    booHalbjahrAenderung = False

    lstFach_AfterUpdate
    
End Sub

Private Sub Form_Close()
    If IstFormularGeoeffnet("frmStart") Then
        Forms.Item("frmStart").Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If booBeenden = False Then
        MsgBox "Schließen nur über die Schaltflächen Formular schließen oder Abbrechen möglich", vbOKOnly, "Notenprogramm FOSBOS"
        Cancel = True
    End If

End Sub

Private Sub lstFach_AfterUpdate()

lstfach_Start:

' jetzt die entsprechenden Schülerdaten lesen
    Set myColSchuelerNoten = SchuelerNotenLesen(Me.lstFach, Me.lstKlasse)
' keine Schülerdaten gefunden
    If myColSchuelerNoten Is Nothing Then
        MsgBox "In dieser Klassengruppe belegt kein Schüler das Fach " & lstFach.Column(1) & "." & vbCrLf & "Es wird das erste Fach in der Liste ausgewählt", vbOKOnly, "FOSBOS Notenverwaltung"
        lstFach.Value = 1
        GoTo lstfach_Start
    End If
' und die entsprechenden Einstellungen für Klasse/Fach
    Set myclsKlasseNoten = NotenEinstellungenKlasseFachLesen(Me.lstFach, Me.lstKlasse)
' und die Halbjahresdaten
    Set myColSchuelerHalbjahr = SchuelerHalbjahresNotenLesen(Me.lstFach, Me.lstKlasse)
' erst mal wieder alles einblenden anschließend ausblenden
    AlleZeilenUndNotenfelderEinblenden
    ZeilenAusblenden
    NotenfelderAusblenden
    AenderungsButtonsEinblenden
    AenderungsButtonsAusblenden
    WeitereSpalteButtonsEinblenden

' Daten füllen
    SchuelerNameSchreiben
    AlleNotenfelderLeeren
    NotenfelderFuellen
End Sub

' Prüfe ob Änderung war, wenn ja dann nachfragen ob gespeichert werden soll
Private Sub lstFach_BeforeUpdate(Cancel As Integer)
    
    If booFormularAenderung Then
        If MsgBox("Sie haben Eingaben geändert. Sollen sie übernommen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If
    
    If booHalbjahrAenderung Then
        If MsgBox("Sie haben Halbjahresleistungen geändert. Sollen sie übernommen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdUebernehmen_Click
        End If
    End If
    
    
End Sub

Private Sub lstKlasse_AfterUpdate()
    Me.lstFach.Requery
    lstFach = 1
    lstFach_AfterUpdate
End Sub

' Prüfe ob Änderung war, wenn ja dann nachfragen ob gespeichert werden soll
Private Sub lstKlasse_BeforeUpdate(Cancel As Integer)
    If booFormularAenderung Then
        If MsgBox("Sie haben Eingaben geändert. Sollen sie übernommen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdSpeichern_Click
        End If
    End If
    If booHalbjahrAenderung Then
        If MsgBox("Sie haben Halbjahresleistungen geändert. Sollen sie übernommen werden?", vbYesNo, "FOSBOS Notenverwaltung") = vbYes Then
            cmdUebernehmen_Click
        End If
    End If
End Sub

Private Sub ZeilenAusblenden()
'Alle Zeilen ausblenden die nicht benötigt werden
Dim intCount As Integer         'allgemeiner Zähler
Dim intAnzahlSa As Integer      'Schulaufgaben
Dim intAnzahlSon As Integer     'sonstige Leistungen
Dim intSchulhalbjahr As Integer 'Schulhalbjahr
Dim strControl1 As String

    For intCount = myColSchuelerNoten.Count + 1 To 35
        strControl1 = "cmd_schueler_" & intCount
        Me.Controls(strControl1).Visible = False
        For intAnzahlSa = 1 To 2
            For intSchulhalbjahr = 1 To 2
                strControl1 = "txt_sa" & intAnzahlSa & "_" & intSchulhalbjahr & "_" & intCount
                Me.Controls(strControl1).Visible = False
            Next intSchulhalbjahr
        Next intAnzahlSa
        For intAnzahlSon = 1 To 7
            For intSchulhalbjahr = 1 To 2
                strControl1 = "txt_md" & intAnzahlSon & "_" & intSchulhalbjahr & "_" & intCount
                Me.Controls(strControl1).Visible = False
            Next intSchulhalbjahr
        Next intAnzahlSon
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_mdg" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = False
        Next intSchulhalbjahr
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_hjg" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = False
        Next intSchulhalbjahr
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_hjgz" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = False
        Next intSchulhalbjahr
    Next intCount

End Sub
Private Sub SchuelerNameSchreiben()
'Schreibe den Schülernamen auf die Schaltfläche
Dim intCount As Integer
Dim StrControl As String

    For intCount = 1 To myColSchuelerNoten.Count
        StrControl = "cmd_schueler_" & intCount
        Me.Controls(StrControl).Visible = True
        Me.Controls(StrControl).Caption = myColSchuelerNoten.Item(intCount).nachname & " " & myColSchuelerNoten.Item(intCount).rufname
        If myColSchuelerNoten.Item(intCount).ind_einstellung Then
            Me.Controls(StrControl).Caption = "**  " & Me.Controls(StrControl).Caption
        End If
    Next intCount

End Sub
Private Sub AlleNotenfelderLeeren()
' leere alle Notenfelder
Dim StrControl As String    'welches Textfeld
Dim intCount As Integer      'Zähler
Dim intSaCount As Integer   'Schulaufgabenzähler
Dim intSonCount As Integer  'Zähler sonstige Leistungen

    For intCount = 1 To myColSchuelerNoten.Count
        For intSaCount = 1 To 2
            StrControl = "txt_sa" & intSaCount & "_1_" & intCount
            Me.Controls(StrControl).Value = vbNullString
        Next intSaCount
        For intSaCount = 1 To 2
            StrControl = "txt_sa" & intSaCount & "_2_" & intCount
            Me.Controls(StrControl).Value = vbNullString
        Next intSaCount
        For intSonCount = 1 To 7
            StrControl = "txt_md" & intSonCount & "_1_" & intCount
            Me.Controls(StrControl).Value = vbNullString
        Next intSonCount
        For intSonCount = 1 To 7
            StrControl = "txt_md" & intSonCount & "_2_" & intCount
            Me.Controls(StrControl).Value = vbNullString
        Next intSonCount
    Next intCount
    
    For intCount = 1 To myColSchuelerHalbjahr.Count
        StrControl = "txt_hjgz_1_" & intCount
        Me.Controls(StrControl).Value = vbNullString
        StrControl = "txt_hjgz_2_" & intCount
        Me.Controls(StrControl).Value = vbNullString
    Next intCount
    
End Sub
Private Sub NotenfelderFuellen()
'Schreibe die Noten der Schüler in die jeweiligen Felder
Dim lngCount As Long        'allgemeiner Zähler
Dim intSaCount As Integer   'Schulaufgabenzähler
Dim intSonCount As Integer  'Zähler sonstige Leistungen
Dim StrControl As String    'welches Textfeld

    For lngCount = 1 To myColSchuelerNoten.Count
        ' Schulaufgaben 1.Halbjahr
        For intSaCount = 1 To myColSchuelerNoten.Item(lngCount).anz_sa_hj1
            StrControl = "txt_sa" & intSaCount & "_1_" & lngCount
            Select Case intSaCount
                Case 1  '1.Schulaufgabe
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_sa1_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 2  '2.Schulaufgabe
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_sa2_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
            End Select
        Next intSaCount
        ' Schulaufgaben 2.Halbjahr
        For intSaCount = 1 To myColSchuelerNoten.Item(lngCount).anz_sa_hj2
            StrControl = "txt_sa" & intSaCount & "_2_" & lngCount
            Select Case intSaCount
                Case 1  '1.Schulaufgabe
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_sa1_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 2  '2.Schulaufgabe
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_sa2_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
            End Select
        Next intSaCount
        ' sonstige Leistungen 1.Halbjahr
        For intSonCount = 1 To myColSchuelerNoten.Item(lngCount).anz_son_hj1
            StrControl = "txt_md" & intSonCount & "_1_" & lngCount
            Select Case intSonCount
                Case 1  '1.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son1_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 2  '2.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son2_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 3  '3.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son3_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 4  '4.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son4_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 5  '5.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son5_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 6  '6.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son6_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 7  '7.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son7_hj1
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 1, lngCount, Me.Controls(StrControl).Value
                    End If
            End Select
        Next intSonCount
        ' sonstige Leistungen 2.Halbjahr
        For intSonCount = 1 To myColSchuelerNoten.Item(lngCount).anz_son_hj2
            StrControl = "txt_md" & intSonCount & "_2_" & lngCount
            Select Case intSonCount
                Case 1  '1.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son1_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 2  '2.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son2_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 3  '3.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son3_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 4  '4.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son4_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 5  '5.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son5_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 6  '6.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son6_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
                Case 7  '7.Leistung
                    Me.Controls(StrControl).Value = myColSchuelerNoten.Item(lngCount).n_son7_hj2
                    If Not IsNull(Me.Controls(StrControl).Value) Then
                        HalbjahrSperrenWennNoteDa 2, lngCount, Me.Controls(StrControl).Value
                    End If
            End Select
        Next intSonCount
    Next lngCount
    For lngCount = 1 To myColSchuelerHalbjahr.Count
            StrControl = "txt_hjgz_1_" & lngCount
            Select Case myColSchuelerHalbjahr.Item(lngCount).jahrgangsstufe_uid
                Case 1, 5, 6, 7 'Vorklassen und Vorkurs
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_vkl_hj1)
                Case 2  '11. Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_11_hj1)
                Case 3  '12.Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_12_hj1)
                Case 4  '13.Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_13_hj1)
            End Select
            StrControl = "txt_hjgz_2_" & lngCount
            Select Case myColSchuelerHalbjahr.Item(lngCount).jahrgangsstufe_uid
                Case 1, 5, 6, 7 'Vorklassen und Vorkurs
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_vkl_hj2)
                Case 2  '11. Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_11_hj2)
                Case 3  '12.Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_12_hj2)
                Case 4  '13.Klasse
                    Me.Controls(StrControl).Value = HalbjahresleistungUebertragen(1, myColSchuelerHalbjahr.Item(lngCount).n_13_hj2)
            End Select
    Next lngCount
End Sub
Private Sub HalbjahrSperrenWennNoteDa(lngHalbjahr As Long, lngSchueler As Long, strValue As Variant)
'Sperre das Feld Halbjahr, wenn einen Einzelnote da ist

Dim StrControl As String

    If (strValue >= 0 And strValue < 16) Or strValue = "-" Then
        StrControl = "txt_hjgz_" & lngHalbjahr & "_" & lngSchueler
        Me.Controls(StrControl).Enabled = False
    End If

End Sub
Private Sub NotenfelderAusblenden()
'Blende Notenfelder der Schüler aus
Dim intCount As Integer
Dim intAnzahlSa As Integer  'Schulaufgaben
Dim intAnzahlSon As Integer 'sonsige Leistungen
Dim StrControl As String

    For intCount = 1 To myColSchuelerNoten.Count
        'Schulaufgaben 1.Halbjahr ausblenden
        If myColSchuelerNoten.Item(intCount).anz_sa_hj1 < 2 Then
            For intAnzahlSa = myColSchuelerNoten.Item(intCount).anz_sa_hj1 + 1 To 2
                StrControl = "txt_sa" & intAnzahlSa & "_1_" & intCount
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSa
        End If
        'Schulaufgaben 2.Halbjahr ausblenden
        If myColSchuelerNoten.Item(intCount).anz_sa_hj2 < 2 Then
            For intAnzahlSa = myColSchuelerNoten.Item(intCount).anz_sa_hj2 + 1 To 2
                StrControl = "txt_sa" & intAnzahlSa & "_2_" & intCount
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSa
        End If
        'Sonstige 1.Halbjahr ausblenden
        If myColSchuelerNoten.Item(intCount).anz_son_hj1 < 7 Then
            For intAnzahlSon = myColSchuelerNoten.Item(intCount).anz_son_hj1 + 1 To 7
                StrControl = "txt_md" & intAnzahlSon & "_1_" & intCount
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        End If
        'Sonstige 2.Halbjahr ausblenden
        If myColSchuelerNoten.Item(intCount).anz_son_hj2 < 7 Then
            For intAnzahlSon = myColSchuelerNoten.Item(intCount).anz_son_hj2 + 1 To 7
                StrControl = "txt_md" & intAnzahlSon & "_2_" & intCount
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        End If
    Next intCount

End Sub
Private Sub AlleZeilenUndNotenfelderEinblenden()
'Blende alles ein
Dim intCount As Integer     'allgemeiner Zähler
Dim intAnzahlSa As Integer  'Schulaufgaben
Dim intAnzahlSon As Integer 'sonsige Leistungen
Dim intSchulhalbjahr As Integer 'Schulhalbjahr
Dim strControl1 As String

    For intCount = 1 To 35
        strControl1 = "cmd_schueler_" & intCount
        Me.Controls(strControl1).Visible = True
        For intAnzahlSa = 1 To 2
            For intSchulhalbjahr = 1 To 2
                strControl1 = "txt_sa" & intAnzahlSa & "_" & intSchulhalbjahr & "_" & intCount
                Me.Controls(strControl1).Visible = True
            Next intSchulhalbjahr
        Next intAnzahlSa
        For intAnzahlSon = 1 To 7
            For intSchulhalbjahr = 1 To 2
                strControl1 = "txt_md" & intAnzahlSon & "_" & intSchulhalbjahr & "_" & intCount
                Me.Controls(strControl1).Visible = True
            Next intSchulhalbjahr
        Next intAnzahlSon
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_mdg" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = True
            Me.Controls(strControl1).Value = ""
        Next intSchulhalbjahr
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_hjg" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = True
            Me.Controls(strControl1).Value = ""
        Next intSchulhalbjahr
        For intSchulhalbjahr = 1 To 2
            strControl1 = "txt_hjgz" & "_" & intSchulhalbjahr & "_" & intCount
            Me.Controls(strControl1).Visible = True
            Me.Controls(strControl1).Value = ""
            Me.Controls(strControl1).BackColor = RGB(255, 255, 255)
            Me.Controls(strControl1).ForeColor = RGB(0, 0, 0)
            Me.Controls(strControl1).Enabled = True
        Next intSchulhalbjahr
    Next intCount
    
End Sub
Private Sub AenderungsButtonsAusblenden()
'Blendet Überflüssige Änderungsbuttons aus
Dim intAnzahlSa As Integer      'Schulaufgaben
Dim intAnzahlSon As Integer     'sonstige Leistungen
Dim StrControl As String

        'Schulaufgaben 1.Halbjahr ausblenden
        If myclsKlasseNoten.anz_sa_hj1 < 2 Then
            For intAnzahlSa = myclsKlasseNoten.anz_sa_hj1 + 1 To 2
                StrControl = "cmd_sa" & intAnzahlSa & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSa
        End If
        'Schulaufgaben 2.Halbjahr ausblenden
        If myclsKlasseNoten.anz_sa_hj2 < 2 Then
            For intAnzahlSa = myclsKlasseNoten.anz_sa_hj2 + 1 To 2
                StrControl = "cmd_sa" & intAnzahlSa & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSa
        End If
        'Sonstige 1.Halbjahr ausblenden
        If myclsKlasseNoten.anz_son_hj1 < 7 Then
            For intAnzahlSon = myclsKlasseNoten.anz_son_hj1 + 1 To 7
                StrControl = "cmd_son" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        End If
        'Sonstige 2.Halbjahr ausblenden
        If myclsKlasseNoten.anz_son_hj2 < 7 Then
            For intAnzahlSon = myclsKlasseNoten.anz_son_hj2 + 1 To 7
                StrControl = "cmd_son" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        End If
End Sub
Private Sub WeitereSpalteButtonsEinblenden()
' alle Buttons für weiter Spalte einblenden
Dim intAnzahlSon As Integer     'sonstige Leistungen
Dim StrControl As String
    'Sonstige 1.Halbjahr einblenden
    Select Case myclsKlasseNoten.anz_son_hj1
        Case 2
            cmd_sone3_1.Visible = True
            For intAnzahlSon = 4 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 3
            cmd_sone3_1.Visible = False
            cmd_sone4_1.Visible = True
            For intAnzahlSon = 5 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 4
            For intAnzahlSon = 3 To 4
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone5_1.Visible = True
            For intAnzahlSon = 6 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 5
            For intAnzahlSon = 3 To 5
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone6_1.Visible = True
            For intAnzahlSon = 7 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 6
            For intAnzahlSon = 3 To 6
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone7_1.Visible = True
        Case 7
            For intAnzahlSon = 3 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
    End Select
    'Sonstige 2.Halbjahr einblenden
    Select Case myclsKlasseNoten.anz_son_hj2
        Case 2
            cmd_sone3_2.Visible = True
            For intAnzahlSon = 4 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 3
            cmd_sone3_2.Visible = False
            cmd_sone4_2.Visible = True
            For intAnzahlSon = 5 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 4
            For intAnzahlSon = 3 To 4
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone5_2.Visible = True
            For intAnzahlSon = 6 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 5
            For intAnzahlSon = 3 To 5
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone6_2.Visible = True
            For intAnzahlSon = 7 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
        Case 6
            For intAnzahlSon = 3 To 6
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
            cmd_sone7_2.Visible = True
        Case 7
            For intAnzahlSon = 3 To 7
                StrControl = "cmd_sone" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = False
            Next intAnzahlSon
    End Select
    
End Sub

Private Sub AenderungsButtonsEinblenden()
' alle Buttons einblenden
Dim intAnzahlSa As Integer      'Schulaufgaben
Dim intAnzahlSon As Integer     'sonstige Leistungen
Dim StrControl As String

        'Schulaufgaben 1.Halbjahr einblenden
            For intAnzahlSa = 1 To 2
                StrControl = "cmd_sa" & intAnzahlSa & "_1"
                Me.Controls(StrControl).Visible = True
            Next intAnzahlSa
        'Schulaufgaben 2.Halbjahr einblenden
            For intAnzahlSa = 1 To 2
                StrControl = "cmd_sa" & intAnzahlSa & "_2"
                Me.Controls(StrControl).Visible = True
            Next intAnzahlSa
        'Sonstige 1.Halbjahr einblenden
            For intAnzahlSon = 1 To 7
                StrControl = "cmd_son" & intAnzahlSon & "_1"
                Me.Controls(StrControl).Visible = True
            Next intAnzahlSon
        'Sonstige 2.Halbjahr einblenden
            For intAnzahlSon = 1 To 7
                StrControl = "cmd_son" & intAnzahlSon & "_2"
                Me.Controls(StrControl).Visible = True
            Next intAnzahlSon
End Sub

'Formular für Einstellungen sonstige Leistungen öffnen
Private Sub FormEinstellungenOeffnen(strSpalte As String, strHalbjahr As String, lngArt As Long, strGewichtung As String, strDatum As String, lngFach As Long)
    DOCMD.OpenForm "frmPUEinstellungenNotenKlasse", acNormal, , , acFormEdit, acDialog, strSpalte & ";" & strHalbjahr & ";" & lngArt & ";" & strGewichtung & ";" & strDatum & ";" & lngFach
End Sub

'Formular für Einstellungen Schulaufgaben öffnen
Private Sub FormEinstellungenSAOeffnen(strSpalte As String, strHalbjahr As String, strDatum As String)
    DOCMD.OpenForm "frmPUEinstellungenSchulaufgabenNotenKlasse", acNormal, , , acFormEdit, acDialog, strSpalte & ";" & strHalbjahr & ";" & strDatum
End Sub

'Formular für Einstellungen Schüler öffnen
Private Sub FormEinstellungenSchuelerOeffnen(StrSchueler As String)
    DOCMD.OpenForm "frmPUEinstellungenNotenSchueler", acNormal, , , acFormEdit, acDialog, StrSchueler
End Sub

Private Sub txt_hjgz_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_22_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_22_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_hjgz_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckHalbjahrEingabeUndSchreibe(2, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md1_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 1

End Sub

Private Sub txt_md1_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 10

End Sub

Private Sub txt_md1_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 11

End Sub

Private Sub txt_md1_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 12

End Sub

Private Sub txt_md1_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 13

End Sub

Private Sub txt_md1_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 14

End Sub

Private Sub txt_md1_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 15

End Sub

Private Sub txt_md1_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 16

End Sub

Private Sub txt_md1_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 17

End Sub

Private Sub txt_md1_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 18

End Sub

Private Sub txt_md1_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 19

End Sub

Private Sub txt_md1_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 2

End Sub

Private Sub txt_md1_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 20

End Sub

Private Sub txt_md1_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 21

End Sub

Private Sub txt_md1_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 22

End Sub

Private Sub txt_md1_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 23

End Sub

Private Sub txt_md1_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 24

End Sub

Private Sub txt_md1_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 25

End Sub

Private Sub txt_md1_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 26

End Sub

Private Sub txt_md1_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 27

End Sub

Private Sub txt_md1_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 28

End Sub

Private Sub txt_md1_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 29

End Sub

Private Sub txt_md1_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 3

End Sub

Private Sub txt_md1_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 30

End Sub

Private Sub txt_md1_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 31

End Sub

Private Sub txt_md1_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 32

End Sub

Private Sub txt_md1_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 33

End Sub

Private Sub txt_md1_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 34

End Sub

Private Sub txt_md1_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 35

End Sub

Private Sub txt_md1_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 4

End Sub

Private Sub txt_md1_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 5

End Sub

Private Sub txt_md1_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 6

End Sub

Private Sub txt_md1_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 7

End Sub

Private Sub txt_md1_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 8

End Sub

Private Sub txt_md1_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 1, 9

End Sub

Private Sub txt_md1_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 1

End Sub

Private Sub txt_md1_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 10

End Sub

Private Sub txt_md1_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 11

End Sub

Private Sub txt_md1_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 12

End Sub

Private Sub txt_md1_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 13

End Sub

Private Sub txt_md1_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 14

End Sub

Private Sub txt_md1_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 15

End Sub

Private Sub txt_md1_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 16

End Sub

Private Sub txt_md1_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 17

End Sub

Private Sub txt_md1_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 18

End Sub

Private Sub txt_md1_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 19

End Sub

Private Sub txt_md1_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 2

End Sub

Private Sub txt_md1_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 20

End Sub

Private Sub txt_md1_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 21

End Sub

Private Sub txt_md1_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 22

End Sub

Private Sub txt_md1_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 23

End Sub

Private Sub txt_md1_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 24

End Sub

Private Sub txt_md1_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 25

End Sub

Private Sub txt_md1_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 26

End Sub

Private Sub txt_md1_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 27

End Sub

Private Sub txt_md1_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 28

End Sub

Private Sub txt_md1_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 29

End Sub

Private Sub txt_md1_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 3

End Sub

Private Sub txt_md1_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 30

End Sub

Private Sub txt_md1_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 31

End Sub

Private Sub txt_md1_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 32

End Sub

Private Sub txt_md1_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 33

End Sub

Private Sub txt_md1_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 34

End Sub

Private Sub txt_md1_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 35

End Sub

Private Sub txt_md1_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 4

End Sub

Private Sub txt_md1_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 5

End Sub

Private Sub txt_md1_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 6

End Sub

Private Sub txt_md1_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 7

End Sub

Private Sub txt_md1_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 8

End Sub

Private Sub txt_md1_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 1, 2, 9

End Sub

Private Sub txt_md2_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 1

End Sub

Private Sub txt_md2_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 10

End Sub

Private Sub txt_md2_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 11

End Sub

Private Sub txt_md2_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 12

End Sub

Private Sub txt_md2_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 13

End Sub

Private Sub txt_md2_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 14

End Sub

Private Sub txt_md2_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 15

End Sub

Private Sub txt_md2_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 16

End Sub

Private Sub txt_md2_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 17

End Sub

Private Sub txt_md2_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 18

End Sub

Private Sub txt_md2_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 19

End Sub

Private Sub txt_md2_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 2

End Sub

Private Sub txt_md2_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 20

End Sub

Private Sub txt_md2_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 21

End Sub

Private Sub txt_md2_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 22

End Sub

Private Sub txt_md2_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 23

End Sub

Private Sub txt_md2_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 24

End Sub

Private Sub txt_md2_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 25

End Sub

Private Sub txt_md2_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 26

End Sub

Private Sub txt_md2_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 27

End Sub

Private Sub txt_md2_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 28

End Sub

Private Sub txt_md2_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 29

End Sub

Private Sub txt_md2_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 3

End Sub

Private Sub txt_md2_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 30

End Sub

Private Sub txt_md2_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 31

End Sub

Private Sub txt_md2_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 32

End Sub

Private Sub txt_md2_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 33

End Sub

Private Sub txt_md2_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 34

End Sub

Private Sub txt_md2_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 35

End Sub

Private Sub txt_md2_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 4

End Sub

Private Sub txt_md2_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 5

End Sub

Private Sub txt_md2_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 6

End Sub

Private Sub txt_md2_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 7

End Sub

Private Sub txt_md2_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 8

End Sub

Private Sub txt_md2_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 1, 9

End Sub

Private Sub txt_md2_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 1

End Sub

Private Sub txt_md2_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 10

End Sub

Private Sub txt_md2_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 11

End Sub

Private Sub txt_md2_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 12

End Sub

Private Sub txt_md2_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 13

End Sub

Private Sub txt_md2_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 14

End Sub

Private Sub txt_md2_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 15

End Sub

Private Sub txt_md2_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 16

End Sub

Private Sub txt_md2_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 17

End Sub

Private Sub txt_md2_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 18

End Sub

Private Sub txt_md2_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 19

End Sub

Private Sub txt_md2_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 2

End Sub

Private Sub txt_md2_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 20

End Sub

Private Sub txt_md2_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 21

End Sub

Private Sub txt_md2_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 22

End Sub

Private Sub txt_md2_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 23

End Sub

Private Sub txt_md2_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 24

End Sub

Private Sub txt_md2_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 25

End Sub

Private Sub txt_md2_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 26

End Sub

Private Sub txt_md2_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 27

End Sub

Private Sub txt_md2_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 28

End Sub

Private Sub txt_md2_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 29

End Sub

Private Sub txt_md2_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 3

End Sub

Private Sub txt_md2_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 30

End Sub

Private Sub txt_md2_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 31

End Sub

Private Sub txt_md2_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 32

End Sub

Private Sub txt_md2_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 33

End Sub

Private Sub txt_md2_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 34

End Sub

Private Sub txt_md2_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 35

End Sub

Private Sub txt_md2_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 4

End Sub

Private Sub txt_md2_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 5

End Sub

Private Sub txt_md2_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 6

End Sub

Private Sub txt_md2_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 7

End Sub

Private Sub txt_md2_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 8

End Sub

Private Sub txt_md2_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 2, 2, 9

End Sub

Private Sub txt_md3_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 1

End Sub

Private Sub txt_md3_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 10

End Sub

Private Sub txt_md3_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 11

End Sub

Private Sub txt_md3_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 12

End Sub

Private Sub txt_md3_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 13

End Sub

Private Sub txt_md3_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 14

End Sub

Private Sub txt_md3_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 15

End Sub

Private Sub txt_md3_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 16

End Sub

Private Sub txt_md3_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 17

End Sub

Private Sub txt_md3_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 18

End Sub

Private Sub txt_md3_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 19

End Sub

Private Sub txt_md3_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 2

End Sub

Private Sub txt_md3_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 20

End Sub

Private Sub txt_md3_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 21

End Sub

Private Sub txt_md3_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 22

End Sub

Private Sub txt_md3_1_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 23

End Sub

Private Sub txt_md3_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 24

End Sub

Private Sub txt_md3_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 25

End Sub

Private Sub txt_md3_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 26

End Sub

Private Sub txt_md3_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 27

End Sub

Private Sub txt_md3_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 28

End Sub

Private Sub txt_md3_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 29

End Sub

Private Sub txt_md3_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 3

End Sub

Private Sub txt_md3_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 30

End Sub

Private Sub txt_md3_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 31

End Sub

Private Sub txt_md3_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 32

End Sub

Private Sub txt_md3_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 33

End Sub

Private Sub txt_md3_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 34

End Sub

Private Sub txt_md3_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 35

End Sub

Private Sub txt_md3_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 4

End Sub

Private Sub txt_md3_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 5

End Sub

Private Sub txt_md3_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 6

End Sub

Private Sub txt_md3_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 7

End Sub

Private Sub txt_md3_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 8

End Sub

Private Sub txt_md3_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 1, 9

End Sub

Private Sub txt_md3_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 1

End Sub

Private Sub txt_md3_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 10

End Sub

Private Sub txt_md3_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 11

End Sub

Private Sub txt_md3_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 12

End Sub

Private Sub txt_md3_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 13

End Sub

Private Sub txt_md3_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 14

End Sub

Private Sub txt_md3_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 15

End Sub

Private Sub txt_md3_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 16

End Sub

Private Sub txt_md3_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 17

End Sub

Private Sub txt_md3_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 18

End Sub

Private Sub txt_md3_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 19

End Sub

Private Sub txt_md3_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 2

End Sub

Private Sub txt_md3_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 20

End Sub

Private Sub txt_md3_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 21

End Sub

Private Sub txt_md3_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 22

End Sub

Private Sub txt_md3_2_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 23

End Sub

Private Sub txt_md3_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 24

End Sub

Private Sub txt_md3_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 25

End Sub

Private Sub txt_md3_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 26

End Sub

Private Sub txt_md3_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 27

End Sub

Private Sub txt_md3_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 28

End Sub

Private Sub txt_md3_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 29

End Sub

Private Sub txt_md3_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 3

End Sub

Private Sub txt_md3_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 30

End Sub

Private Sub txt_md3_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 31

End Sub

Private Sub txt_md3_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 32

End Sub

Private Sub txt_md3_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 33

End Sub

Private Sub txt_md3_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 34

End Sub

Private Sub txt_md3_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 35

End Sub

Private Sub txt_md3_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 4

End Sub

Private Sub txt_md3_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 5

End Sub

Private Sub txt_md3_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 6

End Sub

Private Sub txt_md3_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 7

End Sub

Private Sub txt_md3_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 8

End Sub

Private Sub txt_md3_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md3_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 3, 2, 9

End Sub

Private Sub txt_md3_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 3, 2, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 1

End Sub

Private Sub txt_md4_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 10

End Sub

Private Sub txt_md4_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 11

End Sub

Private Sub txt_md4_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 12

End Sub

Private Sub txt_md4_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 13

End Sub

Private Sub txt_md4_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 14

End Sub

Private Sub txt_md4_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 15

End Sub

Private Sub txt_md4_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 16

End Sub

Private Sub txt_md4_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 17

End Sub

Private Sub txt_md4_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 18

End Sub

Private Sub txt_md4_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 19

End Sub

Private Sub txt_md4_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 2

End Sub

Private Sub txt_md4_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 20

End Sub

Private Sub txt_md4_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 21

End Sub

Private Sub txt_md4_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 22

End Sub

Private Sub txt_md4_1_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 23

End Sub

Private Sub txt_md4_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 24

End Sub

Private Sub txt_md4_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 25

End Sub

Private Sub txt_md4_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 26

End Sub

Private Sub txt_md4_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 27

End Sub

Private Sub txt_md4_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 28

End Sub

Private Sub txt_md4_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 29

End Sub

Private Sub txt_md4_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 3

End Sub

Private Sub txt_md4_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 30

End Sub

Private Sub txt_md4_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 31

End Sub

Private Sub txt_md4_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 32

End Sub

Private Sub txt_md4_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 33

End Sub

Private Sub txt_md4_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 34

End Sub

Private Sub txt_md4_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 35

End Sub

Private Sub txt_md4_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 4

End Sub

Private Sub txt_md4_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 5

End Sub

Private Sub txt_md4_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 6

End Sub

Private Sub txt_md4_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 7

End Sub

Private Sub txt_md4_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 8

End Sub

Private Sub txt_md4_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 1, 9

End Sub

Private Sub txt_md4_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 1

End Sub

Private Sub txt_md4_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 10

End Sub

Private Sub txt_md4_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 11

End Sub

Private Sub txt_md4_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 12

End Sub

Private Sub txt_md4_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 13

End Sub

Private Sub txt_md4_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 14

End Sub

Private Sub txt_md4_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 15

End Sub

Private Sub txt_md4_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 16

End Sub

Private Sub txt_md4_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 17

End Sub

Private Sub txt_md4_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 18

End Sub

Private Sub txt_md4_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 19

End Sub

Private Sub txt_md4_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 2

End Sub

Private Sub txt_md4_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 20

End Sub

Private Sub txt_md4_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 21

End Sub

Private Sub txt_md4_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 22

End Sub

Private Sub txt_md4_2_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 23

End Sub

Private Sub txt_md4_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 24

End Sub

Private Sub txt_md4_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 25

End Sub

Private Sub txt_md4_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 26

End Sub

Private Sub txt_md4_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 27

End Sub

Private Sub txt_md4_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 28

End Sub

Private Sub txt_md4_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 29

End Sub

Private Sub txt_md4_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 3

End Sub

Private Sub txt_md4_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 30

End Sub

Private Sub txt_md4_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 31

End Sub

Private Sub txt_md4_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 32

End Sub

Private Sub txt_md4_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 33

End Sub

Private Sub txt_md4_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 34

End Sub

Private Sub txt_md4_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 35

End Sub

Private Sub txt_md4_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 4

End Sub

Private Sub txt_md4_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 5

End Sub

Private Sub txt_md4_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 6

End Sub

Private Sub txt_md4_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 7

End Sub

Private Sub txt_md4_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 8

End Sub

Private Sub txt_md4_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md4_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 4, 2, 9

End Sub

Private Sub txt_md4_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 4, 2, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 1

End Sub

Private Sub txt_md5_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 10

End Sub

Private Sub txt_md5_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 11

End Sub

Private Sub txt_md5_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 12

End Sub

Private Sub txt_md5_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 13

End Sub

Private Sub txt_md5_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 14

End Sub

Private Sub txt_md5_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 15

End Sub

Private Sub txt_md5_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 16

End Sub

Private Sub txt_md5_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 17

End Sub

Private Sub txt_md5_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 18

End Sub

Private Sub txt_md5_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 19

End Sub

Private Sub txt_md5_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 2

End Sub

Private Sub txt_md5_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 20

End Sub

Private Sub txt_md5_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 21

End Sub

Private Sub txt_md5_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 22

End Sub

Private Sub txt_md5_1_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 23

End Sub

Private Sub txt_md5_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 24

End Sub

Private Sub txt_md5_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 25

End Sub

Private Sub txt_md5_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 26

End Sub

Private Sub txt_md5_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 27

End Sub

Private Sub txt_md5_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 28

End Sub

Private Sub txt_md5_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 29

End Sub

Private Sub txt_md5_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 3

End Sub

Private Sub txt_md5_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 30

End Sub

Private Sub txt_md5_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 31

End Sub

Private Sub txt_md5_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 32

End Sub

Private Sub txt_md5_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 33

End Sub

Private Sub txt_md5_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 34

End Sub

Private Sub txt_md5_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 35

End Sub

Private Sub txt_md5_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 4

End Sub

Private Sub txt_md5_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 5

End Sub

Private Sub txt_md5_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 6

End Sub

Private Sub txt_md5_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 7

End Sub

Private Sub txt_md5_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 8

End Sub

Private Sub txt_md5_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 1, 9

End Sub

Private Sub txt_md5_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 1

End Sub

Private Sub txt_md5_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 10

End Sub

Private Sub txt_md5_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 11

End Sub

Private Sub txt_md5_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 12

End Sub

Private Sub txt_md5_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 13

End Sub

Private Sub txt_md5_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 14

End Sub

Private Sub txt_md5_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 15

End Sub

Private Sub txt_md5_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 16

End Sub

Private Sub txt_md5_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 17

End Sub

Private Sub txt_md5_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 18

End Sub

Private Sub txt_md5_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 19

End Sub

Private Sub txt_md5_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 2

End Sub

Private Sub txt_md5_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 20

End Sub

Private Sub txt_md5_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 21

End Sub

Private Sub txt_md5_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 22

End Sub

Private Sub txt_md5_2_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 23

End Sub

Private Sub txt_md5_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 24

End Sub

Private Sub txt_md5_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 25

End Sub

Private Sub txt_md5_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 26

End Sub

Private Sub txt_md5_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 27

End Sub

Private Sub txt_md5_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 28

End Sub

Private Sub txt_md5_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 29

End Sub

Private Sub txt_md5_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 3

End Sub

Private Sub txt_md5_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 30

End Sub

Private Sub txt_md5_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 31

End Sub

Private Sub txt_md5_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 32

End Sub

Private Sub txt_md5_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 33

End Sub

Private Sub txt_md5_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 34

End Sub

Private Sub txt_md5_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 35

End Sub

Private Sub txt_md5_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 4

End Sub

Private Sub txt_md5_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 5

End Sub

Private Sub txt_md5_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 6

End Sub

Private Sub txt_md5_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 7

End Sub

Private Sub txt_md5_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 8

End Sub

Private Sub txt_md5_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md5_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 5, 2, 9

End Sub

Private Sub txt_md5_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 5, 2, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 1

End Sub

Private Sub txt_md6_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 10

End Sub

Private Sub txt_md6_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 11

End Sub

Private Sub txt_md6_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 12

End Sub

Private Sub txt_md6_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 13

End Sub

Private Sub txt_md6_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 14

End Sub

Private Sub txt_md6_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 15

End Sub

Private Sub txt_md6_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 16

End Sub

Private Sub txt_md6_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 17

End Sub

Private Sub txt_md6_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 18

End Sub

Private Sub txt_md6_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 19

End Sub

Private Sub txt_md6_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 2

End Sub

Private Sub txt_md6_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 20

End Sub

Private Sub txt_md6_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 21

End Sub

Private Sub txt_md6_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 22

End Sub

Private Sub txt_md6_1_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 23

End Sub

Private Sub txt_md6_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 24

End Sub

Private Sub txt_md6_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 25

End Sub

Private Sub txt_md6_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 26

End Sub

Private Sub txt_md6_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 27

End Sub

Private Sub txt_md6_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 28

End Sub

Private Sub txt_md6_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 29

End Sub

Private Sub txt_md6_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 3

End Sub

Private Sub txt_md6_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 30

End Sub

Private Sub txt_md6_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 31

End Sub

Private Sub txt_md6_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 32

End Sub

Private Sub txt_md6_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 33

End Sub

Private Sub txt_md6_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 34

End Sub

Private Sub txt_md6_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 35

End Sub

Private Sub txt_md6_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 4

End Sub

Private Sub txt_md6_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 5

End Sub

Private Sub txt_md6_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 6

End Sub

Private Sub txt_md6_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 7

End Sub

Private Sub txt_md6_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 8

End Sub

Private Sub txt_md6_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 1, 9

End Sub

Private Sub txt_md6_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 1

End Sub

Private Sub txt_md6_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 10

End Sub

Private Sub txt_md6_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 11

End Sub

Private Sub txt_md6_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 12

End Sub

Private Sub txt_md6_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 13

End Sub

Private Sub txt_md6_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 14

End Sub

Private Sub txt_md6_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 15

End Sub

Private Sub txt_md6_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 16

End Sub

Private Sub txt_md6_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 17

End Sub

Private Sub txt_md6_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 18

End Sub

Private Sub txt_md6_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 19

End Sub

Private Sub txt_md6_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 2

End Sub

Private Sub txt_md6_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 20

End Sub

Private Sub txt_md6_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 21

End Sub

Private Sub txt_md6_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 22

End Sub

Private Sub txt_md6_2_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 23

End Sub

Private Sub txt_md6_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 24

End Sub

Private Sub txt_md6_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 25

End Sub

Private Sub txt_md6_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 26

End Sub

Private Sub txt_md6_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 27

End Sub

Private Sub txt_md6_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 28

End Sub

Private Sub txt_md6_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 29

End Sub

Private Sub txt_md6_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 3

End Sub

Private Sub txt_md6_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 30

End Sub

Private Sub txt_md6_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 31

End Sub

Private Sub txt_md6_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 32

End Sub

Private Sub txt_md6_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 33

End Sub

Private Sub txt_md6_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 34

End Sub

Private Sub txt_md6_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 35

End Sub

Private Sub txt_md6_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 4

End Sub

Private Sub txt_md6_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 5

End Sub

Private Sub txt_md6_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 6

End Sub

Private Sub txt_md6_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 7

End Sub

Private Sub txt_md6_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 8

End Sub

Private Sub txt_md6_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md6_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 6, 2, 9

End Sub

Private Sub txt_md6_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 6, 2, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_1_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 1

End Sub

Private Sub txt_md7_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_10_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 10

End Sub

Private Sub txt_md7_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_11_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 11

End Sub

Private Sub txt_md7_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_12_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 12

End Sub

Private Sub txt_md7_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_13_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 13

End Sub

Private Sub txt_md7_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_14_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 14

End Sub

Private Sub txt_md7_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_15_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 15

End Sub

Private Sub txt_md7_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_16_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 16

End Sub

Private Sub txt_md7_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_17_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 17

End Sub

Private Sub txt_md7_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_18_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 18

End Sub

Private Sub txt_md7_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_19_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 19

End Sub

Private Sub txt_md7_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_2_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 2

End Sub

Private Sub txt_md7_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_20_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 20

End Sub

Private Sub txt_md7_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_21_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 21

End Sub

Private Sub txt_md7_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_22_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 22

End Sub

Private Sub txt_md7_1_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_23_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 23

End Sub

Private Sub txt_md7_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_24_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 24

End Sub

Private Sub txt_md7_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_25_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 25

End Sub

Private Sub txt_md7_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_26_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 26

End Sub

Private Sub txt_md7_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_27_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 27

End Sub

Private Sub txt_md7_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_28_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 28

End Sub

Private Sub txt_md7_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_29_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 29

End Sub

Private Sub txt_md7_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_3_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 3

End Sub

Private Sub txt_md7_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_30_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 30

End Sub

Private Sub txt_md7_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_31_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 31

End Sub

Private Sub txt_md7_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_32_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 32

End Sub

Private Sub txt_md7_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_33_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 33

End Sub

Private Sub txt_md7_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_34_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 34

End Sub

Private Sub txt_md7_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_35_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 35

End Sub

Private Sub txt_md7_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_4_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 4

End Sub

Private Sub txt_md7_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_5_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 5

End Sub

Private Sub txt_md7_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_6_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 6

End Sub

Private Sub txt_md7_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_7_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 7

End Sub

Private Sub txt_md7_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_8_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 8

End Sub

Private Sub txt_md7_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_1_9_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 1, 9

End Sub

Private Sub txt_md7_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 1, 9) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_1_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 1

End Sub

Private Sub txt_md7_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 1) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_10_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 10

End Sub

Private Sub txt_md7_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 10) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_11_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 11

End Sub

Private Sub txt_md7_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 11) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_12_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 12

End Sub

Private Sub txt_md7_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 12) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_13_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 13

End Sub

Private Sub txt_md7_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 13) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_14_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 14

End Sub

Private Sub txt_md7_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 14) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_15_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 15

End Sub

Private Sub txt_md7_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 15) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_16_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 16

End Sub

Private Sub txt_md7_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 16) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_17_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 17

End Sub

Private Sub txt_md7_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 17) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_18_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 18

End Sub

Private Sub txt_md7_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 18) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_19_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 19

End Sub

Private Sub txt_md7_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 19) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_2_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 2

End Sub

Private Sub txt_md7_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 2) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_20_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 20

End Sub

Private Sub txt_md7_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 20) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_21_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 21

End Sub

Private Sub txt_md7_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 21) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_22_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 22

End Sub

Private Sub txt_md7_2_22_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 22) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_23_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 23

End Sub

Private Sub txt_md7_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 23) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_24_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 24

End Sub

Private Sub txt_md7_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 24) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_25_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 25

End Sub

Private Sub txt_md7_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 25) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_26_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 26

End Sub

Private Sub txt_md7_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 26) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_27_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 27

End Sub

Private Sub txt_md7_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 27) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_28_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 28

End Sub

Private Sub txt_md7_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 28) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_29_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 29

End Sub

Private Sub txt_md7_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 29) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_3_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 3

End Sub

Private Sub txt_md7_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 3) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_30_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 30

End Sub

Private Sub txt_md7_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 30) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_31_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 31

End Sub

Private Sub txt_md7_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 31) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_32_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 32

End Sub

Private Sub txt_md7_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 32) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_33_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 33

End Sub

Private Sub txt_md7_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 33) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_34_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 34

End Sub

Private Sub txt_md7_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 34) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_35_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 35

End Sub

Private Sub txt_md7_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 35) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_4_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 4

End Sub

Private Sub txt_md7_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 4) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_5_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 5

End Sub

Private Sub txt_md7_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 5) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_6_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 6

End Sub

Private Sub txt_md7_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 6) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_7_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 7

End Sub

Private Sub txt_md7_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 7) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_8_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 8

End Sub

Private Sub txt_md7_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 8) = True Then
        Cancel = True
    End If

End Sub

Private Sub txt_md7_2_9_AfterUpdate()
    CheckArtUndGewichtung 2, 7, 2, 9

End Sub

Private Sub txt_md7_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 7, 2, 9) = True Then
        Cancel = True
    End If

End Sub

' Eingabe von SchulaufgabenNoten (1.Schulaufgabe, 1. Halbjahr, Schüler 1-35)
Private Sub txt_sa1_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 1, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von SchulaufgabenNoten (2.Schulaufgabe, 1. Halbjahr, Schüler 1-35)
Private Sub txt_sa2_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von SchulaufgabenNoten (1.Schulaufgabe, 2. Halbjahr, Schüler 1-35)
Private Sub txt_sa1_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa1_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 1, 2, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von SchulaufgabenNoten (2.Schulaufgabe, 2. Halbjahr, Schüler 1-35)
Private Sub txt_sa2_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 1, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_sa2_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(1, 2, 2, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von sonstigen Leistungen (1.sonst. Leistung, 1. Halbjahr, Schüler 1-35)
Private Sub txt_md1_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 1, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von sonstigen Leistungen (2.sonst. Leistung, 1. Halbjahr, Schüler 1-35)
Private Sub txt_md2_1_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_1_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 1, 35) = True Then
        Cancel = True
    End If
End Sub


' Eingabe von sonstigen Leistungen (1.sonst. Leistung, 2. Halbjahr, Schüler 1-35)
Private Sub txt_md1_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md1_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 1, 2, 35) = True Then
        Cancel = True
    End If
End Sub

' Eingabe von sonstigen Leistungen (2.sonst. Leistung, 2. Halbjahr, Schüler 1-35)
Private Sub txt_md2_2_1_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 1) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_2_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 2) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_3_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 3) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_4_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 4) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_5_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 5) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_6_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 6) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_7_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 7) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_8_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 8) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_9_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 9) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_10_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 10) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_11_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 11) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_12_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 12) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_13_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 13) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_14_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 14) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_15_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 15) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_16_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 16) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_17_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 17) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_18_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 18) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_19_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 19) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_20_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 20) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_21_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 21) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_22_beforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 22) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_23_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 23) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_24_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 24) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_25_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 25) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_26_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 26) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_27_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 27) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_28_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 28) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_29_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 29) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_30_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 30) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_31_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 31) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_32_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 32) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_33_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 33) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_34_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 34) = True Then
        Cancel = True
    End If
End Sub

Private Sub txt_md2_2_35_BeforeUpdate(Cancel As Integer)
    If Not CheckNotenEingabeUndSchreibe(2, 2, 2, 35) = True Then
        Cancel = True
    End If
End Sub

' prüfe die Noteneingabe und schreibe sie beim jeweiligen Schüler zurück
Private Function CheckNotenEingabeUndSchreibe(lngArt As Long, _
                                            lngSpalte As Long, _
                                            lngHalbjahr As Long, _
                                            lngSchueler As Long) As Boolean
Dim StrControl As String
                                            
    Select Case lngArt
        Case 1                                  'Schulaufgaben
            StrControl = "txt_sa" & lngSpalte & "_" & lngHalbjahr & "_" & lngSchueler
        Case 2                                  'sonstige Leistungen
            StrControl = "txt_md" & lngSpalte & "_" & lngHalbjahr & "_" & lngSchueler
    End Select
                                            
    If (Me.Controls(StrControl).Value >= 0 And Me.Controls(StrControl).Value < 16) Or Me.Controls(StrControl).Value = "-" Or Me.Controls(StrControl).Value = "" Then
        Select Case lngArt
            Case 1                              ' Schulaufgaben
                Select Case lngSpalte
                    Case 1                      ' 1. Schulaufgabe
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_sa1_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_sa1_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 2                      ' 2. Schulaufgabe
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_sa2_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_sa2_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                End Select
            Case 2                              ' sonstige Leistungen
                Select Case lngSpalte
                    Case 1                      ' 1. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son1_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son1_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 2                      ' 2. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son2_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son2_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 3                      ' 3. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son3_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son3_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 4                      ' 4. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son4_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son4_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 5                      ' 5. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son5_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son5_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                    Case 6                      ' 6. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son6_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son6_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                         End Select
                    Case 7                      ' 7. sonstige Leistung
                        Select Case lngHalbjahr
                            Case 1              ' 1. Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son7_hj1 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 1, lngSchueler, Me.Controls(StrControl).Value
                            Case 2              ' 2.Halbjahr
                                myColSchuelerNoten.Item(lngSchueler).n_son7_hj2 = Me.Controls(StrControl).Value
                                HalbjahrSperrenWennNoteDa 2, lngSchueler, Me.Controls(StrControl).Value
                        End Select
                End Select
        End Select
        CheckNotenEingabeUndSchreibe = True
        booFormularAenderung = True
    Else
        MsgBox "Es dürfen nur Werte zwischen 0 und 15 Punkten oder '-' eingegeben werden", vbOKOnly, "FOSBOS Noteneingabe"
        CheckNotenEingabeUndSchreibe = False
        Me.Controls(StrControl).SelStart = 0
        Me.Controls(StrControl).SelLength = Len(Me.Controls(StrControl).Text)
        Exit Function

    End If
                                            
End Function
'prüfe ob die Art und Gewichtung für ein Feld gesetzt sind
Private Sub CheckArtUndGewichtung(lngArt As Long, _
                                            lngSpalte As Long, _
                                            lngHalbjahr As Long, _
                                            lngSchueler As Long)
    
    Select Case lngArt
        Case 1                              ' Schulaufgaben
        Case 2                              ' sonstige Leistungen
            Select Case lngSpalte
                Case 1                      ' 1. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son1_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son1_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son1_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son1_2.SetFocus
                            End If
                    End Select
                Case 2                      ' 2. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son2_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son2_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son2_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son2_2.SetFocus
                            End If
                    End Select
                Case 3                      ' 3. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son3_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son3_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son3_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son3_2.SetFocus
                            End If
                    End Select
                Case 4                      ' 4. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son4_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son4_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son4_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son4_2.SetFocus
                            End If
                    End Select
                Case 5                      ' 5. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son5_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son5_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son5_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son5_2.SetFocus
                            End If
                    End Select
                Case 6                      ' 6. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son6_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son6_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son6_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son6_2.SetFocus
                            End If
                    End Select
                Case 7                      ' 7. sonstige Leistung
                    Select Case lngHalbjahr
                        Case 1              ' 1. Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son7_hj1 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son7_1.SetFocus
                            End If
                        Case 2              ' 2.Halbjahr
                            If myColSchuelerNoten.Item(lngSchueler).uid_art_son7_hj2 = 0 Then
                                MsgBox "Bitte die Art und Gewichtung der sonstigen Leistung eingeben", vbOKOnly, "FOSBOS Noteneingabe"
                                Me.cmd_son7_2.SetFocus
                            End If
                    End Select
            End Select
    End Select
End Sub
' prüfe die Eingabe bei der Halbjahresleistung und schreibe sie beim jeweiligen Schüler zurück
Private Function CheckHalbjahrEingabeUndSchreibe(lngHalbjahr As Long, _
                                            lngSchueler As Long) As Boolean
Dim StrControl As String
Dim strControlBerechnet As String
                                            
    StrControl = "txt_hjgz" & "_" & lngHalbjahr & "_" & lngSchueler
    strControlBerechnet = "txt_hjg" & "_" & lngHalbjahr & "_" & lngSchueler
    
    If (Me.Controls(strControlBerechnet).Value >= 0 And Me.Controls(strControlBerechnet) < 16) And _
        Trim(Me.Controls(StrControl)) <> Trim(Me.Controls(strControlBerechnet)) Then
            Me.Controls(StrControl).BackColor = RGB(255, 0, 0)
            Me.Controls(StrControl).ForeColor = RGB(255, 255, 255)
    Else
            Me.Controls(StrControl).BackColor = RGB(255, 255, 255)
            Me.Controls(StrControl).ForeColor = RGB(0, 0, 0)
    End If
    
                                            
    If (Me.Controls(StrControl).Value >= 0 And Me.Controls(StrControl).Value < 16) Or Me.Controls(StrControl).Value = "-" Then
        Select Case lngHalbjahr
            Case 1                              ' 1. Halbjahr
                Select Case myColSchuelerHalbjahr.Item(lngSchueler).jahrgangsstufe_uid
                    Case 1, 5, 6, 7 'Vorklassen und Vorkurs
                        myColSchuelerHalbjahr.Item(lngSchueler).n_vkl_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 2  '11. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_11_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 3  '12. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_12_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 4  '13. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_13_hj1 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                End Select
            Case 2                              ' 2. Halbjahr
                Select Case myColSchuelerHalbjahr.Item(lngSchueler).jahrgangsstufe_uid
                    Case 1, 5, 6, 7 'Vorklassen und Vorkurs
                        myColSchuelerHalbjahr.Item(lngSchueler).n_vkl_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 2  '11. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_11_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 3  '12. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_12_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                    Case 4  '13. Jahrgangsstufe
                        myColSchuelerHalbjahr.Item(lngSchueler).n_13_hj2 = HalbjahresleistungUebertragen(2, Me.Controls(StrControl).Value)
                End Select
        End Select
        CheckHalbjahrEingabeUndSchreibe = True
        booHalbjahrAenderung = True
    Else
        MsgBox "Es dürfen nur Werte zwischen 0 und 15 Punkten eingegeben werden", vbOKOnly, "FOSBOS Noteneingabe"
        CheckHalbjahrEingabeUndSchreibe = False
        Me.Controls(StrControl).SelStart = 0
        Me.Controls(StrControl).SelLength = Len(Me.Controls(StrControl).Text)
        Exit Function

    End If
                                            
End Function

Private Function SpalteHinzufuegen(intSpalte As Integer, intHalbjahr As Integer)
' Blende den SpaltehinzufügenButton aus und mach den nächsten sichtbar
' Blende den Änderungsbutton ein
' Blende alle Textfelder dieser Spalte ein
' öffne das Formular für Gewichtung
' erhöhe für jeden Schüler die Anzahl der sonstigen Leistungen
' erhöhe für die Klassen-Facheinstellung die Anzahl der sonstigen Leistungen

Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim intSchuelerCount As Integer
Dim strSub As String

    strControl1 = "cmd_sone" & intSpalte & "_" & intHalbjahr
    strControl2 = "cmd_sone" & intSpalte + 1 & "_" & intHalbjahr
    StrControl3 = "cmd_son" & intSpalte & "_" & intHalbjahr
    
    Me.Controls(strControl2).Visible = True
    Me.Controls(strControl2).SetFocus
    Me.Controls(strControl1).Visible = False
    Me.Controls(StrControl3).Visible = True
    
    For intSchuelerCount = 1 To myColSchuelerNoten.Count
        strControl1 = "txt_md" & intSpalte & "_" & intHalbjahr & "_" & intSchuelerCount
        Me.Controls(strControl1).Visible = True
    Next intSchuelerCount
    
    Select Case intHalbjahr
        Case 1
            Select Case intSpalte
                Case 3
                    cmd_son3_1_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj1 = 3
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj1 = 3
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If

                Case 4
                    cmd_son4_1_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj1 = 4
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj1 = 4
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 5
                    cmd_son5_1_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj1 = 5
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj1 = 5
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 6
                    cmd_son6_1_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj1 = 6
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj1 = 6
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 7
                    cmd_son7_1_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj1 = 7
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj1 = 7
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
            End Select
        Case 2
            Select Case intSpalte
                Case 3
                    cmd_son3_2_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj2 = 3
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj2 = 3
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 4
                    cmd_son4_2_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj2 = 4
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj2 = 4
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 5
                    cmd_son5_2_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj2 = 5
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj2 = 5
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 6
                    cmd_son6_2_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj2 = 6
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj2 = 6
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
                Case 7
                    cmd_son7_2_Click
                    If booRueckgabeAenderungUebernehmen = True Then
                        For intSchuelerCount = 1 To myColSchuelerNoten.Count
                                myColSchuelerNoten.Item(intSchuelerCount).anz_son_hj2 = 7
                        Next intSchuelerCount
            
                        myclsKlasseNoten.anz_son_hj2 = 7
                    Else    'wieder Rückabwickeln
                        SpalteAusblenden intSpalte, intHalbjahr
                    End If
            End Select

        End Select
    
End Function
Private Function SpalteAusblenden(intSpalte As Integer, intHalbjahr As Integer)

' Rückabwicklung, wenn auf Popup-Maske abbrechen gedrückt wurde
Dim strControl1 As String
Dim strControl2 As String
Dim StrControl3 As String
Dim intSchuelerCount As Integer

    strControl1 = "cmd_sone" & intSpalte & "_" & intHalbjahr
    strControl2 = "cmd_sone" & intSpalte + 1 & "_" & intHalbjahr
    StrControl3 = "cmd_son" & intSpalte & "_" & intHalbjahr
    
    Me.Controls(strControl1).Visible = True
    Me.Controls(strControl1).SetFocus
    Me.Controls(strControl2).Visible = False
    Me.Controls(StrControl3).Visible = False
    
    For intSchuelerCount = 1 To myColSchuelerNoten.Count
        strControl1 = "txt_md" & intSpalte & "_" & intHalbjahr & "_" & intSchuelerCount
        Me.Controls(strControl1).Visible = False
    Next intSchuelerCount
    
End Function

Public Function UebergabeParameterFuellen(intArt As Integer, intSchueler As Integer, intHalbjahr As Integer) As Variant
Dim varUebergabe(0 To 20) As Variant
' Fülle die Schulaufgaben und sonstigen Leistungen in eine Variant-Varaible zur Übergabe an das Popup-Formular
Select Case intArt
    Case 1  'Schulaufgaben
        Select Case intHalbjahr
            Case 1
                varUebergabe(0) = myColSchuelerNoten.Item(intSchueler).dat_sa1_hj1
                varUebergabe(1) = myColSchuelerNoten.Item(intSchueler).dat_sa2_hj1
            Case 2
                varUebergabe(0) = myColSchuelerNoten.Item(intSchueler).dat_sa1_hj2
                varUebergabe(1) = myColSchuelerNoten.Item(intSchueler).dat_sa2_hj2
        End Select
    Case 2  ' sonstige Leistungen
        Select Case intHalbjahr
            Case 1
                varUebergabe(0) = myColSchuelerNoten.Item(intSchueler).uid_art_son1_hj1
                varUebergabe(1) = myColSchuelerNoten.Item(intSchueler).gew_son1_hj1
                varUebergabe(2) = myColSchuelerNoten.Item(intSchueler).dat_son1_hj1
                varUebergabe(3) = myColSchuelerNoten.Item(intSchueler).uid_art_son2_hj1
                varUebergabe(4) = myColSchuelerNoten.Item(intSchueler).gew_son2_hj1
                varUebergabe(5) = myColSchuelerNoten.Item(intSchueler).dat_son2_hj1
                varUebergabe(6) = myColSchuelerNoten.Item(intSchueler).uid_art_son3_hj1
                varUebergabe(7) = myColSchuelerNoten.Item(intSchueler).gew_son3_hj1
                varUebergabe(8) = myColSchuelerNoten.Item(intSchueler).dat_son3_hj1
                varUebergabe(9) = myColSchuelerNoten.Item(intSchueler).uid_art_son4_hj1
                varUebergabe(10) = myColSchuelerNoten.Item(intSchueler).gew_son4_hj1
                varUebergabe(11) = myColSchuelerNoten.Item(intSchueler).dat_son4_hj1
                varUebergabe(12) = myColSchuelerNoten.Item(intSchueler).uid_art_son5_hj1
                varUebergabe(13) = myColSchuelerNoten.Item(intSchueler).gew_son5_hj1
                varUebergabe(14) = myColSchuelerNoten.Item(intSchueler).dat_son5_hj1
                varUebergabe(15) = myColSchuelerNoten.Item(intSchueler).uid_art_son6_hj1
                varUebergabe(16) = myColSchuelerNoten.Item(intSchueler).gew_son6_hj1
                varUebergabe(17) = myColSchuelerNoten.Item(intSchueler).dat_son6_hj1
                varUebergabe(18) = myColSchuelerNoten.Item(intSchueler).uid_art_son7_hj1
                varUebergabe(19) = myColSchuelerNoten.Item(intSchueler).gew_son7_hj1
                varUebergabe(20) = myColSchuelerNoten.Item(intSchueler).dat_son7_hj1
            Case 2
                varUebergabe(0) = myColSchuelerNoten.Item(intSchueler).uid_art_son1_hj2
                varUebergabe(1) = myColSchuelerNoten.Item(intSchueler).gew_son1_hj2
                varUebergabe(2) = myColSchuelerNoten.Item(intSchueler).dat_son1_hj2
                varUebergabe(3) = myColSchuelerNoten.Item(intSchueler).uid_art_son2_hj2
                varUebergabe(4) = myColSchuelerNoten.Item(intSchueler).gew_son2_hj2
                varUebergabe(5) = myColSchuelerNoten.Item(intSchueler).dat_son2_hj2
                varUebergabe(6) = myColSchuelerNoten.Item(intSchueler).uid_art_son3_hj2
                varUebergabe(7) = myColSchuelerNoten.Item(intSchueler).gew_son3_hj2
                varUebergabe(8) = myColSchuelerNoten.Item(intSchueler).dat_son3_hj2
                varUebergabe(9) = myColSchuelerNoten.Item(intSchueler).uid_art_son4_hj2
                varUebergabe(10) = myColSchuelerNoten.Item(intSchueler).gew_son4_hj2
                varUebergabe(11) = myColSchuelerNoten.Item(intSchueler).dat_son4_hj2
                varUebergabe(12) = myColSchuelerNoten.Item(intSchueler).uid_art_son5_hj2
                varUebergabe(13) = myColSchuelerNoten.Item(intSchueler).gew_son5_hj2
                varUebergabe(14) = myColSchuelerNoten.Item(intSchueler).dat_son5_hj2
                varUebergabe(15) = myColSchuelerNoten.Item(intSchueler).uid_art_son6_hj2
                varUebergabe(16) = myColSchuelerNoten.Item(intSchueler).gew_son6_hj2
                varUebergabe(17) = myColSchuelerNoten.Item(intSchueler).dat_son6_hj2
                varUebergabe(18) = myColSchuelerNoten.Item(intSchueler).uid_art_son7_hj2
                varUebergabe(19) = myColSchuelerNoten.Item(intSchueler).gew_son7_hj2
                varUebergabe(20) = myColSchuelerNoten.Item(intSchueler).dat_son7_hj2
        End Select
    End Select
    
    UebergabeParameterFuellen = varUebergabe
    
End Function
Private Sub IndividuelleDatenSchreiben(intSchueler As Integer)
' schreibe die individuellen Daten des Schülers zurück

Dim StrControl As String

    myColSchuelerNoten.Item(intSchueler).ind_einstellung = True         'Schüler hat individuelle Einstellungen
    
    If Not IsNull(strRueckgabeSchuelerNotenSchulaufgabenHj1(0)) And Not strRueckgabeSchuelerNotenSchulaufgabenHj1(0) = "" Then
        myColSchuelerNoten.Item(intSchueler).dat_sa1_hj1 = strRueckgabeSchuelerNotenSchulaufgabenHj1(0)
    End If
    If Not IsNull(strRueckgabeSchuelerNotenSchulaufgabenHj1(1)) And Not strRueckgabeSchuelerNotenSchulaufgabenHj1(1) = "" Then
        myColSchuelerNoten.Item(intSchueler).dat_sa2_hj1 = strRueckgabeSchuelerNotenSchulaufgabenHj1(1)
    End If
        If Not IsNull(strRueckgabeSchuelerNotenSchulaufgabenHj2(0)) And Not strRueckgabeSchuelerNotenSchulaufgabenHj2(0) = "" Then
        myColSchuelerNoten.Item(intSchueler).dat_sa1_hj2 = strRueckgabeSchuelerNotenSchulaufgabenHj2(0)
    End If
    If Not IsNull(strRueckgabeSchuelerNotenSchulaufgabenHj2(1)) And Not strRueckgabeSchuelerNotenSchulaufgabenHj2(1) = "" Then
        myColSchuelerNoten.Item(intSchueler).dat_sa2_hj2 = strRueckgabeSchuelerNotenSchulaufgabenHj2(1)
    End If
    
    If myclsKlasseNoten.anz_son_hj1 > 1 Then
        myColSchuelerNoten.Item(intSchueler).gew_son1_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(0)
        myColSchuelerNoten.Item(intSchueler).gew_son2_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(1)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(0)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(0) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son1_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(0)
        End If
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(1)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(1) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son2_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(1)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj1 > 2 Then
        myColSchuelerNoten.Item(intSchueler).gew_son3_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(2)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(2)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(2) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son3_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(2)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj1 > 3 Then
        myColSchuelerNoten.Item(intSchueler).gew_son4_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(3)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(3)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(3) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son4_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(3)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj1 > 4 Then
        myColSchuelerNoten.Item(intSchueler).gew_son5_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(4)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(4)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(4) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son5_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(4)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj1 > 5 Then
        myColSchuelerNoten.Item(intSchueler).gew_son6_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(5)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(5)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(5) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son6_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(5)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj1 > 6 Then
        myColSchuelerNoten.Item(intSchueler).gew_son7_hj1 = strRueckgabeSchuelerNotenSonstigeGewHj1(6)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj1(6)) And Not strRueckgabeSchuelerNotenSonstigeDatHj1(6) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son7_hj1 = strRueckgabeSchuelerNotenSonstigeDatHj1(6)
        End If
    End If

    If myclsKlasseNoten.anz_son_hj2 > 1 Then
        myColSchuelerNoten.Item(intSchueler).gew_son1_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(0)
        myColSchuelerNoten.Item(intSchueler).gew_son2_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(1)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(0)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(0) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son1_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(0)
        End If
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(1)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(1) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son2_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(1)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj2 > 2 Then
        myColSchuelerNoten.Item(intSchueler).gew_son3_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(2)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(2)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(2) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son3_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(2)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj2 > 3 Then
        myColSchuelerNoten.Item(intSchueler).gew_son4_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(3)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(3)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(3) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son4_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(3)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj2 > 4 Then
        myColSchuelerNoten.Item(intSchueler).gew_son5_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(4)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(4)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(4) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son5_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(4)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj2 > 5 Then
        myColSchuelerNoten.Item(intSchueler).gew_son6_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(5)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(5)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(5) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son6_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(5)
        End If
    End If
    If myclsKlasseNoten.anz_son_hj2 > 6 Then
        myColSchuelerNoten.Item(intSchueler).gew_son7_hj2 = strRueckgabeSchuelerNotenSonstigeGewHj2(6)
        If Not IsNull(strRueckgabeSchuelerNotenSonstigeDatHj2(6)) And Not strRueckgabeSchuelerNotenSonstigeDatHj2(6) = "" Then
            myColSchuelerNoten.Item(intSchueler).dat_son7_hj2 = strRueckgabeSchuelerNotenSonstigeDatHj2(6)
        End If
    End If
    ' setze auf dem CMD-Button des Schülers einen Stern
    
    StrControl = "cmd_schueler_" & intSchueler
    Me.Controls(StrControl).Caption = "**  " & myColSchuelerNoten.Item(intSchueler).nachname & " " & myColSchuelerNoten.Item(intSchueler).rufname
    
    'Formular hat sich geändert setzen um die Daten später auch zu speichern beim Schließen
    booFormularAenderung = True
    
End Sub
' Schülerdaten zur Verwendung im Formular EinstellungenNotenSchueler
Public Property Get Schuelerdaten() As colSchülerNoten
    Set Schuelerdaten = myColSchuelerNoten
End Property
' Klassendaten zur Verwendung im Formular EinstellungenNotenSchueler
Public Property Get Klassendaten() As clsKlasseNoten
    Set Klassendaten = myclsKlasseNoten
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler
Public Property Let RueckgabeAenderungUebernehmen(ByVal booVar As Boolean)
    booRueckgabeAenderungUebernehmen = booVar
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Schulaufgaben HJ 1 Datum)
Public Property Let RueckgabeSchuelerdatenSchulaufgabenHj1(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSchulaufgabenHj1(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Schulaufgaben HJ 2 Datum)
Public Property Let RueckgabeSchuelerdatenSchulaufgabenHj2(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSchulaufgabenHj2(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Sonstige Daten HJ 1 Datum)
Public Property Let RueckgabeSchuelerdatenSonstigeDatHj1(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSonstigeDatHj1(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Sonstige Daten HJ 2 Datum)
Public Property Let RueckgabeSchuelerdatenSonstigeDatHj2(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSonstigeDatHj2(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Sonstige Daten HJ 1 Gewichtung)
Public Property Let RueckgabeSchuelerdatenSonstigeGewHj1(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSonstigeGewHj1(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenSchüler (Sonstige Daten HJ 2 Gewichtung)
Public Property Let RueckgabeSchuelerdatenSonstigeGewHj2(ByVal intCount As Integer, strNewValue As String)
    strRueckgabeSchuelerNotenSonstigeGewHj2(intCount) = strNewValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenKlasse (Art der Leistungserhebung)
Public Property Let RueckgabeKlasseUidArt(ByVal intValue As Integer)
    intRueckgabeKlasseUidArt = intValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenKlasse (Datum der Leistungserhebung)
Public Property Let RueckgabeKlasseDatum(ByVal strValue As String)
    strRueckgabeKlasseDatum = strValue
End Property
' Rückgabewert aus dem Formular EinstellungenNotenKlasse (Gewichtung der Leistungserhebung)
Public Property Let RueckgabeKlasseGewichtung(ByVal strValue As String)
    strRueckgabeKlasseGewichtung = strValue
End Property

