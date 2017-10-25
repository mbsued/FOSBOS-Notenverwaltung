VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private IntFileNr As Integer
Private myColSchueler As colSchüler
Private myColLehrer As colLehrer
Private myColSchule As colSchule
Private myColHalbjahr As colHalbjahrNoten
Private myColHalbjahrKomplett As colHalbjahrNotenKomplett
Private strDateiname As String
Private lngDateiart As Long
' Diese Klasse dient dazu die Daten einzulesen. Welche Daten eingelesen werden wird durch Dateiart übermittelt
' Dateiart  1 = Schülerdaten
'           2 = Lehrerdaten
'           3 = Schuldaten
'           4 = Halbjahresleistungen
'           5 = Halbjahresleistungen (komplett)

'Öffnen der Datei
Public Sub DateiOeffnen()
    
    IntFileNr = FreeFile()
    Open strDateiname For Binary As #IntFileNr
    
End Sub
'Schließen der Datei
Public Sub DateiSchliessen()
    Close #IntFileNr
End Sub
Public Sub DateiLesen()
'Lesen der Datei

Dim strInhalt As String

    strInhalt = Space$(LOF(IntFileNr))
    Get #IntFileNr, , strInhalt
    Select Case lngDateiart
        Case 1
            SchuelerAufsplitten strInhalt, ";"
        Case 2
            LehrerAufsplitten strInhalt, ","
        Case 3
            SchuleAufsplitten strInhalt, ";"
        Case 4
            HalbjahrNotenAufsplitten strInhalt, ";"
        Case 5
            HalbjahrNotenKomplettAufsplitten strInhalt, ";"
    End Select

End Sub

'Aufsplitten der Datei und füllen der Collection Schüler
Private Sub SchuelerAufsplitten(ByVal strInhalt As String, StrDelimiter As String)

Dim ArrayDatei As Variant
Dim ArrayZeile As Variant
Dim ArrayIntWirt As Variant

Dim lngSchuelernummer As Long
Dim strSchulart As String
Dim strKlasse As String
Dim strNachname As String
Dim strRufname As String
Dim strVornamen As String
Dim strGebdat As String
Dim strGebort As String
Dim strGebland As String
Dim strBekenntnis As String
Dim strReligion As String
Dim strGeschlecht As String
Dim strJgst As String
Dim strAusb As String
Dim strSprache As String
Dim strEintrittsdatum As String
Dim strProbezeit As String

Dim lngSchulart As Long
Dim lngGeschlecht As Long
Dim lngJgst As Long
Dim lngAusb As Long
Dim lngSprache As Long

Dim strAusbKurz As String

Dim intCount As Integer

    On Error GoTo Err_Schueler_Aufsplitten
    
    ArrayDatei = Split(strInhalt, vbCrLf, , vbBinaryCompare)
        
    Set myColSchueler = New colSchüler
    
    SysCmd acSysCmdInitMeter, "Schüler einlesen: ", UBound(ArrayDatei)
    
    For intCount = 0 To UBound(ArrayDatei)
            SysCmd acSysCmdUpdateMeter, intCount
        
            ArrayZeile = Split(ArrayDatei(intCount), StrDelimiter, , vbBinaryCompare)
            If UBound(ArrayZeile) > 0 Then
                lngSchuelernummer = ArrayZeile(0)
                strSchulart = ArrayZeile(1)
                strKlasse = ArrayZeile(2)
                strNachname = ArrayZeile(3)
                strRufname = ArrayZeile(4)
                strVornamen = ArrayZeile(5)
                strGebdat = ArrayZeile(6)
                strGebort = ArrayZeile(7)
                strGebland = ArrayZeile(8)
                strBekenntnis = ArrayZeile(9)
                strReligion = ArrayZeile(10)
                strGeschlecht = ArrayZeile(11)
                strJgst = ArrayZeile(12)
                strAusb = Replace(ArrayZeile(13), vbCrLf, "", , , vbBinaryCompare)
                If InStr(1, strAusb, "!", vbBinaryCompare) > 0 Then 'Sprache hängt dran
                    ArrayIntWirt = Split(strAusb, "!", , vbBinaryCompare)
                    strAusb = ArrayIntWirt(0)
                    strSprache = ArrayIntWirt(1)
                End If
                strEintrittsdatum = ArrayZeile(14)
                strProbezeit = ArrayZeile(15)
                lngSchulart = CheckSchulart(strSchulart, strNachname, strRufname)
                lngGeschlecht = CheckGeschlecht(strGeschlecht, strNachname, strRufname)
                lngJgst = CheckJahrgangsstufe(strJgst, strAusb, lngSchulart, strNachname, strRufname)
                lngAusb = CheckAusbildungsrichtung(strJgst, strAusb, strNachname, strRufname)
                lngSprache = CheckSprache(lngAusb, strSprache)
                strAusbKurz = AusbildungsrichtungKurzbezeichnungSetzen(lngAusb)
                myColSchueler.Add lngSchuelernummer, lngSchulart, strKlasse, strNachname, strRufname, strVornamen, strGebdat, _
                                    strGebort, strGebland, strBekenntnis, strReligion, lngGeschlecht, lngJgst, lngAusb, strAusbKurz, strEintrittsdatum, strProbezeit, lngSprache
            End If
    Next intCount
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Schueler_Aufsplitten:
    Exit Sub

Err_Schueler_Aufsplitten:

    FehlermeldungAusgeben "Datenimport", "Aufsplitten der Schüler", Err.Number, Err.Description
    Resume Exit_Schueler_Aufsplitten
    
End Sub
'Aufsplitten der Datei und füllen der Collection Lehrer
Private Sub LehrerAufsplitten(ByVal strInhalt As String, StrDelimiter As String)
Dim ArrayDatei As Variant
Dim ArrayZeile As Variant

Dim strKuerzel As String
Dim strNachname As String
Dim strRufname As String
Dim strAmt As String
Dim strTitel As String
Dim strGeschlecht As String
Dim lngGeschlecht As Long

Dim intCount As Integer

    On Error GoTo Err_Lehrer_Aufsplitten
    
    ArrayDatei = Split(strInhalt, vbCrLf, , vbBinaryCompare)
        
    Set myColLehrer = New colLehrer
    
    SysCmd acSysCmdInitMeter, "Lehrer einlesen: ", UBound(ArrayDatei)
    
    For intCount = 0 To UBound(ArrayDatei)
            SysCmd acSysCmdUpdateMeter, intCount
        
            ArrayZeile = Split(ArrayDatei(intCount), StrDelimiter, , vbBinaryCompare)
            If UBound(ArrayZeile) > 0 Then
                strKuerzel = Replace(ArrayZeile(0), """", vbNullString, , , vbBinaryCompare)
                strNachname = Replace(ArrayZeile(1), """", vbNullString, , , vbBinaryCompare)
                strRufname = Replace(ArrayZeile(2), """", vbNullString, , , vbBinaryCompare)
                strAmt = Replace(ArrayZeile(3), """", vbNullString, , , vbBinaryCompare)
                strTitel = Replace(ArrayZeile(4), """", vbNullString, , , vbBinaryCompare)
                strGeschlecht = Replace(ArrayZeile(5), """", vbNullString, , , vbBinaryCompare)
                lngGeschlecht = CheckGeschlecht(strGeschlecht, strNachname, strRufname)
                myColLehrer.Add strNachname, strRufname, strKuerzel, strAmt, strTitel, lngGeschlecht
            End If

    Next intCount
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Lehrer_Aufsplitten:
    Exit Sub

Err_Lehrer_Aufsplitten:

    FehlermeldungAusgeben "Datenimport", "Aufsplitten der Lehrer", Err.Number, Err.Description
    Resume Exit_Lehrer_Aufsplitten
End Sub
'Aufsplitten der Datei und füllen der Collection Schule
Private Sub SchuleAufsplitten(ByVal strInhalt As String, StrDelimiter As String)
Dim ArrayDatei As Variant
Dim ArrayZeile As Variant

Dim strSchulnummer As String
Dim strSchulart As String
Dim strSchulname As String
Dim strSchulnamekurz As String

Dim intCount As Integer

    On Error GoTo Err_Schule_Aufsplitten
    
    ArrayDatei = Split(strInhalt, vbCrLf, , vbBinaryCompare)
        
    Set myColSchule = New colSchule
    
    SysCmd acSysCmdInitMeter, "Schule einlesen: ", UBound(ArrayDatei)
    
    For intCount = 0 To UBound(ArrayDatei)
            SysCmd acSysCmdUpdateMeter, intCount
        
            ArrayZeile = Split(ArrayDatei(intCount), StrDelimiter, , vbBinaryCompare)
            If UBound(ArrayZeile) > 0 Then
                strSchulnummer = ArrayZeile(0)
                strSchulname = ArrayZeile(1)
                strSchulnamekurz = ArrayZeile(2)
                strSchulart = ArrayZeile(3)
                myColSchule.Add strSchulnummer, strSchulart, strSchulname, strSchulnamekurz
            End If

    Next intCount
    
    SysCmd acSysCmdRemoveMeter
    
Exit_Schule_Aufsplitten:
    Exit Sub

Err_Schule_Aufsplitten:

    FehlermeldungAusgeben "Datenimport", "Aufsplitten der Schule", Err.Number, Err.Description
    Resume Exit_Schule_Aufsplitten
End Sub

'Aufsplitten der Datei und füllen der Collection HalbjahrNoten
Private Sub HalbjahrNotenAufsplitten(ByVal strInhalt As String, StrDelimiter As String)
Dim ArrayDatei As Variant
Dim ArrayZeile As Variant

Dim lngSchuelernummer As Long
Dim lngHalbjahr1 As Long
Dim lngHalbjahr2 As Long

Dim intCount As Integer

    On Error GoTo Err_HalbjahrNoten_Aufsplitten
    
    ArrayDatei = Split(strInhalt, vbCrLf, , vbBinaryCompare)
        
    Set myColHalbjahr = New colHalbjahrNoten
    
    SysCmd acSysCmdInitMeter, "Halbjahresleistungen einlesen: ", UBound(ArrayDatei)
    
    For intCount = 0 To UBound(ArrayDatei)
            SysCmd acSysCmdUpdateMeter, intCount
        
            ArrayZeile = Split(ArrayDatei(intCount), StrDelimiter, , vbBinaryCompare)
            If UBound(ArrayZeile) > 0 Then
                lngSchuelernummer = ArrayZeile(0)
                If Not ArrayZeile(1) = "" Then
                    lngHalbjahr1 = Nz(ArrayZeile(1), -1)
                Else
                    lngHalbjahr1 = -1
                End If
                If Not ArrayZeile(2) = "" Then
                    lngHalbjahr2 = Nz(ArrayZeile(2), -1)
                Else
                    lngHalbjahr2 = -1
                End If
                myColHalbjahr.Add lngSchuelernummer, lngHalbjahr1, lngHalbjahr2
            End If

    Next intCount
    
    SysCmd acSysCmdRemoveMeter
    
Exit_HalbjahrNoten_Aufsplitten:
    Exit Sub

Err_HalbjahrNoten_Aufsplitten:

    FehlermeldungAusgeben "Datenimport", "Aufsplitten der Halbjahresleistungen", Err.Number, Err.Description
    Resume Exit_HalbjahrNoten_Aufsplitten
End Sub

'Aufsplitten der Datei und füllen der Collection HalbjahrNotenKomplett
Private Sub HalbjahrNotenKomplettAufsplitten(ByVal strInhalt As String, StrDelimiter As String)
Dim ArrayDatei As Variant
Dim ArrayZeile As Variant

Dim lngSchuelernummer As Long
Dim strFach As String
Dim lngHalbjahr1 As Long
Dim lngHalbjahr2 As Long
Dim lngPruefung As Long
Dim strName As String
Dim strRufname As String
Dim strSchuljahr As String
Dim strJahrgang As String
Dim strKlasse As String

Dim intCount As Integer

    On Error GoTo Err_HalbjahrNoten_Aufsplitten
    
    ArrayDatei = Split(strInhalt, vbCrLf, , vbBinaryCompare)
        
    Set myColHalbjahrKomplett = New colHalbjahrNotenKomplett
    
    SysCmd acSysCmdInitMeter, "Halbjahresleistungen einlesen: ", UBound(ArrayDatei)
    
    For intCount = 0 To UBound(ArrayDatei)
            SysCmd acSysCmdUpdateMeter, intCount
        
            ArrayZeile = Split(ArrayDatei(intCount), StrDelimiter, , vbBinaryCompare)
            If UBound(ArrayZeile) > 0 Then
                lngSchuelernummer = ArrayZeile(0)
                strFach = ArrayZeile(1)
                If Not ArrayZeile(2) = "" Then
                    lngHalbjahr1 = Nz(ArrayZeile(2), -1)
                Else
                    lngHalbjahr1 = -1
                End If
                If Not ArrayZeile(3) = "" Then
                    lngHalbjahr2 = Nz(ArrayZeile(3), -1)
                Else
                    lngHalbjahr2 = -1
                End If
                If Not ArrayZeile(4) = "" Then
                    lngPruefung = Nz(ArrayZeile(4), -1)
                Else
                    lngPruefung = -1
                End If
                strName = ArrayZeile(5)
                strRufname = ArrayZeile(6)
                strSchuljahr = ArrayZeile(7)
                strJahrgang = ArrayZeile(8)
                strKlasse = ArrayZeile(9)
                myColHalbjahrKomplett.Add lngSchuelernummer, strFach, lngHalbjahr1, lngHalbjahr2, lngPruefung, strName, strRufname, strSchuljahr, strJahrgang, strKlasse
            End If

    Next intCount
    
    SysCmd acSysCmdRemoveMeter
    
Exit_HalbjahrNoten_Aufsplitten:
    Exit Sub

Err_HalbjahrNoten_Aufsplitten:

    FehlermeldungAusgeben "Datenimport", "Aufsplitten der Halbjahresleistungen", Err.Number, Err.Description
    Resume Exit_HalbjahrNoten_Aufsplitten
End Sub

' Alle Schüler
Public Property Get Schueler() As colSchüler
    Set Schueler = myColSchueler
End Property
' Alle Lehrer
Public Property Get Lehrer() As colLehrer
    Set Lehrer = myColLehrer
End Property
' Alle Schulen
Public Property Get Schule() As colSchule
    Set Schule = myColSchule
End Property
'Halbjahresleistungen
Public Property Get HalbjahrNoten() As colHalbjahrNoten
    Set HalbjahrNoten = myColHalbjahr
End Property
'Halbjahresleistungen komplett
Public Property Get HalbjahrNotenKomplett() As colHalbjahrNotenKomplett
    Set HalbjahrNotenKomplett = myColHalbjahrKomplett
End Property
' Dateiname
Public Property Get Dateiname() As String
    Dateiname = strDateiname
End Property

Public Property Let Dateiname(ByVal vNewValue As String)
    strDateiname = vNewValue
End Property

Public Property Get Dateiart() As Long
    Dateiart = lngDateiart
End Property

Public Property Let Dateiart(ByVal vNewValue As Long)
    lngDateiart = vNewValue
End Property

Private Function CheckSchulart(strSchulart As String, strNachname As String, strRufname As String) As Long
'Überprüfe die Schulart

    On Error GoTo Err_Check_Schulart
    
    Select Case UCase(strSchulart)
        Case "FOS"
            CheckSchulart = 1
        Case "BOS"
            CheckSchulart = 2
        Case Else
            Err.Raise 600, "Notenverwaltung FOSBOS", "Falsche Schulart: " & strSchulart & " beim Schüler " & strRufname & " " & strNachname
    End Select
    
Exit_Check_Schulart:
    Exit Function

Err_Check_Schulart:

    FehlermeldungAusgeben "Datenimport", "Überprüfen der Schulart", Err.Number, Err.Description
    Resume Exit_Check_Schulart
    
End Function
Private Function CheckGeschlecht(strGeschlecht As String, strNachname As String, strRufname As String) As Long
'Überprüfe das Geschlecht
    
    On Error GoTo Err_Check_Geschlecht
    
    Select Case UCase(strGeschlecht)
        Case "W"
            CheckGeschlecht = 1
        Case "M"
            CheckGeschlecht = 2
        Case Else
            Err.Raise 601, "Notenverwaltung FOSBOS", "Falsches Geschlecht: " & strGeschlecht & " bei Person " & strRufname & " " & strNachname
    End Select
    
Exit_Check_Geschlecht:
    Exit Function

Err_Check_Geschlecht:

    FehlermeldungAusgeben "Datenimport", "Überprüfen des Geschlechts", Err.Number, Err.Description
    Resume Exit_Check_Geschlecht
    
End Function
Private Function CheckJahrgangsstufe(strJgst As String, strAusb As String, lngSchulart As Long, strNachname As String, strRufname As String) As Long
'Überprüfe die Jahrgangsstufe

    On Error GoTo Err_Check_Jahrgangsstufe
    
    Select Case UCase(strJgst)
        Case "VKL", "10"
            If IsNull(strAusb) Or strAusb = "" Then     'Integrationsvorklasse
                CheckJahrgangsstufe = 7
            Else
                CheckJahrgangsstufe = 1
            End If
        Case "11"
            If UCase(strAusb) = "VG" Then       'BOS-Vorkurs ganzjährig
                CheckJahrgangsstufe = 5
            Else
                If UCase(strAusb) = "VH" Then   'BOS Vorkurs halbjährig
                    CheckJahrgangsstufe = 6
                Else
                    If lngSchulart = 2 Then     'BOS dann Vorklasse
                        CheckJahrgangsstufe = 1
                    Else
                        CheckJahrgangsstufe = 2 'FOS 11. Jahrgangsstufe
                    End If
                End If
            End If
        Case "12"
            CheckJahrgangsstufe = 3
        Case "13"
            CheckJahrgangsstufe = 4
        Case Else
            Err.Raise 602, "Notenverwaltung FOSBOS", "Falsche Jahrgangsstufe: " & strJgst & " beim Schüler " & strRufname & " " & strNachname
    End Select
    
Exit_Check_Jahrgangsstufe:
    Exit Function

Err_Check_Jahrgangsstufe:

    FehlermeldungAusgeben "Datenimport", "Überprüfen der Jahrgangsstufe", Err.Number, Err.Description
    Resume Exit_Check_Jahrgangsstufe
    
End Function
Private Function CheckAusbildungsrichtung(strJgst As String, strAusb As String, strNachname As String, strRufname As String) As Long
'Überprüfe die Jahrgangsstufe

    On Error GoTo Err_Check_Ausbildungsrichtung
    
    Select Case UCase(strAusb)
        Case "A"
            CheckAusbildungsrichtung = 1
        Case "ABU"
            CheckAusbildungsrichtung = 1
        Case "G"
            CheckAusbildungsrichtung = 2
        Case "U"
            CheckAusbildungsrichtung = 3
        Case "I", "IW", "E"
            CheckAusbildungsrichtung = 4
        Case "S"
            CheckAusbildungsrichtung = 5
        Case "T"
            CheckAusbildungsrichtung = 6
        Case "W"
            CheckAusbildungsrichtung = 7
        Case "WVR"
            CheckAusbildungsrichtung = 7
        Case "VG", "VH"
            If strJgst = "11" Then
                CheckAusbildungsrichtung = 8
            Else
                Err.Raise 603, "Notenverwaltung FOSBOS", "Falsche Ausbildungsrichtung: " & strAusb & " beim Schüler " & strRufname & " " & strNachname
            End If
        Case ""
            If strJgst = "10" Then
                CheckAusbildungsrichtung = 9
            Else
                Err.Raise 603, "Notenverwaltung FOSBOS", "Falsche Ausbildungsrichtung: " & strAusb & " beim Schüler " & strRufname & " " & strNachname
            End If
        Case Else
            Err.Raise 603, "Notenverwaltung FOSBOS", "Falsche Ausbildungsrichtung: " & strAusb & " beim Schüler " & strRufname & " " & strNachname
    End Select
    
Exit_Check_Ausbildungsrichtung:
    Exit Function

Err_Check_Ausbildungsrichtung:

    FehlermeldungAusgeben "Datenimport", "Überprüfen der Ausbildungsrichtung", Err.Number, Err.Description
    Resume Exit_Check_Ausbildungsrichtung

End Function
Private Function AusbildungsrichtungKurzbezeichnungSetzen(lngAusb As Long) As String
' setze die Kurzbezeichnung der Ausbildungsrichtung
    Select Case lngAusb
        Case 1
            AusbildungsrichtungKurzbezeichnungSetzen = "A"
        Case 2
            AusbildungsrichtungKurzbezeichnungSetzen = "G"
        Case 3
            AusbildungsrichtungKurzbezeichnungSetzen = "U"
        Case 4
            AusbildungsrichtungKurzbezeichnungSetzen = "I"
        Case 5
            AusbildungsrichtungKurzbezeichnungSetzen = "S"
        Case 6
            AusbildungsrichtungKurzbezeichnungSetzen = "T"
        Case 7
            AusbildungsrichtungKurzbezeichnungSetzen = "W"
        Case 8
            AusbildungsrichtungKurzbezeichnungSetzen = "VKU"
        Case 9
            AusbildungsrichtungKurzbezeichnungSetzen = "FVI"
        Case Else
            AusbildungsrichtungKurzbezeichnungSetzen = "X"
    End Select
End Function
Private Function CheckSprache(lngAusb As Long, strSprache As String) As Long
' Sprache anhand der übergebenen Sprache setzen nur bei Int. Wirtschaft
' Französisch = 19 ( aus Tabelle TabFach)
' Spanisch = 20
' Französisch fortg. = 41
' Spanisch fortg. = 42
' ansonsten 0 zurückliefern

    If lngAusb = 4 Then
        Select Case UCase(strSprache)
            Case "F"    'Französisch
                CheckSprache = 19
            Case "SP"   'Spanisch
                CheckSprache = 20
            Case "FF"   'Französisch fortgeführt
                CheckSprache = 41
            Case "SF"   'Spanisch fortgeführt
                CheckSprache = 42
            Case Else
                CheckSprache = 0
        End Select
    Else
        CheckSprache = 0
    End If
    
End Function
