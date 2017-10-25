Attribute VB_Name = "ModImport"
Option Compare Database
Option Explicit
' Importieren der Schülerdaten in die jeweiligen Tabellen
' Schritt 1: Einlesen der Daten in eine Collection
' Schritt 2: Prüfung ob der Schüler existiert, wenn ja: stimmt Klasse/Klassengruppe?
'                                                           ja: sind Einträge der Fächer vorhanden?
'                                                               nein: SchülerNoteneinträge anlegen
'                                                           nein: Dem Schüler die Klassengruppe zuweisen und die SchülerNoteneinträge anlegen
'                                              wenn nein: gibt es die Klasse?
'                                                           ja: Schüler anlegen und SchülerNoteneinträge anlegen
'                                                           nein: Klasse anlegen, KlassenNoteneinträge anlegen und Schülernoteneinträge anlegen

Sub SchuelerImportieren(strDateiname As String)

Dim myClsDatei As clsDatei
Dim myClsDb As clsDB
Dim myColSchueler As colSchüler

On Error GoTo Err_SchuelerImportieren

' Einlesen der Schüler

    Set myClsDatei = New clsDatei
    
    myClsDatei.Dateiname = strDateiname
    myClsDatei.Dateiart = 1
    myClsDatei.DateiOeffnen
    myClsDatei.DateiLesen
    myClsDatei.DateiSchliessen
    
    Set myColSchueler = myClsDatei.Schueler
    
' Schüler in Datenbank aufnehmen
    
    Set myClsDb = New clsDB
    
    Set myClsDb.Schueler = myColSchueler
    myClsDb.SchuelerImportieren
    MeldungAusgeben "Import der Schülerdaten erfolgreich abgeschlossen!"
    
Exit_SchuelerImportieren:

    Set myColSchueler = Nothing
    Set myClsDb = Nothing
    Set myClsDatei = Nothing
    Exit Sub
    
Err_SchuelerImportieren:
    
    FehlermeldungAusgeben "ModImport", Err.Source, Err.Number, Err.Description
    Resume Exit_SchuelerImportieren
End Sub

Sub LehrerImportieren(strDateiname As String)

Dim myClsDatei As clsDatei
Dim myClsDb As clsDB
Dim myColLehrer As colLehrer

On Error GoTo Err_LehrerImportieren

' Einlesen der Lehrer

    Set myClsDatei = New clsDatei
    
    myClsDatei.Dateiname = strDateiname
    myClsDatei.Dateiart = 2
    myClsDatei.DateiOeffnen
    myClsDatei.DateiLesen
    myClsDatei.DateiSchliessen
    
    Set myColLehrer = myClsDatei.Lehrer
    
' Schüler in Datenbank aufnehmen
    
    Set myClsDb = New clsDB
    
    Set myClsDb.Lehrer = myColLehrer
    myClsDb.LehrerImportieren
    MeldungAusgeben "Import der Lehrerdaten erfolgreich abgeschlossen!"
    
Exit_LehrerImportieren:

    Set myColLehrer = Nothing
    Set myClsDb = Nothing
    Set myClsDatei = Nothing
    Exit Sub
    
Err_LehrerImportieren:
    
    FehlermeldungAusgeben "ModImport", Err.Source, Err.Number, Err.Description
    Resume Exit_LehrerImportieren
End Sub
Sub SchuleImportieren(strDateiname As String)

Dim myClsDatei As clsDatei
Dim myClsDb As clsDB
Dim myColSchule As colSchule

On Error GoTo Err_SchuleImportieren

' Einlesen der Lehrer

    Set myClsDatei = New clsDatei
    
    myClsDatei.Dateiname = strDateiname
    myClsDatei.Dateiart = 3
    myClsDatei.DateiOeffnen
    myClsDatei.DateiLesen
    myClsDatei.DateiSchliessen
    
    Set myColSchule = myClsDatei.Schule
    
' Schule in Datenbank aufnehmen
    
    Set myClsDb = New clsDB
    
    Set myClsDb.Schule = myColSchule
    myClsDb.SchuleImportieren
    MeldungAusgeben "Import der Schuldaten erfolgreich abgeschlossen!"
    
Exit_SchuleImportieren:

    Set myColSchule = Nothing
    Set myClsDb = Nothing
    Set myClsDatei = Nothing
    Exit Sub
    
Err_SchuleImportieren:
    
    FehlermeldungAusgeben "ModImport", Err.Source, Err.Number, Err.Description
    Resume Exit_SchuleImportieren
End Sub
Sub HalbjahresNotenImportieren(strDateiname As String, lngFachUid As Long, lngKlassengruppeUid As Long, lngJahrgangsstufeUid As Long)

Dim myClsDatei As clsDatei
Dim myClsDb As clsDB
Dim myColHalbjahr As colHalbjahrNoten

On Error GoTo Err_HalbjahrImportieren

' Einlesen der Halbjahresergebnisse

    Set myClsDatei = New clsDatei
    
    myClsDatei.Dateiname = strDateiname
    myClsDatei.Dateiart = 4
    myClsDatei.DateiOeffnen
    myClsDatei.DateiLesen
    myClsDatei.DateiSchliessen
    
    Set myColHalbjahr = myClsDatei.HalbjahrNoten
    
' Noten in Datenbank aufnehmen
    
    Set myClsDb = New clsDB
    
    Set myClsDb.Halbjahr = myColHalbjahr
    myClsDb.HalbjahresNotenImportieren lngFachUid, lngKlassengruppeUid, lngJahrgangsstufeUid
    MeldungAusgeben "Import der Halbjahresleistungen erfolgreich abgeschlossen!"
    
Exit_HalbjahrImportieren:

    Set myColHalbjahr = Nothing
    Set myClsDb = Nothing
    Set myClsDatei = Nothing
    Exit Sub
    
Err_HalbjahrImportieren:
    
    FehlermeldungAusgeben "ModImport", Err.Source, Err.Number, Err.Description
    Resume Exit_HalbjahrImportieren
End Sub

Sub HalbjahresNotenKomplettImportieren(strDateiname As String)

Dim myClsDatei As clsDatei
Dim myClsDb As clsDB
Dim myColHalbjahrKomplett As colHalbjahrNotenKomplett

On Error GoTo Err_HalbjahrKomplettImportieren

' Einlesen der Halbjahresergebnisse

    Set myClsDatei = New clsDatei
    
    myClsDatei.Dateiname = strDateiname
    myClsDatei.Dateiart = 5
    myClsDatei.DateiOeffnen
    myClsDatei.DateiLesen
    myClsDatei.DateiSchliessen
    
    Set myColHalbjahrKomplett = myClsDatei.HalbjahrNotenKomplett
    
' Noten in Datenbank aufnehmen
    
    Set myClsDb = New clsDB
    
    Set myClsDb.HalbjahrKomplett = myColHalbjahrKomplett
    myClsDb.HalbjahresNotenKomplettImportieren
    MeldungAusgeben "Import der Halbjahresleistungen erfolgreich abgeschlossen!"
    
Exit_HalbjahrKomplettImportieren:

    Set myColHalbjahrKomplett = Nothing
    Set myClsDb = Nothing
    Set myClsDatei = Nothing
    Exit Sub
    
Err_HalbjahrKomplettImportieren:
    
    FehlermeldungAusgeben "ModImport", Err.Source, Err.Number, Err.Description
    Resume Exit_HalbjahrKomplettImportieren
End Sub

