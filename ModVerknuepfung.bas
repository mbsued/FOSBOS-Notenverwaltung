Attribute VB_Name = "ModVerknuepfung"
Option Compare Database
Option Explicit
 
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private dbs As DAO.Database
 
' Funktion zum Neuverknüpfen des Backends
' Quelle: www.dbwiki.net oder www.dbwiki.de
'
' strPath: Verzeichnis, in dem das/die Backend(s) liegt/liegen
' (Es darf nur EIN Verzeichnis sein, kann jedoch mehrere BE-Dateien enthalten)
' Rückgabe: True bei Erfolg der kompletten Neuverknüpfung
 
Function RelinkDb(strPath As String) As Boolean
Dim flag As Boolean, bStat As Boolean
Dim i As Long, nCount As Long
Dim strFile As String, strConnect As String
Dim tdf As DAO.TableDef
Dim rs() As DAO.Recordset
Dim cFiles As New VBA.Collection
 
  On Error GoTo Fehler
 
  If (strPath = "") Then Err.Raise 23001, "RelinkDB", "Leere Pfadangabe"
  If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
  ' Testen, ob Verzeichnis Dateien enthält...
  If Dir(strPath & "\*") = "" Then Err.Raise 23002, _
                           "RelinkDB", "Ungülige Pfadangabe"
 
  Set dbs = CurrentDb
  ' Zwischenspeicher, ob Statusleiste angezeigt ist...
  bStat = Application.GetOption("Show Status Bar")
  Application.SetOption "Show Status Bar", True   'Statusleiste anzeigen und
  ' Meldung "Neuverknüpfen" anzeigen
  SysCmd acSysCmdInitMeter, "Neuverknüpfen der Daten mit dem Backend...", 100
 
  ' Testen, ob alle benötigten Dateien im
  ' Backendverzeichnis vorhanden sind (s. Funktion unten)
  strFile = TestFilesOK(strPath)
 
  If strFile <> "" Then Err.Raise 23003, "RelinkDB", _
                        "Benötigte Datei " & strFile & " nicht gefunden." & _
                        vbNewLine & "...Abbruch!"
 
  ' Zahl der verknüpften Tabellen ermitteln für Fortschrittsanzeige...
  nCount = dbs.OpenRecordset("SELECT COUNT(*) FROM MSysObjects WHERE " & _
                             "MSysObjects.Database IS NOT NULL", dbOpenSnapshot)(0)
  i = 0
  ' Diese beiden Optionen sollen den Sperrmechanismus von JET beschleunigen
  DBEngine.SetOption dbLockDelay, 20
  DBEngine.SetOption dbLockRetry, 5
 
  ' Alle (verknüpften) Tabellen durchgehen...
  For Each tdf In dbs.TableDefs
 
    ' Fortschrittsanzeige in Statusleiste einstellen...
    SysCmd acSysCmdUpdateMeter, , Int(5 + 95 * i / nCount)
 
    strConnect = tdf.Connect
    ' Verknüpfte Tabellen haben in der Eigenschaft "Connect" keinen Leer-String...
    ' ...aber evtl. eine ODBC-/Excel-/ETC.-Verknüpfnung.
    If strConnect <> "" Then
      If Left(strConnect, 9) = ";DATABASE" Then
        ' Diese Tabellen ausschließen.
        i = i + 1
 
        ' Dateiname des Backends aus Connect extrahieren; hier kommt die Funktion
        '  InstrRev() zum Einsatz, die erst ab A2000 zur Verfügung steht. Einen
        '  Ersatz für A97 gibt es hier im Wiki
 
        strFile = Mid(strConnect, 1 + InStrRev(strConnect, "\"))
        On Error Resume Next
        Err.Clear
 
        '  Dateinamen zur Collection hinzufügen. Wenn ein gleichnamiger Eintrag
        ' (Key!) bereits besteht, schlägt dies fehl. Andernfalls wird flag auf
        ' True gesetzt. (s.u. >>FLAG)
 
        cFiles.Add strFile, "1" & strFile
        flag = (Err.Number = 0)
        On Error GoTo Fehler
 
        ' Neuen Verknüpfungsstring erzeugen
        strConnect = ";DATABASE=" & strPath & "\" & strFile
 
        ' Nur verknüpfen, wenn es wirklich erforderlich ist
        If tdf.Connect <> strConnect Then
          tdf.Connect = strConnect
 
          ' Und das ist die eigentliche Verknüpfungsanweisung:
          tdf.RefreshLink
        End If
      End If
    End If
 
    '>> FLAG
    ' Wenn auf eine Tabelle im Backend ein Recordset geöffnet wird, so
    ' wird die Zahl der Sperrversuche zum Backend von JET herabgesetzt.
    ' Dies resultiert in einer höheren Verknüpfungsgeschwindigkeit
    ' für die weiteren Tabellen dieses Backend. Nach Erfahrungswerten
    ' beschleunigt dies den Vorgang um das 2-3-fache!
    '  Deshalb wird hier für jede Backend-Datei genau ein Recordset in
    ' einem Recordset-Array geöffnet...
    If flag Then
      ReDim Preserve rs(cFiles.Count - 1)
      Set rs(cFiles.Count - 1) = dbs.OpenRecordset(tdf.name, dbOpenDynaset)
    End If
 
  Next tdf
  ' DBEngine-Optionen wieder auf Standardwerte einstellen.
  DBEngine.SetOption dbLockDelay, 100
  DBEngine.SetOption dbLockRetry, 20
 
  dbs.TableDefs.Refresh
  RelinkDb = True
  SysCmd acSysCmdRemoveMeter  ' Fortschrittsanzeige in Statusleiste entfernen
  SysCmd acSysCmdSetStatus, " Verknüpfungen OK!"  ' Erfolgsmeldung anzeigen
  Sleep 2000      ' ...und 2 sek warten, damit man sie auch lesen kann. ;-)
  SysCmd acSysCmdClearStatus
  Application.SetOption "Show Status Bar", bStat  ' Statusleiste anzeigen, je nach
                                                  ' Zustand vor der Neuverknüpfung
Ende:
  Erase rs    ' Recordset-Array löschen; setzt alle Recordsets auf Nothing
  Set cFiles = Nothing
  Set tdf = Nothing
  Set dbs = Nothing
  Exit Function
 
Fehler:
  MsgBox Err.Description, vbCritical, "mdlRelink / RelinkDB"
  Resume Ende
 
End Function
 
' Hilfsfunktion, die ermittelt, ob alle notwendigen
' Backend-Dateien im angegebenen Verzeichnis vorhanden sind.
' Rückgabe dann ""; andernfalls String mit Name der fehlenden Datei.
Public Function TestFilesOK(strPath As String) As String
Dim tdf As TableDef, strConnect As String
On Error GoTo Fehler
 
  For Each tdf In dbs.TableDefs
    strConnect = tdf.Connect
    If strConnect <> "" Then
      If Left(strConnect, 9) = ";DATABASE" Then
        strConnect = Mid(strConnect, 1 + InStrRev(strConnect, "\"))
        If Dir(strPath & "\" & strConnect) = "" Then
          TestFilesOK = strConnect
          ' Hier Abbruch, falls Datei fehlt
          Exit For
        End If
      End If
    End If
  Next tdf
 
Ende:
  Set tdf = Nothing
  Exit Function
 
Fehler:
  MsgBox Err.Description, vbCritical, "mdlRelink / TestFilesOK"
  Resume Ende
End Function
