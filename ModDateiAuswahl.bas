Attribute VB_Name = "ModDateiAuswahl"
Option Explicit
 
Public Const OFN_READONLY _
    As Long = &H1
Public Const OFN_EXPLORER _
    As Long = &H80000
Public Const OFN_LONGNAMES _
    As Long = &H200000
Public Const OFN_CREATEPROMPT _
    As Long = &H2000
Public Const OFN_NODEREFERENCELINKS _
    As Long = &H100000
Public Const OFN_OVERWRITEPROMPT _
    As Long = &H2
Public Const OFN_HIDEREADONLY _
    As Long = &H4
Public Const OFS_FILE_OPEN_FLAGS _
    As Long = OFN_EXPLORER _
    Or OFN_LONGNAMES _
    Or OFN_CREATEPROMPT _
    Or OFN_NODEREFERENCELINKS
Public Const OFS_FILE_SAVE_FLAGS _
    As Long = OFN_EXPLORER _
    Or OFN_LONGNAMES _
    Or OFN_OVERWRITEPROMPT _
    Or OFN_HIDEREADONLY
 
Public Type OPENFILENAME
  nStructSize             As Long
  hwndOwner               As Long
  hInstance               As Long
  sFilter                 As String
  sCustomFilter           As String
  nCustFilterSize         As Long
  nFilterIndex            As Long
  sFile                   As String
  nFileSize               As Long
  sFileTitle              As String
  nTitleSize              As Long
  sInitDir                As String
  sDlgTitle               As String
  Flags                   As Long
  nFileOffset             As Integer
  nFileExt                As Integer
  sDefFileExt             As String
  nCustDataSize           As Long
  fnHook                  As Long
  sTemplateName           As String
End Type
 
Public Declare Function GetActiveWindow _
    Lib "user32.dll" () As Long
Public Declare Function CommDlgExtendedError _
    Lib "comdlg32.dll" () As Long
Public Declare Function GetOpenFileName _
    Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName _
    Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetShortPathName _
    Lib "kernel32.dll" Alias _
    "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
 
Public Function vbGetOpenFilename( _
       Optional sFilter _
       As String = vbNullString, _
       Optional sDefaultFileExtension _
       As String = vbNullString, _
       Optional sInitDirectory _
       As String = vbNullString, _
       Optional sDialogTitle _
       As String = vbNullString, _
       Optional lFilterIndex _
       As Long = 0, _
       Optional sInitFilename _
       As String = vbNullString, _
       Optional bReadOnlySelected _
       As Boolean = False, _
       Optional sRtnShortPath _
       As String = vbNullString, _
       Optional lExtendedError _
       As Long = 0, _
       Optional bShowSave _
       As Boolean = False) _
       As String
  '// -----------------------------------------------------
  '// Methode:  | Aufruf des Windows-Dateidialogs
  '// -----------------------------------------------------
  '// Parameter:| sämtliche Parameter sind optional:
  '//           | sFilter               Dateifilter
  '//           | sDefaultFileExtension Dateiendung
  '//           | sInitDirectory        Startordner
  '//           | sDialogTitle          Titel des Dialogs
  '//           | lFilterIndex          Index Filterauswahl
  '//           | sInitFilename         Voreinge. Dateiname
  '//           | bReadOnlySelected     Schreibschutz
  '//           | sRtnShortPath         8.3-Pfad
  '//           | lExtendedError        mögliche Fehler
  '//           | bShowSave             Dialogtyp
  '// -----------------------------------------------------
  '// Autor:    | Stefan Kulpa
  '//           | EDV Innovation & Consulting - Dormagen
  '// -----------------------------------------------------
  '// Beispiel: |?vbGetOpenFilename
  '//           | Datei öffnen
  '//           |?vbGetOpenFilename(bShowSave:=True)
  '//           | Datei speichern
  '//           |------------------------------------------
  '//           | sFilter = "Programme" & _
  '//           |            vbNullChar & "*.exe"
  '//           |?vbGetOpenFilename(sFilter:=sFilter, _
  '//           |      sDefaultFileExtension:="*.exe", _
  '//           |      sInitDirectory:="C:WINNT")
  '// -----------------------------------------------------
  Dim lCnt            As Long
  Dim lResult         As Long
  Dim sChar           As String
  Dim sHelp           As String
  Dim sBuffer         As String
  Dim sLongPath       As String
  Dim sFilePath       As String
  Dim sFileName       As String
  Dim slpstrTitle     As String
  Dim slpstrFile      As String
  Dim slpstrFileTitle As String
  Static sLastDir     As String
  Dim uOFN            As OPENFILENAME
 
  uOFN.nStructSize = Len(uOFN)
  uOFN.hwndOwner = GetActiveWindow()
  '// -----------------------------------------------------
  '// Übergebene Parameter "trimmen"
  '// -----------------------------------------------------
  sFilter = Trim(sFilter)
  sDefaultFileExtension = Trim(sDefaultFileExtension)
  sInitDirectory = Trim(sInitDirectory)
  sDialogTitle = Trim(sDialogTitle)
  sInitFilename = Trim(sInitFilename)
  '// -----------------------------------------------------
  '// Filter setzen
  '// Format: "Name" n "Ext." n "Name" n "Ext." ... nn
  '// -----------------------------------------------------
  If Len(sFilter) = 0 Then
    sFilter = "Alle Dateien" & vbNullChar & "*.*" & _
        vbNullChar & vbNullChar
  End If
  uOFN.sFilter = sFilter
  uOFN.nFilterIndex = IIf(lFilterIndex = 0, 1, _
      lFilterIndex)
  '// -----------------------------------------------------
  '// Parameter 2: sDefaultFileExtension
  '// -----------------------------------------------------
  If Len(sDefaultFileExtension) > 0 Then
    uOFN.sDefFileExt = sDefaultFileExtension
  Else
    uOFN.sDefFileExt = "*.*"
  End If
  '// -----------------------------------------------------
  '// Parameter 3: sInitDirectory
  '// -----------------------------------------------------
  If Len(sInitDirectory) = 0 Then
    If Len(sLastDir) > 0 Then
      sInitDirectory = sLastDir
    Else: sInitDirectory = CurDir
    End If
  End If
  uOFN.sInitDir = sInitDirectory
  '// -----------------------------------------------------
  '// Parameter 4: sDialogTitle
  '// -----------------------------------------------------
  If Len(Trim(sDialogTitle)) > 0 Then
    uOFN.sDlgTitle = sDialogTitle
  Else
    If Not bShowSave Then
      uOFN.sDlgTitle = "Datei öffnen"
    Else
      uOFN.sDlgTitle = "Datei speichern unter"
    End If
  End If
  '// -----------------------------------------------------
  '// Parameter 5: lDialogFlags
  '// -----------------------------------------------------
  uOFN.Flags = OFS_FILE_OPEN_FLAGS
  '// -----------------------------------------------------
  '// Parameter 6: sInitFilename
  '// -----------------------------------------------------
  If Len(sInitFilename) = 0 Then
    uOFN.sFile = Space$(1024) & vbNullChar
  Else
    uOFN.sFile = sInitFilename & Space$(1024) & vbNullChar
  End If
  uOFN.nFileSize = Len(uOFN.sFile)
  uOFN.sFileTitle = Space$(512) & vbNullChar
  uOFN.nTitleSize = Len(uOFN.sFileTitle)
  '// -----------------------------------------------------
  '// Funktion aufrufen und auswerten
  '// -----------------------------------------------------
  If bShowSave Then
    lResult = GetSaveFileName(uOFN)
  Else
    lResult = GetOpenFileName(uOFN)
  End If
  If lResult Then
    sLongPath = _
        VBA.Left(uOFN.sFile, VBA.InStr(uOFN.sFile, _
        vbNullChar) - 1)
    If Len(sLongPath) > 0 Then
      sBuffer = String(260, 0)
      GetShortPathName sLongPath, sBuffer, Len(sBuffer)
      sRtnShortPath = _
          VBA.Left(sBuffer, VBA.InStr(sBuffer, _
          vbNullChar) - 1)
    End If
    bReadOnlySelected = _
        Abs((uOFN.Flags And OFN_READONLY) = OFN_READONLY)
    vbGetOpenFilename = sLongPath
  Else
    lExtendedError = CommDlgExtendedError
  End If
 
End Function


