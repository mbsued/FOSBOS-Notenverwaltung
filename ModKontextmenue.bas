Attribute VB_Name = "ModKontextmenue"
Option Compare Database
Option Explicit
Public Function ErzeugeKontextmenue()

' Anlegen des Kontextmen�1

Dim cmb As CommandBar
    On Error Resume Next

    CommandBars("Kontextmen�1").Delete


    Set cmb = CommandBars.Add("Kontextmen�1", _
               msoBarPopup, False, False)
    With cmb
        .Controls.Add msoControlButton, _
                  4, , , True  ' Drucken
        .Controls.Add msoControlButton, _
                  109, , , True  ' Seitenansicht
        .Controls.Add msoControlButton, _
                  12951, , , True  ' PDF oder XPS
        .Controls.Add msoControlButton, _
                  14782, , , True  ' Schlie�en
    End With
        
End Function
