Attribute VB_Name = "ModFrmFunktionen"
Option Compare Database
Option Explicit
'Modul für Formularoperationen
Public Function IstFormularGeoeffnet(strFormularname As String) As Boolean

    If SysCmd(acSysCmdGetObjectState, acForm, strFormularname) <> 0 Then

        If Forms(strFormularname).CurrentView <> 0 Then

            IstFormularGeoeffnet = True

        End If

    End If

End Function

