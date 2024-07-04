Attribute VB_Name = "modConectarCaida"
Option Explicit

Public AlphaB As Byte
Public Caida As Integer
Public ModoCaida As Byte
Public CaidaConst As Integer
Public Sub IniciarCaida(Modo As Byte)
    If Modo = 0 Then
        Caida = 0
    End If
    ModoCaida = Modo
        
    Call Audio.PlayWave(SND_CAIDA)
End Sub

Public Sub EfectoCaida()
    CaidaConst = 580
    If ModoCaida = 0 Then
        If Caida < CaidaConst Then
            Caida = Caida + 10
        End If
    Else
        If Caida > 0 Then
            Caida = Caida - 10
        Else
            If MsgBox("Seguro que deseas salir?", vbYesNo) = vbYes Then
                prgRun = False
            Else
                IniciarCaida 0
            End If
        End If
    End If

End Sub
