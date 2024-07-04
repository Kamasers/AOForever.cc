Attribute VB_Name = "mod_setup"
Option Explicit
Public CONALFAB As Boolean
Type tSetings
    transArboles As Boolean
    AlphaBlending As Boolean
    EfectosPelea As Boolean
    LimitFps As Boolean
    noche As Boolean
    NoFullScreen As Boolean
    videoMemory As Boolean
    rememberPass As Boolean
    tdsCursors As Boolean
    bmpCapture As Boolean
    VSync As Boolean
End Type
Public settingFile As String
Public tSetup As tSetings
Public Sub SaveIni()
    With tSetup
        Call WriteVar(settingFile, "Init", "AlphaBlending", IIf(.AlphaBlending = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "TreeTransparence", IIf(.transArboles = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "FightingEfects", IIf(.EfectosPelea = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "FpsLimit", IIf(.LimitFps = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "Night", IIf(.noche = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "NoFullScreen", IIf(.NoFullScreen = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "VideoMemory", IIf(.videoMemory = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "RememberPass", IIf(.rememberPass = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "Cursors", IIf(.tdsCursors = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "BmpScreenshot", IIf(.bmpCapture = True, "1", "0"))
        Call WriteVar(settingFile, "Init", "VSync", IIf(.VSync = True, "1", "0"))
    End With
End Sub

Public Sub LoadIni()
    settingFile = App.path & "/init/Settings.ini"
    With tSetup
        .AlphaBlending = IIf(GetVar(settingFile, "Init", "AlphaBlending") = "1", True, False)
        .transArboles = IIf(GetVar(settingFile, "Init", "TreeTransparence") = "1", True, False)
        .EfectosPelea = IIf(GetVar(settingFile, "Init", "FightingEfects") = "1", True, False)
        .LimitFps = IIf(GetVar(settingFile, "Init", "FpsLimit") = "1", True, False)
        .noche = IIf(GetVar(settingFile, "Init", "Night") = "1", True, False)
        .NoFullScreen = IIf(GetVar(settingFile, "Init", "NoFullScreen") = "1", True, False)
        .videoMemory = IIf(GetVar(settingFile, "Init", "VideoMemory") = "1", True, False)
        .rememberPass = IIf(GetVar(settingFile, "Init", "RememberPass") = "1", True, False)
        .tdsCursors = IIf(GetVar(settingFile, "Init", "Cursors") = "1", True, False)
        .bmpCapture = IIf(GetVar(settingFile, "Init", "BmpScreenshot") = "1", True, False)
        .VSync = IIf(GetVar(settingFile, "Init", "VSync") = "1", True, False)
    End With
End Sub
