Attribute VB_Name = "mod_Gui"
Option Explicit

'MADE BY: EL_SANTO
'*************DECLARACIONES RENDER CONNECT*************
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                                                                ByVal lpPrevWndFunc As Long, _
                                                                ByVal hwnd As Long, _
                                                                ByVal msg As Long, _
                                                                ByVal wParam As Long, _
                                                                ByVal lParam As Long) As Long
 

Public GuiColor As Long
Public GuiColor1 As Long
Public TitilaCounter As Long
Public Titila As Boolean 'asc = 124
Public FpsCounter As Integer
Public LastFpsCount As Long
Public FpsShow As Integer
Type tGuiObject
    StartX As Integer
    StartY As Integer
    EndX As Integer
    EndY As Integer
    Tipo As eGuiType
    Texto As String
    ActualColor As Byte
    Evento As Long 'al clickearse llama al evento.
    '**************Solo para cajas de texto***************
    CallEventWithEnter As Boolean
    HasFocus As Boolean
    IsPassWd As Boolean
    MaxLenght As Byte
    Width As Integer
    TieneBarrita As Boolean
    PassTmp As String
    
    '*****************************************************
    '*****************Solo para labels********************
    SetFocusOn As Byte
    '*****************************************************
End Type

Enum eGuiType
    Cmd = 1
    TxtBox = 2
    Label = 3
End Enum

Enum eGuiMode
    Nada = 1 'No hay ningun evento
    Over = 2 'El mouse encima
    MouseDown = 3 'El mouse apretando
End Enum

Public Gui() As tGuiObject
Public MaxGuiObj As Byte

Public Const TxtBox_H As Integer = 22
Public loopGui As Long
Public Const PassWord_GuiIndex As Byte = 2 'El index dentro del array de la gui del txtPasswd
Public Const Nombre_GuiIndex As Byte = 1 'El index dentro del array de la gui del txtNombre

Private YaInicializo As Boolean

Public Sub InitGui()
    MaxGuiObj = 2
    
    ReDim Gui(1 To MaxGuiObj)
    Gui(2).Evento = CallBack(AddressOf Conectarse)
    
    #If Desarrollo = 1 Then
        InitializeTxtBox Gui(1), 284, 400, 336, False, , True, , "Sun"
        InitializeTxtBox Gui(2), 284, 400, 400, True, , False, True, Decrypt("¢êù£òab")
    #Else
        InitializeTxtBox Gui(1), 284, 400, 336, False, , True, , ""
        InitializeTxtBox Gui(2), 284, 400, 400, True, , False, True, ""
    #End If

End Sub
Private Sub LimpiarArray()
    Dim x As Long, Vacio As tGuiObject
    For x = LBound(Gui) To UBound(Gui)
        Gui(x) = Vacio
    Next
End Sub
Private Sub InitializeLabel(ByRef Caption As String, ByRef Objeto As tGuiObject, ByVal x As Integer, _
                             ByVal Y As Integer, ByVal SetFocusOn As Byte)
    With Objeto
        .Tipo = Label
        .Width = Engine_GetTextWidth(cfonts(1), Caption)
        .StartY = Y
        .EndY = .StartY + 25
        .StartX = x - Round(.Width / 2)
        .EndX = x + Round(.Width / 2)
        .Texto = Caption
        .SetFocusOn = SetFocusOn
    End With
End Sub
Private Function PassChar(ByVal Lengh As Integer) As String
    Dim LoopC As Long
    If Lengh <= 0 Then PassChar = vbNullString: Exit Function
    For LoopC = 1 To Lengh
        PassChar = PassChar & "*"
    Next LoopC
End Function

Private Sub InitializeTxtBox(ByRef Objeto As tGuiObject, ByVal Width As Integer, ByVal x As Integer, _
                             ByVal Y As Integer, Optional ByVal Password As Boolean = False, _
                             Optional ByVal MaxLenght As Byte = 18, _
                             Optional ByVal StartWithFocus As Boolean = False, _
                             Optional ByVal EventWithEnter As Boolean = False, _
                             Optional ByVal InitText As String = vbNullString)
    With Objeto
        .Tipo = TxtBox
        .CallEventWithEnter = EventWithEnter
        .Width = Width
        .StartY = Y
        .EndY = .StartY + TxtBox_H
        .StartX = x - Round(.Width / 2)
        .EndX = x + Round(.Width / 2)
        .CallEventWithEnter = EventWithEnter
        .HasFocus = StartWithFocus
        .IsPassWd = Password
        .MaxLenght = MaxLenght
        .Texto = InitText
        .PassTmp = PassChar(Len(InitText))
    End With
End Sub

Public Sub Connect_KeyPress(KeyAscii As Integer) 'ByRef Buff As String, KeyAscii As Integer)

    For loopGui = 1 To MaxGuiObj
        With Gui(loopGui)
            If .Tipo = TxtBox And .HasFocus Then
                If ((KeyAscii = vbKeyBack)) And Len(.Texto) > 0 Then
                    .Texto = Left(.Texto, Len(.Texto) - 1)
                    .PassTmp = PassChar(Len(.Texto))
                    Exit Sub
                End If
                
                If KeyAscii = vbKeyReturn Then
                    If .CallEventWithEnter Then
                        CallWindowProc .Evento, 0&, 0&, 0&, 0&
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
                'If KeyAscii = vbKeyEscape 'And .ExitWithEscape = True Then
                '    prgRun = False
                 '   Exit Sub
               ' End If
                
                If KeyAscii >= vbKeySpace And KeyAscii <= 250 And Len(.Texto) < .MaxLenght Then
                    .Texto = .Texto + Chr$(KeyAscii)
                    .PassTmp = PassChar(Len(.Texto))
                End If
                
                
            End If
        End With
    Next loopGui
End Sub

Public Sub Gui_Click(Optional ByVal DblClick As Boolean = False)
    Dim lastfocus As Byte, cambiofocus As Boolean, forcefocus As Byte
    For loopGui = 1 To MaxGuiObj
        With Gui(loopGui)
            If GuiEvent(Gui(loopGui), frmConnect.Mx, frmConnect.mY) = Over Then
                If .Tipo = Cmd Then
                
                    CallWindowProc .Evento, 0&, 0&, 0&, 0&
                        
                ElseIf .Tipo = TxtBox Then
                    If DblClick Then
                        .Texto = vbNullString
                        .PassTmp = vbNullString
                    End If
                    .HasFocus = True
                    cambiofocus = True
                ElseIf .Tipo = Label Then
                    forcefocus = .SetFocusOn
                    Exit For
                End If
            Else
                If .Tipo = TxtBox Then
                    If .HasFocus = True Then
                        lastfocus = loopGui
                    End If
                    .HasFocus = False 'Le sacamos el foco
                End If
            End If
        End With
    Next loopGui
    
    If forcefocus <> 0 Then
        For loopGui = 1 To MaxGuiObj
            With Gui(loopGui)
                If .Tipo = TxtBox Then
                    If loopGui = forcefocus Then
                        .HasFocus = True
                    Else
                        .HasFocus = False
                    End If
                End If
            End With
        Next loopGui
        Exit Sub
    End If
    
    If cambiofocus = False Then 'el focus no se fue a otro textbox.
        If lastfocus > 0 And lastfocus <= MaxGuiObj Then
            Gui(lastfocus).HasFocus = True
        End If
    End If
    If frmConnect.Mx > frmConnect.imgOlvidePass.Left And frmConnect.Mx < (frmConnect.imgOlvidePass.Left + frmConnect.imgOlvidePass.Width) Then
        If frmConnect.mY > frmConnect.imgOlvidePass.Top And frmConnect.mY < (frmConnect.imgOlvidePass.Top + frmConnect.imgOlvidePass.Height) Then
            frmOlvidePass.Show , frmConnect
        End If
    End If
    If frmConnect.Mx > frmConnect.imgBorrarPj.Left And frmConnect.Mx < (frmConnect.imgBorrarPj.Left + frmConnect.imgBorrarPj.Width) Then
        If frmConnect.mY > frmConnect.imgBorrarPj.Top And frmConnect.mY < (frmConnect.imgBorrarPj.Top + frmConnect.imgBorrarPj.Height) Then
            frmBorrarPj.Show , frmConnect
        End If
    End If
End Sub

Private Function CallBack(param As Long) As Long
    CallBack = param
End Function



Public Function GuiEvent(obj As tGuiObject, x As Integer, Y As Integer) As eGuiMode
    GuiEvent = eGuiMode.Nada
    With obj
        If .Tipo = TxtBox Then
            'Debug.Print "start: " & .StartY
            
            'Debug.Print "end: " & .EndY
        End If
        If x > .StartX And x < .EndX Then 'Esta en el rango horizontal del objeto
            If Y > .StartY And Y < .EndY Then 'Tambien en el vertical, ejecutamos accion
                GuiEvent = eGuiMode.Over
            End If
        End If
    End With
End Function




'*****************Render Connect************** De aca para abajo es la parte del engine del render connect.

Sub DrawGui()
    Dim HayOver As Boolean, Modo As eGuiMode

    'DrawText 491, 303, "Usuario", -1, AlphaB ' 4, AlphaB 'D3DColorARGB(AlphaB, 255, 255, 255)
    'DrawText 486, 356, "Password", -1, AlphaB ' 4, AlphaB

    'guicolor = d3dcolorargb(255,255,255,255)
    'guicolor1 = d3dcolorargb(135,255,255,255)
    'DrawText 10, 10, "X: " & frmConnect.Mx, D3DColorARGB(AlphaB, 255, 255, 255), 4, AlphaB
    'DrawText 10, 20, "Y: " & frmConnect.My, D3DColorARGB(AlphaB, 255, 255, 255), 4, AlphaB
    If GetTickCount - TitilaCounter > 700 Then
        Titila = Not Titila
        TitilaCounter = GetTickCount
    End If
    Dim CALCULO As Integer
    For loopGui = 1 To MaxGuiObj
        With Gui(loopGui)
            Modo = GuiEvent(Gui(loopGui), frmConnect.Mx, frmConnect.mY)
            If .Tipo = Cmd Then
                If HayOver = False Then ' asi no lo pone en false aunque haya un over
                    HayOver = (Modo = Over)
                End If
                If Modo = Over Then
                    If frmConnect.mb = vbLeftButton Then
                        Modo = MouseDown
                    End If
                End If
                
            ElseIf .Tipo = TxtBox Then
                If HayOver = False Then ' asi no lo pone en false aunque haya un over
                    HayOver = (Modo = Over)
                End If
                CALCULO = (Round(Engine_GetTextWidth(cfonts(2), IIf(.IsPassWd, .PassTmp, .Texto)) / 2))
                DrawText .StartX + (.Width / 2) - CALCULO, .StartY - (CaidaConst - Caida), IIf(.IsPassWd, .PassTmp, .Texto), IIf(.HasFocus, D3DColorXRGB(255, 255, 255), -1), AlphaB, , 0, 2  ' 4, AlphaB
                If .HasFocus = True And Titila = True Then
                    DrawText .StartX + (.Width / 2) + CALCULO + IIf(Len(.Texto) > 0, 0, 0), .StartY - (CaidaConst - Caida), Chr$(124), D3DColorXRGB(255, 255, 255), AlphaB, , 0
                End If
            ElseIf .Tipo = Label Then
                DrawText .StartX, .StartY - (CaidaConst - Caida), .Texto, IIf(Modo = Over, D3DColorXRGB(150, 150, 150), -1), AlphaB, False
            End If
        End With
    Next loopGui
    
    If frmConnect.Mx > frmConnect.imgOlvidePass.Left And frmConnect.Mx < (frmConnect.imgOlvidePass.Left + frmConnect.imgOlvidePass.Width) Then
        If frmConnect.mY > frmConnect.imgOlvidePass.Top And frmConnect.mY < (frmConnect.imgOlvidePass.Top + frmConnect.imgOlvidePass.Height) Then
            DrawText frmConnect.imgOlvidePass.Left, 424 - (CaidaConst - Caida), "OlvidÈ mi contraseÒa", D3DColorXRGB(255, 255, 255), AlphaB, False, 0
            If Not frmConnect.MousePointer = 11 Then
                HayOver = True
            End If
        Else
            DrawText frmConnect.imgOlvidePass.Left, 424 - (CaidaConst - Caida), "OlvidÈ mi contraseÒa", D3DColorXRGB(180, 180, 180), AlphaB, False, 0
        End If
    Else
        DrawText frmConnect.imgOlvidePass.Left, 424 - (CaidaConst - Caida), "OlvidÈ mi contraseÒa", D3DColorXRGB(180, 180, 180), AlphaB, False, 0
    End If
    
    If frmConnect.Mx > frmConnect.imgBorrarPj.Left And frmConnect.Mx < (frmConnect.imgBorrarPj.Left + frmConnect.imgBorrarPj.Width) Then
        If frmConnect.mY > frmConnect.imgBorrarPj.Top And frmConnect.mY < (frmConnect.imgBorrarPj.Top + frmConnect.imgBorrarPj.Height) Then
            DrawText frmConnect.imgBorrarPj.Left, 424 - (CaidaConst - Caida), "Borrar personaje", D3DColorXRGB(255, 255, 255), AlphaB, False, 0
            If Not frmConnect.MousePointer = 11 Then
                HayOver = True
            End If
        Else
            DrawText frmConnect.imgBorrarPj.Left, 424 - (CaidaConst - Caida), "Borrar personaje", D3DColorXRGB(180, 180, 180), AlphaB, False, 0
        End If
    Else
        DrawText frmConnect.imgBorrarPj.Left, 424 - (CaidaConst - Caida), "Borrar personaje", D3DColorXRGB(180, 180, 180), AlphaB, False, 0
    End If
    
    If frmConnect.Visible And Not frmConnect.MousePointer = 11 Then
        If HayOver Then
            frmConnect.MousePointer = vbCustom
            frmConnect.MouseIcon = picMouseIcon
        Else
            frmConnect.MousePointer = vbNormal
        End If
    End If
    
End Sub

Public Sub Conectarse()

    
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.state <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = Gui(Nombre_GuiIndex).Texto
    
    Dim aux As String
    aux = Gui(PassWord_GuiIndex).Texto
    
#If SeguridadAlkon Then
    UserPassword = md5.GetMD5String(aux)
    Call md5.MD5Reset
#Else
    UserPassword = aux
#End If
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    End If
    
End Sub

Public Sub RecuperoPass(ByVal Nick As String, ByVal NewPass As String, ByVal pin As String)

    
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.state <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = Nick
    UserPassword = NewPass
    UserPin = pin
    EstadoLogin = E_MODO.RecuperarPass
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    
End Sub

Public Sub BorroPj(ByVal Nick As String, ByVal Pass As String, ByVal pin As String)

    
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.state <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = Nick
    UserPassword = Pass
    UserPin = pin
    EstadoLogin = E_MODO.BorrarPj
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

    
End Sub
