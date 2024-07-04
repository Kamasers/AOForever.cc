Attribute VB_Name = "mod_T1v1"
Option Explicit
'AOForever AO

'Torneos 1vs1

'Programado por El_Santo

Private Const mapa_Torneos As Byte = 192

Private Const salaespera_inicioX As Byte = 33
Private Const salaespera_inicioY As Byte = 45

Private Const salapelea_esquina1X As Byte = 55
Private Const salapelea_esquina1y As Byte = 63
Private Const salapelea_esquina2X As Byte = 67
Private Const salapelea_esquina2y As Byte = 65

Public Type tTorneos
    nextPelea As Byte
    Activo As Boolean
    premioOro As Long
    EmpezoPelea As Boolean
    Cupos As Byte 'Cupos del torneo
    ActualCupos As Byte 'Cuantos cupos estan llenos?
    ListaUsers() As Integer
    luchando(1 To 2) As Integer
    siguienteronda() As Integer
    siguienteCupos As Byte
    maxRojas As Integer
    ClaseProhibida(1 To NUMCLASES) As Boolean
    NumProhibidas As Byte
    cuentaRegresiva As Byte
    rondasFinal(1) As Byte
    tiempoTimeout As Integer
End Type
''
Public Torneo1 As tTorneos

Public Sub NuevoTorneo(ByVal UserIndex As Integer, Cupos As Byte, ByVal maxRojas As Byte, premioOro As Long, ClaseProhibida() As Boolean)
    With Torneo1
        If CuposValidos(Cupos) = False Then
            Call WriteConsoleMsg(UserIndex, "El numero de cupos es invalido, use 2, 4, 8, 16, 32, 64 ,128", FontTypeNames.FONTTYPE_CENTINELA)
            Exit Sub
       End If
        Dim ProhibidasStr As String, LoopC As Long, Count As Byte
        .Activo = True
        .EmpezoPelea = False
        .ActualCupos = 0
        .premioOro = premioOro
        .Cupos = Cupos
        .tiempoTimeout = 60 * 7
        .siguienteCupos = Cupos / 2
        ReDim .siguienteronda(1 To .siguienteCupos)
        ReDim .ListaUsers(1 To Cupos)
        For LoopC = 1 To NUMCLASES
            .ClaseProhibida(LoopC) = ClaseProhibida(LoopC)
        Next LoopC
        .maxRojas = maxRojas
        For LoopC = 1 To NUMCLASES
            If .ClaseProhibida(LoopC) = True Then
                Count = Count + 1
                If Count = 1 Then
                    ProhibidasStr = ListaClases(LoopC)
                Else
                    ProhibidasStr = ProhibidasStr & ", " & ListaClases(LoopC)
                End If
            End If
        Next LoopC
        
        .NumProhibidas = Count
        Call MensajeTorneo(UserList(UserIndex).Name & _
                            " está organizando un [Torneo 1vs1] para " & Cupos & _
                            " participantes con un premio de " & .premioOro & " monedas de oro." & _
                            IIf(.maxRojas > 0, " Maximo de pociones rojas: " & .maxRojas _
                            & ".", "") & " Escribe '/PARTICIPAR 1VS1' para ingresar", FontTypeNames.FONTTYPE_GUILD)
        'Call MensajeTorneo(UserList(UserIndex).Name & _
                            " está organizando un [Torneo 1vs1] para " & cupos & _
                            " participantes. Clases prohibidas: " & Count & _
                            "(" & ProhibidasStr & ")" & _
                            IIf(.maxRojas > 0, "Maximo de pociones rojas: " & .maxRojas _
                            & ".", ".") & " Escribe '/PARTICIPAR 1VS1' para ingresar", FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub


Private Function Potion_Red(ByVal ID As Integer) As Long
 
    '@@ Función que devuelve las pociones rojas del usuario.
 
    Dim LoopC As Long
    Dim Total As Long
 
    With UserList(ID)
 
        For LoopC = 1 To .CurrentInventorySlots
            If .Invent.Object(LoopC).ObjIndex = 38 Then
                Total = Total + .Invent.Object(LoopC).Amount
            End If
        Next LoopC
 
        Potion_Red = Total
 
    End With
 
End Function

Private Sub CheckUsuario(ByVal UserIndex As Integer, ByRef lError As String)
    
        If UserIndex > 0 Then
            With UserList(UserIndex)
                If .EnEvento = True Then
                        lError = "Estás en un evento."
                    Exit Sub
                End If
                If .flags.Muerto <> 0 Then
                    lError = "Estás muerto."
                    Exit Sub
                End If
                If MapInfo(.Pos.map).Pk = True Then
                        lError = "Estás en una zona insegura."
                    Exit Sub
                End If
            End With
        End If
End Sub
Public Sub IngresarUsuario(ByVal UserIndex As Integer)
    With Torneo1
        Dim LoopC As Long, encontroLugar As Byte
        
        If .Activo = False Then
            Call WriteConsoleMsg(UserIndex, "Este evento no esta activo..", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Esta permitida su clase?
        If .NumProhibidas > 0 Then
            If .ClaseProhibida(UserList(UserIndex).clase) Then
                Call WriteConsoleMsg(UserIndex, "Tu clase esta prohibida en este torneo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If .maxRojas > 0 Then
            Dim userrojas As Long
            userrojas = Potion_Red(UserIndex)
            If userrojas > .maxRojas Then
                Call WriteConsoleMsg(UserIndex, "Tienes demasiadas pociones rojas, maximo: " & .maxRojas & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        Dim errr As String
        Call CheckUsuario(UserIndex, errr)
        
        If LenB(errr) > 0 Then _
            Call WriteConsoleMsg(UserIndex, errr, FontTypeNames.FONTTYPE_INFO): Exit Sub
        
        
        If UserList(UserIndex).Torneo.EnTorneo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya estas en el torneo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Si todavia no empezo la pelea, podriamos conseguirle un cupo de alguien que deslogio.
        If .Cupos = .ActualCupos Then
                Call WriteConsoleMsg(UserIndex, "Los cupos del torneo ya fueron completados. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        'Todo tranquilo, hay cupos y no empezo todavia.
        ElseIf .Cupos > .ActualCupos And .EmpezoPelea = False Then
            For LoopC = 1 To .Cupos
                If .ListaUsers(LoopC) = 0 Then 'Le encontramos un lugar
                    encontroLugar = LoopC
                End If
            Next LoopC
            .ListaUsers(encontroLugar) = UserIndex
            UserList(UserIndex).Torneo.EnTorneo = True
            UserList(UserIndex).Torneo.SuCupo = encontroLugar
            UserList(UserIndex).EnEvento = True
            .ActualCupos = .ActualCupos + 1
            
            MensajeTorneo "Torneo 1vs1> " & UserList(UserIndex).Name & " ha ingresado al torneo"
            
            Call WarpUserChar(UserIndex, mapa_Torneos, salaespera_inicioX + RandomNumber(1, 5), salaespera_inicioY, True)
            If .ActualCupos = .Cupos Then 'se llenaron los cupos. Comenzamos el torneo
                generarNuevaPelea
                
                .EmpezoPelea = True
                MensajeTorneo "Torneo 1vs1> Los cupos estan llenos. El torneo ha comenzado!!! "  '
            End If
        End If
        
    End With
End Sub
Private Sub generarNuevaPelea()
    With Torneo1
        Dim i As Long, Users(1) As Integer, gotonew As Boolean, passtoNew As Integer, findNext As Integer, LoopC As Long
        Dim c As Byte
        For i = 1 To .Cupos
            If .ListaUsers(i) <> 0 Then _
                Users(c) = .ListaUsers(i): c = c + 1
                If c = 2 Then Exit For
        Next i
        If Users(0) = -1 Then
            Call WriteConsoleMsg(Users(1), "Torneo 1vs1> Has pasado automaticamente a la siguiente ronda debido a que tu contrincante se desconecto", FontTypeNames.FONTTYPE_INFOBOLD)
            gotonew = True
            passtoNew = Users(1)
        ElseIf Users(1) = -1 Then
            Call WriteConsoleMsg(Users(0), "Torneo 1vs1> Has pasado automaticamente a la siguiente ronda debido a que tu contrincante se desconecto", FontTypeNames.FONTTYPE_INFOBOLD)
            gotonew = True
            passtoNew = Users(0)
        End If
        If gotonew = True Then
            
            Call WarpUserChar(Users(passtoNew), mapa_Torneos, salaespera_inicioX + RandomNumber(1, 5), salaespera_inicioY, True)
            If .siguienteCupos > 1 Then
                
                For LoopC = 1 To .siguienteCupos
                    If .siguienteronda(LoopC) = 0 Then
                        findNext = LoopC
                        .siguienteronda(LoopC) = Users(passtoNew)
                        UserList(Users(passtoNew)).Torneo.SuCupo = LoopC 'el index de cupo es ahora de la siguiente ronda.
                        Exit For
                    End If
                Next LoopC
                For LoopC = 1 To .Cupos
                    If (.ListaUsers(LoopC) = Users(0)) Or (.ListaUsers(LoopC) = Users(1)) Then
                        .ListaUsers(LoopC) = 0
                    End If
                Next LoopC
                If findNext = .siguienteCupos Then 'se llenaron los puestos de la ronda siguiente, se actualizan las variables para una nueva ronda.
                    'pasamos a la siguiente ronda.
                    ReDim .ListaUsers(1 To .siguienteCupos)
                    For LoopC = 1 To .siguienteCupos
                        .ListaUsers(LoopC) = .siguienteronda(LoopC)
                    Next LoopC
                    For LoopC = 1 To .siguienteCupos
                        .ListaUsers(LoopC) = .siguienteronda(LoopC)
                        If esPar(LoopC) Then
                            UserList(.ListaUsers(LoopC)).Torneo.contrincante = .ListaUsers(LoopC - 1)
                        Else
                            UserList(.ListaUsers(LoopC)).Torneo.contrincante = .ListaUsers(LoopC + 1)
                        End If
                    Next LoopC
                    .Cupos = .siguienteCupos
                    .siguienteCupos = .siguienteCupos / 2
                    
                    ReDim .siguienteronda(1 To .siguienteCupos)
                    
                    If .siguienteCupos = 1 Then _
                        MensajeTorneo "Torneo 1vs1> GRAN FINAL!!! " & UserList(.ListaUsers(1)).Name & " vs " & UserList(.ListaUsers(2)).Name
                    
                    Call generarNuevaPelea
                    
                Else 'Todavia ha de ser peleado en esta ronda
                    Call generarNuevaPelea
                End If
                'ponemos al ganador en la siguiente ronda.
                
            Else 'Es la final
                
                    MensajeTorneo "Torneo 1vs1> " & UserList(Users(passtoNew)).Name & " ha ganado el torneo. El otro usuario finalista se ha desconectado"
                    .EmpezoPelea = False
                    .Activo = False
 
                    With UserList(Users(passtoNew))
                        .Stats.GLD = UserList(Users(passtoNew)).Stats.GLD + Torneo1.premioOro
                        .EnEvento = False
                        .Torneo.EnTorneo = False
                    End With
                    
            End If
        Else
            If c = 2 Then
                Call WarpUserChar(Users(0), mapa_Torneos, salapelea_esquina1X, salapelea_esquina1y, True)
                Call WarpUserChar(Users(1), mapa_Torneos, salapelea_esquina2X, salapelea_esquina2y, True)
                Call WritePauseToggle(Users(0))
                Call WritePauseToggle(Users(1))
                
                .cuentaRegresiva = 11
                .luchando(0) = Users(0)
                .luchando(1) = Users(1)
                UserList(Users(0)).Torneo.contrincante = Users(1)
                UserList(Users(1)).Torneo.contrincante = Users(0)
                
            End If
        End If
        
    End With
End Sub
Public Sub PasaunSegundo()
    With Torneo1
        If .Activo Then
            If .EmpezoPelea = False Then
                If .tiempoTimeout > 0 Then
                        .tiempoTimeout = .tiempoTimeout - 1
                ElseIf .tiempoTimeout = 0 Then
                        Call MensajeTorneo("Torneo1vs1> El torneo se ha cancelado por falta de participantes")
                        Call CancelarTorneo
                End If
            End If
            If .cuentaRegresiva > 0 Then
                .cuentaRegresiva = .cuentaRegresiva - 1
                If .cuentaRegresiva > 0 Then
                    Call WriteConsoleMsg(.luchando(0), "Torneo 1vs1> " & .cuentaRegresiva & "...", FontTypeNames.FONTTYPE_INFOBOLD)
                    Call WriteConsoleMsg(.luchando(1), "Torneo 1vs1> " & .cuentaRegresiva & "...", FontTypeNames.FONTTYPE_INFOBOLD)
                Else
                
                    Call WriteConsoleMsg(.luchando(0), "Torneo 1vs1> YA!!!..", FontTypeNames.FONTTYPE_CONSEJOCAOS)
                    Call WriteConsoleMsg(.luchando(1), "Torneo 1vs1> YA!!!..", FontTypeNames.FONTTYPE_CONSEJOCAOS)
                    Call WritePauseToggle(.luchando(0))
                    Call WritePauseToggle(.luchando(1))
                    
                End If
            End If
        End If
    End With
End Sub
Public Sub proccessDeathOrDisconnect(ByVal UserIndex As Integer)
    Dim iContrincante As Integer, LoopC As Integer
    
    
    
    Dim findNext As Byte
    Dim enc As Byte
    With Torneo1
        For LoopC = 1 To .Cupos
            If .ListaUsers(LoopC) = UserIndex Then
                enc = LoopC
            End If
        Next LoopC
        If enc = 0 Then
            For LoopC = 1 To .siguienteCupos
                If .siguienteronda(LoopC) = UserIndex Then
                    enc = LoopC
                   .siguienteronda(enc) = -1
                   Exit Sub
                End If
            Next LoopC
             'Ya habia pasado a la siguiente ronda y se desconecto
                 
        End If
            
        If .siguienteCupos > 1 Then
        
            With UserList(UserIndex)
                .EnEvento = False
                .Torneo.EnTorneo = False
                iContrincante = .Torneo.contrincante
                .Torneo.contrincante = 0
            End With
            
            Call WarpUserChar(UserIndex, 1, 50, 50, True)
            
            For LoopC = 1 To .siguienteCupos
                If .siguienteronda(LoopC) = 0 Then
                    findNext = LoopC
                    .siguienteronda(LoopC) = iContrincante
                    UserList(iContrincante).Torneo.SuCupo = LoopC 'el index de cupo es ahora de la siguiente ronda.
                    Exit For
                End If
            Next LoopC
            For LoopC = 1 To .Cupos
                If (.ListaUsers(LoopC) = UserIndex) Or (.ListaUsers(LoopC) = iContrincante) Then
                    .ListaUsers(LoopC) = 0
                End If
            Next LoopC
            If findNext = .siguienteCupos Then 'se llenaron los puestos de la ronda siguiente, se actualizan las variables para una nueva ronda.
                'pasamos a la siguiente ronda.
                ReDim .ListaUsers(1 To .siguienteCupos)
                For LoopC = 1 To .siguienteCupos
                    .ListaUsers(LoopC) = .siguienteronda(LoopC)
                Next LoopC
                For LoopC = 1 To .siguienteCupos
                    .ListaUsers(LoopC) = .siguienteronda(LoopC)
                    If esPar(LoopC) Then
                        UserList(.ListaUsers(LoopC)).Torneo.contrincante = .ListaUsers(LoopC - 1)
                    Else
                        UserList(.ListaUsers(LoopC)).Torneo.contrincante = .ListaUsers(LoopC + 1)
                    End If
                Next LoopC
                
                    
                .Cupos = .siguienteCupos
                .siguienteCupos = .siguienteCupos / 2
                ReDim .siguienteronda(1 To .siguienteCupos)
                
                If .siguienteCupos = 1 Then _
                    MensajeTorneo "Torneo 1vs1> GRAN FINAL!!! " & UserList(.ListaUsers(1)).Name & " vs " & UserList(.ListaUsers(2)).Name
                
                generarNuevaPelea
                
            Else 'Todavia ha de ser peleado en esta ronda
                Call generarNuevaPelea
            End If
            'ponemos al ganador en la siguiente ronda.
            
        Else 'Es la final
            
            Dim contB As Byte
            If .luchando(0) = UserIndex Then contB = 1 Else contB = 0
            
            .rondasFinal(contB) = .rondasFinal(contB) + 1
            
            If .rondasFinal(contB) < 3 Then
                'nueva ronda
                .cuentaRegresiva = 11
                
                Call WarpUserChar(.luchando(0), mapa_Torneos, salapelea_esquina1X, salapelea_esquina1y, True)
                Call WarpUserChar(.luchando(1), mapa_Torneos, salapelea_esquina2X, salapelea_esquina2y, True)
                Call WritePauseToggle(.luchando(0))
                Call WritePauseToggle(.luchando(1))
                
                RevivirUsuario UserIndex
            Else 'GANO 3 RONDAS YA, TERMINA TORNEO
                MensajeTorneo "Torneo 1vs1> " & UserList(iContrincante).Name & " ha ganado el torneo. El finalista gana el 30% del premio TOTAL"
                .EmpezoPelea = False
                .Activo = False
                With UserList(iContrincante)
                    .Stats.GLD = UserList(iContrincante).Stats.GLD + Torneo1.premioOro
                    .EnEvento = False
                    .Torneo.EnTorneo = False
                End With
                With UserList(UserIndex)
                    .Stats.GLD = UserList(UserIndex).Stats.GLD + (Torneo1.premioOro * 0.3)
                    .EnEvento = False
                    .Torneo.EnTorneo = False
                End With
                Call WarpUserChar(iContrincante, 1, 50, 51, True)
            End If
        End If
    End With
End Sub
Public Sub userSaleTorneo(ByVal UI As Integer)
    With Torneo1
        If .EmpezoPelea = False Then
            Dim i As Long
            For i = 1 To .Cupos
                If .ListaUsers(i) = UI Then _
                    .ListaUsers(i) = 0: Exit For
            Next i
            .ActualCupos = .ActualCupos - 1
            Call WarpUserChar(UI, 1, 50, 50, True)
            MensajeTorneo "Torneo 1vs1> " & UserList(UI).Name & " se desconecto, hay un nuevo cupo disponible"
        Else
            MensajeTorneo "Torneo 1vs1> " & UserList(UI).Name & " se desconecto, su contrincante continua victorioso por abandono a la siguiente ronda"
            Call proccessDeathOrDisconnect(UI)
        End If
    
    End With
End Sub

Private Function esPar(ByVal n As Byte) As Boolean
    Dim nstr As String, comp As String
    nstr = str(n)
    If LenB(nstr) > 1 Then
        comp = Right$(nstr, 1) 'sacamos la cifra de la derecha, para saber si es par o impar
    Else
        comp = nstr
    End If
    esPar = False
     If (comp = "0") Or (comp = "2") Or (comp = "4") Or (comp = "6") Or (comp = "8") Then
        esPar = True
    End If
End Function

Public Function CuposValidos(ByVal Cupos As Byte) As Boolean
    If (Cupos = 2 Or Cupos = 4 Or Cupos = 8 Or Cupos = 16 Or Cupos = 32 Or Cupos = 64 Or Cupos = 128) Then
        CuposValidos = True
    Else
        CuposValidos = False
    End If
End Function

Public Sub LimpiarDatos()
    With Torneo1
        .Activo = False
        .Cupos = 0
        .ActualCupos = 0
        .cuentaRegresiva = 0
        .maxRojas = 0
        Dim X As Long
        For X = 1 To NUMCLASES
            .ClaseProhibida(X) = False
        Next X
    End With
End Sub

Public Sub CancelarTorneo()
    With Torneo1
            
            .Activo = False
            
            Dim i As Long
            For i = 1 To .Cupos
                If .ListaUsers(i) > 0 Then
                    With UserList(.ListaUsers(i))
                        .Torneo.EnTorneo = False
                        .EnEvento = False
                        .Torneo.contrincante = 0
                        .Torneo.SuCupo = 0
                    End With
                    Call WarpUserChar(.ListaUsers(i), 1, 50, 50, True)
                    
                End If
            Next i
            
            
            
    End With
End Sub

Public Sub MensajeTorneo(ByVal Texto As String, Optional ByVal ft As FontTypeNames = FontTypeNames.FONTTYPE_GUILD)
    Dim data As String
    data = PrepareMessageConsoleMsg(Texto, ft)
    Call modSendData.SendData(SendTarget.ToAll, 0, data)
End Sub



















