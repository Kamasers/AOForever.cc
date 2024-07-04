Attribute VB_Name = "mod_DeathMatch"
Option Explicit

Private Type tDeath
    UIndex() As Integer
    Activo As Boolean
    cuentaRegresiva As Integer
    CuposRestantes As Byte
    Cupos As Byte
    CaenItems As Boolean
    UsersRestantes As Byte
    Comenzado As Boolean
End Type
'AVISO CUANDO COMIENZA
Public Death As tDeath
'Death: 295, 48,81 espera
'death: 295,49,50 pelea
Public Const MAPA_DEATH As Integer = 49
Private Const ESPERA_X As Byte = 29
Private Const ESPERA_Y As Byte = 30
Private Const PELEA_X As Byte = 66
Private Const PELEA_Y As Byte = 30
Public Sub IniciarDeath(ByVal Cupos As Byte, ByVal CaenItems As Boolean)
        
    With Death
        If .Activo = True Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Ya hay un DeathMatch en curso", FontTypeNames.FONTTYPE_INFOBOLD))
            Exit Sub
        End If
        .Activo = True
        .cuentaRegresiva = 10
        .Cupos = Cupos
        .CuposRestantes = Cupos
        .CaenItems = CaenItems
        .UsersRestantes = Cupos
        .Comenzado = False
        Call LimpiarMapa(MAPA_DEATH)
        ReDim .UIndex(1 To Cupos)
        Call MensajeGlobal("DeathMatch> El evento ha comenzado. Cupos disponibles: " & Cupos & ". " & IIf(CaenItems = False, "Los items no se caen", "Con caída de items ") & ". Escribe /ENTRARDEATH para ingresar", FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Public Sub Muere_Death(ByVal userindex As Integer, Optional ByVal desconexion As Boolean = False)
    With UserList(userindex)
        If .UserDeath.EnDeath = False Then Exit Sub
        .UserDeath.EnDeath = False
        .EnEvento = False
        If Death.CaenItems Then
            If Death.Comenzado = True Then
                Call TirarTodosLosItems(userindex)
            End If
        End If
        Call WarpUserChar(userindex, .UserDeath.lastPos.map, .UserDeath.lastPos.X, .UserDeath.lastPos.Y, True, , True)
        Dim loopC As Long
        For loopC = 1 To Death.Cupos
            If Death.UIndex(loopC) = userindex Then
                Death.UIndex(loopC) = 0
            End If
        Next loopC
        
        If desconexion = True And Death.Comenzado = False Then
            Death.CuposRestantes = Death.CuposRestantes + 1
        End If
        If desconexion Then
            If .Stats.GLD >= 200000 Then
                .Stats.GLD = .Stats.GLD - 200000
            Else
                .Stats.GLD = 0
            End If
            If Death.Comenzado = True Then
                Call MensajeGlobal("DeathMatch> " & .Name & " se ha desconectado" & IIf(Death.UsersRestantes > 2, ". Quedan " & Death.UsersRestantes - 1 & " usuarios vivos.", ""), FontTypeNames.FONTTYPE_GUILD)
            Else
                Call MensajeGlobal("DeathMatch> Se ha liberado un cupo por la desconexión de " & .Name, FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            Call MensajeGlobal("DeathMatch> " & .Name & " ha muerto" & IIf(Death.UsersRestantes > 2, ". Quedan " & Death.UsersRestantes & " usuarios vivos.", "."), FontTypeNames.FONTTYPE_GUILD)
        End If
        If Not desconexion And Death.Comenzado = True Then
            Death.UsersRestantes = Death.UsersRestantes - 1
            
            If Death.UsersRestantes = 1 Then
                Call Death_Finish
            End If
        End If
    End With
End Sub

Private Sub Death_Finish()
    With Death
        '.Activo = False
       ' .Comenzado = False
        '.CuentaRegresiva = 10
        Dim loopC As Long, Winner As Integer
        For loopC = 1 To .Cupos
            If .UIndex(loopC) > 0 Then
                Winner = .UIndex(loopC)
                Exit For
            End If
        Next loopC
        If Winner <= 0 Then Exit Sub 'Raro, pero por las dudas
        Call MensajeGlobal("DeathMatch> Evento finalizado. Ganador: " & UserList(Winner).Name & ". Premio: 200000 monedas de oro", FontTypeNames.FONTTYPE_GUILD)
        With UserList(Winner)
            .Stats.GLD = .Stats.GLD + 500000
            Call WriteUpdateGold(Winner)
            If Death.CaenItems = False Then
                WarpUserChar Winner, .UserDeath.lastPos.map, .UserDeath.lastPos.X, .UserDeath.lastPos.Y, True
                Death.Activo = False
                Death.Comenzado = False
                Death.cuentaRegresiva = 10
            Else
                WriteConsoleMsg Winner, "DeathMatch> Tenés 1 minuto para agarrar tus items.", FontTypeNames.FONTTYPE_GUILD
                .UserDeath.SecondsBack = 60
            End If
        End With
        
    End With
End Sub
'Listo el modulo, creo. Ahora hay que hacer los paquetes y las llamadas(Cuando muere, desconecta)
Public Sub EnterDeath(ByVal userindex As Integer)
    With UserList(userindex)
        Dim lError As String '<=esta es la variable
        Call PuedeDeath(userindex, lError)
        If LenB(lError) <> 0 Then
            Call WriteConsoleMsg(userindex, "DeathMatch> " & lError, FontTypeNames.FONTTYPE_INFO)
            Exit Sub 'Si tiene algun error, le decimos cual es y salimos.
        End If
        With .UserDeath
            .EnDeath = True
            .lastPos = UserList(userindex).Pos
        End With
        .EnEvento = True
        With Death
            .CuposRestantes = .CuposRestantes - 1
            Dim loopC As Long, find As Byte
            For loopC = 1 To .Cupos
                If .UIndex(loopC) <= 0 Then
                    find = CByte(loopC)
                    Exit For
                End If
            Next loopC
            .UIndex(find) = userindex
            WarpUserChar userindex, MAPA_DEATH, ESPERA_X, ESPERA_Y, True, , True
            Call MensajeGlobal("DeathMatch> " & UserList(userindex).Name & " ha ingresado al evento.", FontTypeNames.FONTTYPE_GUILD)
            If .CuposRestantes = 0 Then
                Death_Go
            End If
        End With
    End With
End Sub

Public Sub PassSecondDeath()
    With Death
        'Death_Finish
        If .Activo And .Comenzado = True And .cuentaRegresiva >= 0 Then
            Select Case .cuentaRegresiva
                Case 0
                    Call MensajeGlobal("DeathMatch> ¡Ya!", FontTypeNames.FONTTYPE_GUILD)
                    Call DEATH_GO1
                
                Case Else
                    Call MensajeGlobal("DeathMatch> ¡" & .cuentaRegresiva & "!", FontTypeNames.FONTTYPE_GUILD)
            
            End Select
            .cuentaRegresiva = .cuentaRegresiva - 1
        End If
    End With
End Sub
Sub CancelarDeath()
    With Death
        If .Activo = False Then Exit Sub
        Dim X As Long
        For X = 1 To .Cupos
            If .UIndex(X) > 0 Then
                WarpUserChar .UIndex(X), UserList(.UIndex(X)).UserDeath.lastPos.map, UserList(.UIndex(X)).UserDeath.lastPos.X, UserList(.UIndex(X)).UserDeath.lastPos.Y, True, , True
                UserList(.UIndex(X)).UserDeath.EnDeath = False
                UserList(.UIndex(X)).EnEvento = False
            End If
        Next X
        .Activo = False
        .CaenItems = False
        .Comenzado = False
        .cuentaRegresiva = 10
        Call MensajeGlobal("DeathMatch> El evento ha sido cancelado", FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Function death_PuedeAtacar(ByVal userindex As Integer) As Boolean
    With Death
        If .Activo = True And .Comenzado = True And .cuentaRegresiva <= 0 Then
            death_PuedeAtacar = True
            Exit Function
        End If
        
        If .Activo = True And .cuentaRegresiva > 0 Then
            death_PuedeAtacar = False
            WriteConsoleMsg userindex, "DeathMatch> Espera que termine la cuenta regresiva", FontTypeNames.FONTTYPE_GUILD
        End If
    End With
End Function

Private Sub DEATH_GO1()
    Dim X As Long
    For X = 1 To Death.Cupos
        If Death.UIndex(X) > 0 Then
            WritePauseToggle Death.UIndex(X)
        End If
    Next X
End Sub
Private Sub Death_Go()
    With Death
        .Comenzado = True
        '.CuentaRegresiva = 10
        
        Dim X As Long
        For X = 1 To .Cupos
            If .UIndex(X) > 0 Then
                WarpUserChar .UIndex(X), MAPA_DEATH, PELEA_X, PELEA_Y, True, , True
                WritePauseToggle .UIndex(X)
            End If
        Next X
    End With
End Sub

Private Sub PuedeDeath(ByVal userindex As Integer, ByRef lError As String)
    With UserList(userindex)
       
        
        If Death.Activo = False Then
            lError = "Evento inactivo"
            Exit Sub
        End If
        
        If Death.CuposRestantes <= 0 Then
            lError = "Cupos completos"
            Exit Sub
        End If
        
        If .UserDeath.EnDeath = True Then
            lError = "Ya estás en el evento"
            Exit Sub
        End If
        
        If (.flags.Muerto <> 0) Then
            lError = "Estás muerto"
            Exit Sub
        End If
        
        If (.Counters.Pena <> 0) Then
            lError = "Estás en la cárcel"
            Exit Sub
        End If
        
        If .Stats.ELV < 25 Then
            lError = "Necesitas ser nivel 25"
            Exit Sub
        End If
    
        If MapInfo(.Pos.map).Pk = True Then
            lError = "Estás en una zona insegura"
            Exit Sub
        End If
        
        If .EnEvento = True Then
            lError = "Ya estás en un evento"
            Exit Sub
        End If
        
        If .Stats.GLD < 200000 Then
            lError = "No tenes suficiente oro"
            Exit Sub
        End If
        
        
        
    End With
End Sub

Public Sub MensajeGlobal(ByVal Chat As String, ByVal FontIndex As FontTypeNames)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Chat, FontIndex))
End Sub
