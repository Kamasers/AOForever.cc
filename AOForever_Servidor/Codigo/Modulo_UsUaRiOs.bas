Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal victimIndex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 11/03/2010
'***************************************************

    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(victimIndex).Stats.ELV) * 2
    
    With UserList(attackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        If TriggerZonaPelea(victimIndex, attackerIndex) <> TRIGGER6_PERMITE Then

            EraCriminal = criminal(attackerIndex)
            
            With .Reputacion
                If Not criminal(victimIndex) Then
                    .AsesinoRep = .AsesinoRep + vlASESINO * 2
                    If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                    .BurguesRep = 0
                    .NobleRep = 0
                    .PlebeRep = 0
                Else
                    .NobleRep = .NobleRep + vlNoble
                    If .NobleRep > MAXREP Then .NobleRep = MAXREP
                End If
            End With
            
            If criminal(attackerIndex) Then
                If Not EraCriminal Then Call RefreshCharStatus(attackerIndex)
            Else
                If EraCriminal Then Call RefreshCharStatus(attackerIndex)
            End If
        End If
        
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
        'Call WriteConsoleMsg(VictimIndex, "¡" & .name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteMultiMessage(attackerIndex, eMessages.HaveKilledUser, victimIndex, DaExp)
        Call WriteMultiMessage(victimIndex, eMessages.UserKill, attackerIndex)
        
        'Call UserDie(VictimIndex)
        Call FlushBuffer(victimIndex)
        
        'Log
        Call LogAsesinato(.Name & " asesino a " & UserList(victimIndex).Name)
    End With
End Sub

Public Sub RevivirUsuario(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(userindex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        
        If .flags.Navegando = 1 Then
            Call ToogleBoatBody(userindex)
        Else
            Call DarCuerpoDesnudo(userindex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        If .flags.Traveling Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteMultiMessage(userindex, eMessages.CancelHome)
        End If
        
        Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(userindex)
    End With
End Sub

Public Sub ToogleBoatBody(ByVal userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 13/01/2010
'Gives boat body depending on user alignment.
'***************************************************

    Dim Ropaje As Integer
    
    With UserList(userindex)
        
 
        .Char.Head = 0
        
        ' Barco de armada
        'If .Faccion.ArmadaReal = 1 Then
            'Char.body = iFragataReal
            
        ' Barco de caos
        'ElseIf .Faccion.FuerzasCaos = 1 Then
            'Char.body = iFragataCaos
        
        'Barcos neutrales

            Select Case .Invent.BarcoObjIndex
                Case 474
                    .Char.body = 84
                
                Case 475
                    .Char.body = 85
                
                Case 476
                    .Char.body = 86
            End Select

        
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End With

End Sub

Public Sub ChangeUserChar(ByVal userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    With UserList(userindex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
    End With
End Sub

Public Function GetWeaponAnim(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 03/29/10
'
'***************************************************
    Dim tmp As Integer

    With UserList(userindex)
        tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = tmp
                Exit Function
            End If
        End If
        
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
    End With
End Function

Public Sub EnviarFama(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim L As Long
    
    With UserList(userindex).Reputacion
        L = (-.AsesinoRep) + _
            (-.BandidoRep) + _
            .BurguesRep + _
            (-.LadronesRep) + _
            .NobleRep + _
            .PlebeRep
        L = Round(L / 6)
        
        .Promedio = L
    End With
    
    Call WriteFame(userindex)
End Sub

Public Sub EraseUserChar(ByVal userindex As Integer, ByVal IsAdminInvisible As Boolean)
'*************************************************
'Author: Unknown
'Last modified: 08/01/2009
'08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
'*************************************************

On Error GoTo ErrorHandler
    
    With UserList(userindex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call EnviarDatosASlot(userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        End If
        
        Call QuitarUser(userindex, .Pos.map)
        
        MapData(.Pos.map, .Pos.X, .Pos.Y).userindex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Public Sub RefreshCharStatus(ByVal userindex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/07/2009
'Refreshes the status and tag of UserIndex.
'04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
'*************************************************
    Dim ClanTag As String
    Dim NickColor As Byte
    
    With UserList(userindex)
        If .GuildIndex > 0 Then
            ClanTag = modGuilds.GuildName(.GuildIndex)
            ClanTag = " <" & ClanTag & ">"
        End If
        
        NickColor = GetNickColor(userindex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageUpdateTagAndStatus(userindex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageUpdateTagAndStatus(userindex, NickColor, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Call ToogleBoatBody(userindex)
            End If
            
            Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
        'ustedes se preguntaran que hace esto aca?
      'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
       'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
      'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
       Dim NuevaA As Boolean

       Dim GI     As Integer

       Dim tStr   As String

       GI = .GuildIndex

       If GI > 0 Then
               NuevaA = False

               If Not modGuilds.m_ValidarPermanencia(userindex, True, NuevaA) Then
                       Call WriteConsoleMsg(userindex, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)
               End If

               If NuevaA Then
                       Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación " & modGuilds.GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
                       tStr = modGuilds.GuildName(GI)
                       Call LogClanes("¡El clan " & tStr & " cambio de alineación!")
               End If
  
       End If
    End With
End Sub

Public Function GetNickColor(ByVal userindex As Integer) As Byte
'*************************************************
'Author: ZaMa
'Last modified: 15/01/2010
'
'*************************************************
    
    With UserList(userindex)
        
        If criminal(userindex) Then
            GetNickColor = eNickColor.ieCriminal
        Else
            GetNickColor = eNickColor.ieCiudadano
        End If
    End With
    
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal userindex As Integer, _
        ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ButIndex As Boolean = False)
'*************************************************
'Author: Unknown
'Last modified: 15/01/2010
'23/07/2009: Budi - Ahora se envía el nick
'15/01/2010: ZaMa - Ahora se envia el color del nick.
'*************************************************

On Error GoTo Errhandler

    Dim CharIndex As Integer
    Dim ClanTag As String
    Dim NickColor As Byte
    Dim UserName As String
    Dim Privileges As Byte
    
    With UserList(userindex)
    
        If InMapBounds(map, X, Y) Then
            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = userindex
            End If
            
            'Place character on map if needed
            If toMap Then MapData(map, X, Y).userindex = userindex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)
                End If
                
                NickColor = GetNickColor(userindex)
                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .showName Then
                    UserName = .Name
                    
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else
                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then _
                                UserName = UserName & " <" & ClanTag & ">"
                        Else
                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else
                                If LenB(ClanTag) <> 0 Then _
                                    UserName = UserName & " <" & ClanTag & ">"
                            End If
                        End If
                    End If
                End If
            
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, _
                            .Char.CharIndex, X, Y, _
                            .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                            UserName, NickColor, Privileges)
            Else
                'Hide the name and clan - set privs as normal user
                 Call AgregarUser(userindex, .Pos.map, ButIndex)
            End If
        End If
    End With
Exit Sub

Errhandler:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(userindex)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/19/2009
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
'12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
'02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
'11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
'02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
'*************************************************
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    
On Error GoTo Errhandler
    
    WasNewbie = EsNewbie(userindex)
    
    With UserList(userindex)
        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea si alcanzó el máximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub
            End If
            
            'Store it!
            Call Statistics.UserLevelUp(userindex)
            
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
            Call WriteConsoleMsg(userindex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            
            If .Stats.ELV = 1 Then
                Pts = 10
            Else
                'For multiple levels being rised at once
                Pts = Pts + 5
            End If
            
            .Stats.ELV = .Stats.ELV + 1
            
            Call CheckRanking(eRankings.Nivel, userindex, .Stats.ELV)
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
                If .Stats.ELV = 2 Then
                .Stats.ELU = 450
                ElseIf .Stats.ELV = 3 Then
                .Stats.ELU = 675
                ElseIf .Stats.ELV = 4 Then
                .Stats.ELU = 1012
                ElseIf .Stats.ELV = 5 Then
                .Stats.ELU = 1518
                ElseIf .Stats.ELV = 6 Then
                .Stats.ELU = 2277
                ElseIf .Stats.ELV = 7 Then
                .Stats.ELU = 3416
                ElseIf .Stats.ELV = 8 Then
                .Stats.ELU = 5124
                ElseIf .Stats.ELV = 9 Then
                .Stats.ELU = 7886
                ElseIf .Stats.ELV = 10 Then
                .Stats.ELU = 11529
                ElseIf .Stats.ELV = 11 Then
                .Stats.ELU = 14988
                ElseIf .Stats.ELV = 12 Then
                .Stats.ELU = 19484
                ElseIf .Stats.ELV = 13 Then
                .Stats.ELU = 25329
                ElseIf .Stats.ELV = 14 Then
                .Stats.ELU = 32928
                ElseIf .Stats.ELV = 15 Then
                .Stats.ELU = 42806
                ElseIf .Stats.ELV = 16 Then
                .Stats.ELU = 55648
                ElseIf .Stats.ELV = 17 Then
                .Stats.ELU = 72342
                ElseIf .Stats.ELV = 18 Then
                .Stats.ELU = 94045
                ElseIf .Stats.ELV = 19 Then
                .Stats.ELU = 122259
                ElseIf .Stats.ELV = 20 Then
                .Stats.ELU = 158937
                ElseIf .Stats.ELV = 21 Then
                .Stats.ELU = 206618
                ElseIf .Stats.ELV = 22 Then
                .Stats.ELU = 268603
                ElseIf .Stats.ELV = 23 Then
                .Stats.ELU = 349184
                ElseIf .Stats.ELV = 24 Then
                .Stats.ELU = 453939
                ElseIf .Stats.ELV = 25 Then
                .Stats.ELU = 544727
                ElseIf .Stats.ELV = 26 Then
                .Stats.ELU = 667632
                ElseIf .Stats.ELV = 27 Then
                .Stats.ELU = 784406
                ElseIf .Stats.ELV = 28 Then
                .Stats.ELU = 941287
                ElseIf .Stats.ELV = 29 Then
                .Stats.ELU = 1129544
                ElseIf .Stats.ELV = 30 Then
                .Stats.ELU = 1355453
                ElseIf .Stats.ELV = 31 Then
                .Stats.ELU = 1626544
                ElseIf .Stats.ELV = 32 Then
                .Stats.ELU = 1951853
                ElseIf .Stats.ELV = 33 Then
                .Stats.ELU = 2342224
                ElseIf .Stats.ELV = 34 Then
                .Stats.ELU = 3372803
                ElseIf .Stats.ELV = 35 Then
                .Stats.ELU = 4047364
                ElseIf .Stats.ELV = 36 Then
                .Stats.ELU = 5828204
                ElseIf .Stats.ELV = 37 Then
                .Stats.ELU = 6993845
                ElseIf .Stats.ELV = 38 Then
                .Stats.ELU = 8392614
                ElseIf .Stats.ELV = 39 Then
                .Stats.ELU = 10071137
                ElseIf .Stats.ELV = 40 Then
                .Stats.ELU = 120853640
                ElseIf .Stats.ELV = 41 Then
                .Stats.ELU = 145024370
                ElseIf .Stats.ELV = 42 Then
                .Stats.ELU = 174029240
                ElseIf .Stats.ELV = 43 Then
                .Stats.ELU = 208835090
                ElseIf .Stats.ELV = 44 Then
                .Stats.ELU = 417670180
                ElseIf .Stats.ELV = 45 Then
                .Stats.ELU = 835340360
                ElseIf .Stats.ELV = 46 Then
                .Stats.ELU = 1670680720
                Else
                .Stats.ELU = 0
                End If
            
            'Calculo subida de vida
            Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
        
            AumentoHP = Promedio + gethp(Promedio)
        
            Select Case .clase
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Pirat
                    AumentoHIT = 3
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Thief
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Worker
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTTrabajador
                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case eClass.Bandit
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                    AumentoSTA = AumentoStBandido
                
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHp = .Stats.MaxHp + AumentoHP
            If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MaxHIT = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                    .Stats.MinHIT = STAT_MAXHIT_OVER36
            End If
            
            'Notificamos al user
            If AumentoHP > 0 Then
                Call WriteConsoleMsg(userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoSTA > 0 Then
                Call WriteConsoleMsg(userindex, "Has ganado " & AumentoSTA & " puntos de energía.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoMANA > 0 Then
                Call WriteConsoleMsg(userindex, "Has ganado " & AumentoMANA & " puntos de maná.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoHIT > 0 Then
                Call WriteConsoleMsg(userindex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(userindex, "Tu golpe mínimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            
            .Stats.MinHp = .Stats.MaxHp
            
            If .Stats.ELV = 25 Then
                GI = .GuildIndex
                If GI > 0 Then
                    If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                        'We get here, so guild has factionary alignment, we have to expulse the user
                        Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                        Call WriteConsoleMsg(userindex, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
            End If

        Loop
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(userindex) And WasNewbie Then
            Call QuitarNewbieObj(userindex)
            If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
                Call WarpUserChar(userindex, 1, 50, 50, True)
                Call WriteConsoleMsg(userindex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'Send all gained skill points at once (if any)
        If Pts > 0 Then
            Call WriteLevelUp(userindex, Pts)
            
            .Stats.SkillPts = .Stats.SkillPts + Pts
            
            Call WriteConsoleMsg(userindex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
        End If
        
    End With
    
    Call WriteUpdateUserStats(userindex)
Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub
Public Function gethp(ByVal Promedio As Double) As Double
    Dim X As Integer
    X = RandomNumber(1, 1000)
    If Int(Promedio) <> Promedio Then
        Select Case X
            Case 1 To 250
                gethp = 0.5
            Case 251 To 500
                gethp = -0.5
            Case 501 To 750
                gethp = 0.5
            Case 751 To 1000
                gethp = -0.5
        End Select
    Else
        Select Case X
            Case 1 To 250
                gethp = 1
            Case 251 To 500
                gethp = -1
            Case 501 To 750
                gethp = 1
            Case 751 To 1000
                gethp = -1
        End Select
    End If
End Function
Public Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PuedeAtravesarAgua = UserList(userindex).flags.Navegando = 1 _
                    Or UserList(userindex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As eHeading)
'*************************************************
'Author: Unknown
'Last modified: 13/07/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
'13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
'13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim CasPerPos As WorldPos
    Dim HayUser As Integer
    
    sailing = PuedeAtravesarAgua(userindex)
    nPos = UserList(userindex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    If MoveToLegalPos(UserList(userindex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        
        If MapInfo(UserList(userindex).Pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(userindex).Pos.map, nPos.X, nPos.Y).userindex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            'Call WriteConsoleMsg(UserIndex, "COLISIONA", FontTypeNames.FONTTYPE_CITIZEN)
            
            HayUser = MapData(UserList(userindex).Pos.map, nPos.X + 1, nPos.Y).userindex
            
            If HayUser > 0 Then
                ''Call WriteConsoleMsg(UserIndex, "COLISIONA", FontTypeNames.FONTTYPE_CITIZEN)
            End If
            
            If CasperIndex > 0 Then
            
                ' Los admins invisibles no pueden patear caspers
                If Not (UserList(userindex).flags.AdminInvisible = 1) Then
                    
                    If TriggerZonaPelea(userindex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
                        End If
                    End If
    
                    CasperHeading = InvertHeading(nHeading)
                    CasPerPos = UserList(CasperIndex).Pos
                    Call HeadtoPos(CasperHeading, CasPerPos)
    
                    With UserList(CasperIndex)
                        
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then _
                            Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, CasPerPos.X, CasPerPos.Y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and user pos
                        .Pos = CasPerPos
                        .Char.heading = CasperHeading
                        MapData(.Pos.map, CasPerPos.X, CasPerPos.Y).userindex = CasperIndex
                    
                    End With
                
                    'Actualizamos las áreas de ser necesario
                    Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If

            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not UserList(userindex).flags.AdminInvisible = 1 Then _
                Call SendData(SendTarget.ToPCAreaButIndex, userindex, PrepareMessageCharacterMove(UserList(userindex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not ((UserList(userindex).flags.AdminInvisible = 1) And CasperIndex <> 0) Then
            Dim oldUserIndex As Integer
            
            oldUserIndex = MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex
            
            ' Si no hay intercambio de pos con nadie
            If oldUserIndex = userindex Then
                MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = 0
            End If
            
            UserList(userindex).Pos = nPos
            UserList(userindex).Char.heading = nHeading
            MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).userindex = userindex
            Call DoTileEvents(userindex, UserList(userindex).Pos.map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y)
            
            'Actualizamos las áreas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)
        Else
            Call WritePosUpdate(userindex)
        End If

    Else
        Call WritePosUpdate(userindex)
    End If
    
    If UserList(userindex).Counters.Trabajando Then _
        UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

    If UserList(userindex).Counters.Ocultando Then _
        UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function

Sub ChangeUserInv(ByVal userindex As Integer, ByVal slot As Byte, ByRef Object As UserOBJ)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(userindex).Invent.Object(slot) = Object
    Call WriteChangeInventorySlot(userindex, slot)
End Sub

Function NextOpenCharIndex() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim loopC As Long
    
    For loopC = 1 To MAXCHARS
        If CharList(loopC) = 0 Then
            NextOpenCharIndex = loopC
            NumChars = NumChars + 1
            
            If loopC > LastChar Then _
                LastChar = loopC
            
            Exit Function
        End If
    Next loopC
End Function

Function NextOpenUser() As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim loopC As Long
    
    For loopC = 1 To MaxUsers + 1
        If loopC > MaxUsers Then Exit For
        If (UserList(loopC).ConnID = -1 And UserList(loopC).flags.UserLogged = False) Then Exit For
    Next loopC
    
    NextOpenUser = loopC
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GuildI As Integer
    
    With UserList(userindex)
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(sendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
#If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim tempStr As String
        TempDate = Now - .LogOnTime
        TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        tempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.GLD & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    With UserList(userindex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.ArmadaReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        
        ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
        
        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
    Dim CharFile As String
    Dim Ban As String
    Dim BanDetailPath As String
    
    BanDetailPath = App.path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejército real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legión oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        End If

        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    
    With UserList(userindex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To .CurrentInventorySlots
            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    End With
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim j As Long
    Dim CharFile As String, tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, tmp, Asc("-"))
            ObjCant = ReadField(2, tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(userindex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(userindex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(userindex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal userindex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 02/04/2010
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserList(userindex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(userindex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(userindex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(userindex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(userindex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> userindex Then
            Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    
    If EsMascotaCiudadano(NpcIndex, userindex) Then
        Call VolverCriminal(userindex)
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    Else
        EraCriminal = criminal(userindex)
        
        'Reputacion
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
           If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call VolverCriminal(userindex)
           End If
        
        ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
           UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR / 2
           If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
            UserList(userindex).Reputacion.PlebeRep = MAXREP
        End If
        
        If Npclist(NpcIndex).MaestroUser <> userindex Then
            'hacemos que el npc se defienda
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
        End If
        
        If EraCriminal And Not criminal(userindex) Then
            Call VolverCiudadano(userindex)
        End If
    End If
End Sub
Public Function PuedeApuñalar(ByVal userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(userindex).clase = eClass.Assasin
        End If
    End If
End Function

Public Function PuedeAcuchillar(ByVal userindex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 25/01/2010 (ZaMa)
'
'***************************************************
    
    With UserList(userindex)
        If .clase = eClass.Pirat Then
            If .Invent.WeaponEqpObjIndex > 0 Then
                PuedeAcuchillar = (ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1)
            End If
        End If
    End With
    
End Function

Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer, ByVal Acerto As Boolean)

    With UserList(userindex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            
            If .Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
            
            Dim Lvl As Integer
            Lvl = .Stats.ELV
            
            If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
            If .Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
            

                .Stats.UserSkills(Skill) = .Stats.UserSkills(Skill) + 1
                Call WriteConsoleMsg(userindex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & .Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                
                .Stats.Exp = .Stats.Exp + 50
                If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                
                Call WriteConsoleMsg(userindex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
                
                Call WriteUpdateExp(userindex)
                Call CheckUserLevel(userindex)
        End If
    End With
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal userindex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 12/01/2010 (ZaMa)
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
'27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
'21/07/2009: Marco - Al morir se desactiva el comercio seguro.
'16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
'27/11/2009: Budi - Al morir envia los atributos originales.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    With UserList(userindex)
        'Sonido
        If .Genero = eGenero.Mujer Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, e_SoundIndex.MUERTE_MUJER)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, e_SoundIndex.MUERTE_HOMBRE)
        End If
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        If Not .GuildIndex = 0 Or Not .GuildIndex > CANTIDADDECLANES Then
            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg(.Name & " ha muerto en el MAPA " & .Pos.map & " X: " & .Pos.X & " Y: " & .Pos.Y, FontTypeNames.FONTTYPE_VENENO))
        End If
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        ' No se activa en arenas
        If TriggerZonaPelea(userindex, userindex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(userindex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
        End If
        
        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
        End If
        
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        Call PerdioNpc(userindex)

        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(userindex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(userindex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(userindex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(userindex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call SetInvisible(userindex, UserList(userindex).Char.CharIndex, False)
        End If
        
        If TriggerZonaPelea(userindex, userindex) <> eTrigger6.TRIGGER6_PERMITE Then
            ' << Si es newbie no pierde el inventario >>
            If Not EsNewbie(userindex) Then
                If Not EsGM(userindex) Then Call TirarTodo(userindex)
            Else
                Call TirarTodosLosItemsNoNewbies(userindex)
            End If
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(userindex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot2 > 0 Then
            Call Desequipar(userindex, .Invent.AnilloEqpSlot2)
        End If
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.EscudoEqpSlot)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If
        
        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            If criminal(userindex) Then
                .Char.body = 145
                .Char.Head = 501
            Else
                .Char.body = 8
                .Char.Head = 500
            End If
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            
            .Char.body = 87
        End If
        
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
            ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0
            End If
        Next i
        
        .NroMascotas = 0
        
        Call mod_Retos3vs3.Death(userindex)
        Call mod_DeathMatch.Muere_Death(userindex)
        Call Muere_HungerGames(userindex)
        Call eventDie(userindex)
        '<< Actualizamos clientes >>
        Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(userindex)
        Call WriteUpdateStrenghtAndDexterity(userindex)
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(userindex)
        Call MuereReto(userindex)
        Call Event_1vs1_Aim_Melee.Death_Event(userindex)
        If UserList(userindex).Torneo.EnTorneo = True Then
            Call proccessDeathOrDisconnect(userindex)
        End If
    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If EsNewbie(Muerto) Then Exit Sub
    
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If criminal(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name
                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                    .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
            End If
            
            If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
                .Faccion.Reenlistadas = 200  'jaja que trucho
                
                'con esto evitamos que se vuelva a reenlistar
            End If
        Else
            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name
                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                    .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            End If
        End If
        
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
            .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
            
        Call CheckRanking(eRankings.Matados, Atacante, .Stats.UsuariosMatados)
    End With
    
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim loopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.map = Pos.map
    nPos.X = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If loopC > 15 Then
            Exit Do
        End If
        
        For tY = Pos.Y - loopC To Pos.Y + loopC
            For tX = Pos.X - loopC To Pos.X + loopC
                
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.X + loopC
                        tY = Pos.Y + loopC
                    End If
                End If
            
            Next tX
        Next tY
        
        loopC = loopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal FX As Boolean, Optional ByVal Teletransported As Boolean, Optional ByVal StablePos As Boolean = True)
'**************************************************************
'Author: Unknown
'Last Modify Date: 13/11/2009
'15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
'13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
'**************************************************************
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    If Not EsGM(userindex) Then
    
        ''If StablePos Then
            Dim nPos As WorldPos
            Dim oPos As WorldPos
            oPos.map = map
            oPos.X = X
            oPos.Y = Y
            Call ClosestStablePos1(oPos, nPos)
            map = nPos.map
            X = nPos.X
            Y = nPos.Y
        ''End If
    End If
    With UserList(userindex)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        Call WriteRemoveAllDialogs(userindex)
        
        OldMap = .Pos.map
        OldX = .Pos.X
        OldY = .Pos.Y

        Call EraseUserChar(userindex, .flags.AdminInvisible = 1)
        
        If OldMap <> map Then
            Call WriteChangeMap(userindex, map, MapInfo(.Pos.map).MapVersion, MapInfo(map).Name)
            Call WritePlayMidi(userindex, val(ReadField(1, MapInfo(map).Music, 45)))
            
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean
            nextMap = IIf(distanceToCities(map).distanceToCity(.Hogar) >= 0, True, False)
            previousMap = IIf(distanceToCities(.Pos.map).distanceToCity(.Hogar) >= 0, True, False)

            If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                .flags.lastMap = .Pos.map
            ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
                .flags.lastMap = 0
            ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                .flags.lastMap = .flags.lastMap
            End If
        
        End If
        
        .Pos.X = X
        .Pos.Y = Y
        .Pos.map = map
        
        Call MakeUserChar(True, map, userindex, map, X, Y)
        Call WriteUserCharIndexInServer(userindex)
        Call DoTileEvents(userindex, map, X, Y)
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(userindex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SetInvisible(userindex, .Char.CharIndex, True)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        If Teletransported Then
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                Call WriteMultiMessage(userindex, eMessages.CancelHome)
            End If
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(userindex)
        
        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(userindex, True)
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(userindex)
        
        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(userindex)
                End If
            Else
                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(userindex)
                End If
            End If
        End If
      
    End With
End Sub

Private Sub WarpMascotas(ByVal userindex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 11/05/2009
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
'************************************************
    Dim i As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp As Boolean
    Dim index As Integer
    Dim iMinHP As Integer
    
    NroPets = UserList(userindex).NroMascotas
    canWarp = (MapInfo(UserList(userindex).Pos.map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(userindex).MascotasIndex(i)
        
        If index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(userindex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHp
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(userindex).MascotasType(i) = petType

            End If
        ElseIf UserList(userindex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(userindex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And canWarp Then
            index = SpawnNpc(petType, UserList(userindex).Pos, False, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If index = 0 Then
                Call WriteConsoleMsg(userindex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(userindex).MascotasIndex(i) = index

                ' Nos aseguramos de que conserve el hp, si estaba dañado
                Npclist(index).Stats.MinHp = IIf(iMinHP = 0, Npclist(index).Stats.MinHp, iMinHP)
            
                Npclist(index).MaestroUser = userindex
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(userindex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(userindex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(userindex).NroMascotas = NroPets
End Sub

Public Sub WarpMascota(ByVal userindex As Integer, ByVal PetIndex As Integer)
'************************************************
'Author: ZaMa
'Last Modified: 18/11/2009
'Warps a pet without changing its stats
'************************************************
    Dim petType As Integer
    Dim NpcIndex As Integer
    Dim iMinHP As Integer
    Dim TargetPos As WorldPos
    
    With UserList(userindex)
        
        TargetPos.map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
        
        NpcIndex = .MascotasIndex(PetIndex)
            
        'Store data and remove NPC to recreate it after warp
        petType = .MascotasType(PetIndex)
        
        ' Guardamos el hp, para restaurarlo cuando se cree el npc
        iMinHP = Npclist(NpcIndex).Stats.MinHp
        
        Call QuitarNPC(NpcIndex)
        
        ' Restauramos el valor de la variable
        .MascotasType(PetIndex) = petType
        .NroMascotas = .NroMascotas + 1
        NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        
        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
        ' Exception: Pets don't spawn in water if they can't swim
        If NpcIndex = 0 Then
            Call WriteConsoleMsg(userindex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
        Else
            .MascotasIndex(PetIndex) = NpcIndex

            With Npclist(NpcIndex)
                ' Nos aseguramos de que conserve el hp, si estaba dañado
                .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
            
                .MaestroUser = userindex
                .Movement = TipoAI.SigueAmo
                .Target = 0
                .TargetNPC = 0
            End With
            
            Call FollowAmo(NpcIndex)
        End If
    End With
End Sub


''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************
    Dim isNotVisible As Boolean
    Dim HiddenPirat As Boolean
    
    With UserList(userindex)
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
            
            isNotVisible = (.flags.Oculto Or .flags.invisible)
            If isNotVisible Then
                .flags.invisible = 0
                
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .clase = eClass.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToogleBoatBody(userindex)
                            Call WriteConsoleMsg(userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                                NingunEscudo, NingunCasco)
                            HiddenPirat = True
                        End If
                    End If
                End If
                
                .flags.Oculto = 0
                
                ' Para no repetir mensajes
                If Not HiddenPirat Then Call WriteConsoleMsg(userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                Call SetInvisible(userindex, .Char.CharIndex, False)

            End If
            
            If .flags.Traveling = 1 Then
                Call WriteMultiMessage(userindex, eMessages.CancelHome)
            End If
            
            Call WriteConsoleMsg(userindex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(userindex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(userindex).ConnIDValida Then
            UserList(userindex).Counters.Saliendo = False
            UserList(userindex).Counters.Salir = 0
            Call WriteConsoleMsg(userindex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(userindex).Counters.Salir = IIf((UserList(userindex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(userindex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ViejoNick As String
    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
    End If
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
#If ConUpTime Then
        Dim TempSecs As Long
        Dim tempStr As String
        TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
        tempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
    
    End If
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim CharFile As String
    
On Error Resume Next
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub VolverCriminal(ByVal userindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/02/2010
'Nacho: Actualiza el tag al cliente
'21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
'**************************************************************
    With UserList(userindex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)

        End If
    End With
    
    Call RefreshCharStatus(userindex)
End Sub

Sub VolverCiudadano(ByVal userindex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 21/06/2006
'Nacho: Actualiza el tag al cliente.
'**************************************************************
    With UserList(userindex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
    End With
    
    Call RefreshCharStatus(userindex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function

Public Sub SetInvisible(ByVal userindex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim sndNick As String

With UserList(userindex)
    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, userindex, PrepareMessageSetInvisible(userCharIndex, invisible))
    
    sndNick = .Name
    
    If invisible Then
        sndNick = sndNick & " " & TAG_USER_INVISIBLE
    Else
        If .GuildIndex > 0 Then
            sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
        End If
    End If
    
    Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, userindex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
End With
End Sub

Public Sub SetConsulatMode(ByVal userindex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/06/10
'
'***************************************************

Dim sndNick As String

With UserList(userindex)
    sndNick = .Name
    
    If .flags.EnConsulta Then
        sndNick = sndNick & " " & TAG_CONSULT_MODE
    Else
        If .GuildIndex > 0 Then
            sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
End With
End Sub

Public Function IsArena(ByVal userindex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 10/11/2009
'Returns true if the user is in an Arena
'**************************************************************
    IsArena = (TriggerZonaPelea(userindex, userindex) = TRIGGER6_PERMITE)
End Function

Public Sub PerdioNpc(ByVal userindex As Integer)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/01/2010 (ZaMa)
'The user loses his owned npc
'18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
'**************************************************************

    Dim PetIndex As Long
    
    With UserList(userindex)
        If .flags.OwnedNpc > 0 Then
            Npclist(.flags.OwnedNpc).Owner = 0
            .flags.OwnedNpc = 0
            
            ' Dejan de atacar las mascotas
            If .NroMascotas > 0 Then
                For PetIndex = 1 To MAXMASCOTAS
                    If .MascotasType(PetIndex) > 0 Then Call FollowAmo(PetIndex)
                Next PetIndex
            End If
        End If
    End With
End Sub

Public Sub ApropioNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/01/2010 (zaMa)
'The user owns a new npc
'18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
'19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
'**************************************************************

    With UserList(userindex)
        ' Los admins no se pueden apropiar de npcs
        If EsGM(userindex) Then Exit Sub
        
        'No aplica a zonas seguras
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(.Pos.map).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueño, lo perdio aca
        Npclist(NpcIndex).Owner = userindex
        .flags.OwnedNpc = NpcIndex
    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(userindex, True)
End Sub

Public Function GetDireccion(ByVal userindex As Integer, ByVal OtherUserIndex As Integer) As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve la direccion hacia donde esta el usuario
'**************************************************************
    Dim X As Integer
    Dim Y As Integer
    
    X = UserList(userindex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(userindex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"
    End If

End Function

Public Function SameFaccion(ByVal userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 17/11/2009
'Devuelve True si son de la misma faccion
'**************************************************************
    SameFaccion = (esCaos(userindex) And esCaos(OtherUserIndex)) Or _
                    (esArmada(userindex) And esArmada(OtherUserIndex))
End Function

Public Function FarthestPet(ByVal userindex As Integer) As Integer
'**************************************************************
'Author: ZaMa
'Last Modify Date: 18/11/2009
'Devuelve el indice de la mascota mas lejana.
'**************************************************************
On Error GoTo Errhandler
    
    Dim PetIndex As Integer
    Dim Distancia As Integer
    Dim OtraDistancia As Integer
    
    With UserList(userindex)
        If .NroMascotas = 0 Then Exit Function
    
        For PetIndex = 1 To MAXMASCOTAS
            ' Solo pos invocar criaturas que exitan!
            If .MascotasIndex(PetIndex) > 0 Then
                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                    Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                        Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex
                        End If
                    End If
                End If
            End If
        Next PetIndex
    End With

    Exit Function
    
Errhandler:
    Call LogError("Error en FarthestPet")
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal userindex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 11/20/2009
'
'*************************************************

With UserList(userindex).Stats
    If .UserSkills(Skill) < MAXSKILLPOINTS Then
        If Allocation Then
            .ExpSkills(Skill) = 0
        Else
            .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
        End If
        
        .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
    Else
        .ExpSkills(Skill) = 0
        .EluSkills(Skill) = 0
    End If
End With

End Sub

Public Function HasEnoughItems(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks Wether the user has the required amount of items in the inventory or not
'**************************************************************

    Dim slot As Long
    Dim ItemInvAmount As Long
    
    For slot = 1 To UserList(userindex).CurrentInventorySlots
        ' Si es el item que busco
        If UserList(userindex).Invent.Object(slot).ObjIndex = ObjIndex Then
            ' Lo sumo a la cantidad total
            ItemInvAmount = ItemInvAmount + UserList(userindex).Invent.Object(slot).Amount
        End If
    Next slot

    HasEnoughItems = Amount <= ItemInvAmount
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, ByVal userindex As Integer) As Long
'**************************************************************
'Author: ZaMa
'Last Modify Date: 25/11/2009
'Cheks the amount of items the user has in offerSlots.
'**************************************************************
    Dim slot As Byte
    
    For slot = 1 To MAX_OFFER_SLOTS
            ' Si es el item que busco
        If UserList(userindex).ComUsu.Objeto(slot) = ObjIndex Then
            ' Lo sumo a la cantidad total
            TotalOfferItems = TotalOfferItems + UserList(userindex).ComUsu.cant(slot)
        End If
    Next slot

End Function

Public Function getMaxInventorySlots(ByVal userindex As Integer) As Byte
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If UserList(userindex).Invent.MochilaEqpObjIndex > 0 Then
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(userindex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
Else
    getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
End If
End Function

Public Sub goHome(ByVal userindex As Integer)
Dim Distance As Integer
Dim tiempo As Long

With UserList(userindex)
    If .flags.Muerto = 1 Then
        If .flags.lastMap = 0 Then
            Distance = distanceToCities(.Pos.map).distanceToCity(.Hogar)
        Else
            Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
        End If
        
        tiempo = (Distance + 1) * 30 'segundos
        
        .Counters.goHome = tiempo / 6 'Se va a chequear cada 6 segundos.
        
        .flags.Traveling = 1

        Call WriteMultiMessage(userindex, eMessages.Home, Distance, tiempo, , MapInfo(Ciudades(.Hogar).map).Name)
    Else
        Call WriteConsoleMsg(userindex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End With
End Sub


Public Sub setHome(ByVal userindex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 30/04/2010
'30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
'***************************************************
    If newHome < eCiudad.cUllathorpe Or newHome > cArghal Then Exit Sub
    UserList(userindex).Hogar = newHome
    
    Call WriteChatOverHead(userindex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
End Sub
