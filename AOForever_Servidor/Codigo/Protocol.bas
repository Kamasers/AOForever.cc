Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

Public OnlineString As String
Public OnlineNum As Integer
''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue


Private Enum ServerPacketID
    UpdateDiam
    ListaMercado
    CerrarPartyForm
    ListarViajes
    CreateProjectile
    CreateDamage            ' CDMG
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMidi                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    
    ShowGuildAlign
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    SendAura
    PartyForm
    MovimientSW
    closeclient
    SacarScreen
    EnviarScreen
    SendRanking
End Enum

Private Enum ClientPacketID
    BorrarPj
    RequestRanking
    PedirScreen
    Mandarscreen
    ConsultaGM
    SolicitudCambioPj
    DatosCambioPj
    ConfirmarCambioPJ
    RecuperarPass
    Viajar
    EcharDisolverParty
    SavePartyPorc
    requestpartyform
    CrearParty
    AceptarParty
    SolicitudParty
    
    DragToPos
    DragInventario
    DepositarTodo
    RetirarTodo
    DisolverClan
    ReanudarClan
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    ReleasePet              '/LIBERAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    GuildFundation
    Ping                    '/PING
    ItemUpgrade
    GMCommands
    InitCrafting
    Home
    ShowGuildNews
    ShareNpc                '/COMPARTIR
    StopSharingNpc
    Consultation
    Retar
    AceptarReto
    AbandonarReto
    PublicarUsr
    ComprarUsr
    CancelarUsr
    RequestListaMercado
    sendReto
    acceptReto
    
    Send_Reto2
    Accept_Reto2
    Refuse_Reto2
    Send_Reto3
    Accept_Reto3
    Refuse_Reto3
    Cheat
    Do_Death
    Enter_Death
    Cancel_Death
    Do_Aim
    Enter_Aim
    DoJDH
    Enter_JDH
    Cancel_JDH
    DarDiamantes
    Canjear
    CrearTorneo
    EntrarTorneo
    CancelarTorneo
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 128

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
End Enum



''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(UserIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.ThrowDices _
      Or packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar _
      Or packetID = ClientPacketID.RecuperarPass _
      Or packetID = ClientPacketID.BorrarPj) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
    Select Case packetID
        Case ClientPacketID.BorrarPj
            Call HandleBorrarPj(UserIndex)
            
        Case ClientPacketID.RequestRanking
            Call HandleRequestRanking(UserIndex)
            
        Case ClientPacketID.PedirScreen
            Call HandlePedirScreen(UserIndex)
            
        Case ClientPacketID.Mandarscreen
            Call HandleMandarScreen(UserIndex)
            
        Case ClientPacketID.Cheat
            Call HandleCheater(UserIndex)
            
        Case ClientPacketID.ConsultaGM
            Call HandleConsultaGM(UserIndex)
            
        Case ClientPacketID.RecuperarPass
            Call HandleRecuperarPass(UserIndex)
            
        Case ClientPacketID.EcharDisolverParty
            Call HandleEcharDisolverParty(UserIndex)
            
        Case ClientPacketID.SolicitudParty
            Call HandleSolicitudParty(UserIndex)
            
        Case ClientPacketID.CrearParty
            Call HandleCrearParty(UserIndex)

        Case ClientPacketID.AceptarParty
            Call HandleAceptarParty(UserIndex)
            
        Case ClientPacketID.requestpartyform
            Call HandleRequestPartyForm(UserIndex)
            
        Case ClientPacketID.SavePartyPorc
            Call HandleSavePartyPerc(UserIndex)
            
        Case ClientPacketID.DragInventario
            Call mod_DragDrop.HandleDragInventory(UserIndex)
           
        Case ClientPacketID.Viajar
            Call HandleViajar(UserIndex)
        
        Case ClientPacketID.DragToPos
            Call mod_DragDrop.HandleDragToPos(UserIndex)
        
        Case ClientPacketID.DepositarTodo
            Call handleDepositarTodo(UserIndex)
            
        Case ClientPacketID.RetirarTodo
            Call handleRetirarTodo(UserIndex)
            
        Case ClientPacketID.DisolverClan
            Call HandleDisolverClan(UserIndex)
           
        Case ClientPacketID.ReanudarClan
            Call HandleReanudarClan(UserIndex)
            
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.ThrowDices              'TIRDAD
            Call HandleThrowDices(UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
            
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)

        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASEÑA
            Call HandleChangePassword(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(UserIndex)
            
        Case ClientPacketID.GuildFundation
            Call HandleGuildFundation(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
            
        Case ClientPacketID.ItemUpgrade
            Call HandleItemUpgrade(UserIndex)
        
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(UserIndex)
        
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
        
        Case ClientPacketID.ShowGuildNews
            Call HandleShowGuildNews(UserIndex)
            
        Case ClientPacketID.ShareNpc
            Call HandleShareNpc(UserIndex)
            
        Case ClientPacketID.StopSharingNpc
            Call HandleStopSharingNpc(UserIndex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(UserIndex)
            
        Case ClientPacketID.sendReto
            Call handleSendReto(UserIndex)
           
        Case ClientPacketID.acceptReto
            Call handleAcceptReto(UserIndex)
        
        Case ClientPacketID.Retar
            Call HandleRetar(UserIndex)
        
        Case ClientPacketID.AceptarReto
            Call HandleAceptarReto(UserIndex)
            
        Case ClientPacketID.AbandonarReto
            Call HandleAbandonarReto(UserIndex)
            
        Case ClientPacketID.Send_Reto2
           ' Call HandleSendReto2(UserIndex)
            
        Case ClientPacketID.Accept_Reto2
          '  Call HandleAcceptReto2(UserIndex)
            
        Case ClientPacketID.Refuse_Reto2
           ' Call handleRefuseReto2(UserIndex)
            
        Case ClientPacketID.Send_Reto3
            Call HandleSendReto3(UserIndex)
            
        Case ClientPacketID.Accept_Reto3
            Call HandleAcceptReto3(UserIndex)
            
        Case ClientPacketID.Refuse_Reto3
            Call handleRefuseReto3(UserIndex)
            
        Case ClientPacketID.PublicarUsr
            Call HandlePublicarUser(UserIndex)
            
        Case ClientPacketID.ComprarUsr
            Call handleComprarUser(UserIndex)
        
        Case ClientPacketID.CancelarUsr
            Call handleCancelarUser(UserIndex)
        
        Case ClientPacketID.RequestListaMercado
            Call HandleRequestMercado(UserIndex)
            
        Case ClientPacketID.Do_Death
            Call HandleDoDeath(UserIndex)
        
        Case ClientPacketID.Enter_Death
            Call HandleEnterDeath(UserIndex)
            
        Case ClientPacketID.Cancel_Death
            Call HandleCancelDeath(UserIndex)
            
        Case ClientPacketID.DoJDH
            Call HandleDoJDH(UserIndex)
        
        Case ClientPacketID.Enter_JDH
            Call HandleEnterJDH(UserIndex)
            
        Case ClientPacketID.Cancel_JDH
            Call HandleCancelJDH(UserIndex)
            
        Case ClientPacketID.Enter_Aim
            Call HandleEnterAIM(UserIndex)
            
        Case ClientPacketID.Do_Aim
            Call HandleDoAIM(UserIndex)
            
        Case ClientPacketID.CrearTorneo
            Call HandleCrearTorneo(UserIndex)
            
        Case ClientPacketID.EntrarTorneo
            Call HandleEntrarTorneo(UserIndex)
            
        Case ClientPacketID.CancelarTorneo
            Call HandleCancelarTorneo(UserIndex)
            
            
#If SeguridadAlkon Then
        Case Else
            Call HandleIncomingDataEx(UserIndex)
#Else
        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
#End If
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(UserIndex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
    End If
End Sub

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex
            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, _
                eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, _
                eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1) 'Target
                Call .WriteInteger(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteLong(Arg2) 'Expe
            
            Case eMessages.UserKill '"¡" & .name & " te ha matado!"
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda así.
                Call .WriteASCIIString(StringArg1) 'Call .WriteByte(CByte(Arg2))
                
        End Select
    End With
Exit Sub ''

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

Dim Command As Byte

With UserList(UserIndex)
    Call .incomingData.ReadByte
    
    Command = .incomingData.PeekByte
    
    Select Case Command
        Case eGMCommands.GMMessage                '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case eGMCommands.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case eGMCommands.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case eGMCommands.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
        
        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
        
        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
        
        Case eGMCommands.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)
        
        Case eGMCommands.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)
        
        Case eGMCommands.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)
        
        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
        
        Case eGMCommands.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case eGMCommands.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case eGMCommands.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
        
        Case eGMCommands.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case eGMCommands.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
        
        Case eGMCommands.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
        
        Case eGMCommands.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
        
        Case eGMCommands.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case eGMCommands.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case eGMCommands.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
        
        Case eGMCommands.Forgive                 '/PERDON
            Call HandleForgive(UserIndex)
        
        Case eGMCommands.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
        
        Case eGMCommands.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
        
        Case eGMCommands.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
        
        Case eGMCommands.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
        
        Case eGMCommands.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
        
        Case eGMCommands.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
        
        Case eGMCommands.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
        
        Case eGMCommands.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
        
        Case eGMCommands.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
        
        Case eGMCommands.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
        
        Case eGMCommands.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
        
        Case eGMCommands.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case eGMCommands.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
        
        Case eGMCommands.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(UserIndex)
        
        Case eGMCommands.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
        
        Case eGMCommands.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
        
        Case eGMCommands.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)
        
        Case eGMCommands.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
        
        Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
        
        Case eGMCommands.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
        
        Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
        
        Case eGMCommands.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)
        
        Case eGMCommands.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)
        
        Case eGMCommands.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
        
        Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        
        Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
        
        Case eGMCommands.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
        
        Case eGMCommands.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
        
        Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
        
        Case eGMCommands.DumpIPTables            '/DUMPSECURITY
            Call HandleDumpIPTables(UserIndex)
        
        Case eGMCommands.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case eGMCommands.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case eGMCommands.AskTrigger              '/TRIGGER with no args
            Call HandleAskTrigger(UserIndex)
        
        Case eGMCommands.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case eGMCommands.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(UserIndex)
        
        Case eGMCommands.GuildBan                '/BANCLAN
            Call HandleGuildBan(UserIndex)
        
        Case eGMCommands.BANIP                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case eGMCommands.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case eGMCommands.CreateItem              '/CI
            Call HandleCreateItem(UserIndex)
        
        Case eGMCommands.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
        
        Case eGMCommands.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)
        
        Case eGMCommands.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)
        
        Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case eGMCommands.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case eGMCommands.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case eGMCommands.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case eGMCommands.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case eGMCommands.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case eGMCommands.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case eGMCommands.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case eGMCommands.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case eGMCommands.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case eGMCommands.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case eGMCommands.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(UserIndex)
        
        Case eGMCommands.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(UserIndex)
        
        Case eGMCommands.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case eGMCommands.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(UserIndex)
        
        Case eGMCommands.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(UserIndex)
        
        Case eGMCommands.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
        
        Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(UserIndex)
        
        Case eGMCommands.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case eGMCommands.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)
        
        Case eGMCommands.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)
        
        Case eGMCommands.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case eGMCommands.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)
        
        Case Declaraciones.eGMCommands.DoBackUp               '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(UserIndex)
        
        Case eGMCommands.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
        
        Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
        
        Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(UserIndex)
        
        Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
        
        Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case eGMCommands.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case eGMCommands.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case eGMCommands.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)
        
        Case eGMCommands.night                   '/NOCHE
            Call HandleNight(UserIndex)
        
        Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case eGMCommands.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case eGMCommands.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case eGMCommands.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case eGMCommands.Restart                 '/REINICIAR
            Call HandleRestart(UserIndex)
        
        Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(UserIndex)
        
        Case eGMCommands.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case eGMCommands.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case eGMCommands.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
        
        Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(UserIndex)
            
        Case eGMCommands.Conteo
            Call HandleConteo(UserIndex)
            
        Case eGMCommands.AdminCargos             '/ADMIN ' GSZAO
            Call HandleAdminCargos(UserIndex)
            
    End Select
End With

Exit Sub

Errhandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & _
                  ". Paquete: " & Command)

End Sub

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
With UserList(UserIndex)
    Call .incomingData.ReadByte
    If .flags.TargetNpcTipo = eNPCType.Gobernador Then
        Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
    Else
        If .flags.Muerto = 1 Then
            'Si es un mapa común y no está en cana
            If UCase$(MapInfo(.Pos.map).Restringir) = "NO" And .Counters.Pena = 0 Then
                If .flags.Traveling = 0 Then
                    If Ciudades(.Hogar).map <> .Pos.map Then
                        Call goHome(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes usar este comando aquí.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(UserIndex).incomingData.length < 53 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    
    UserName = buffer.ReadASCIIString()
    
#If SeguridadAlkon Then
    Password = buffer.ReadASCIIStringFixed(32)
#Else
    Password = buffer.ReadASCIIString()
#End If
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
        
#If SeguridadAlkon Then
    If Not MD5ok(buffer.ReadASCIIStringFixed(16)) Then
        Call WriteErrorMsg(UserIndex, "El cliente está dañado, por favor descarguelo nuevamente desde www.argentumonline.com.ar")
    Else
#End If
        
        If BANCheck(UserName) Then
            Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a AOForever debido a tu mal comportamiento.")
        ElseIf Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectUser(UserIndex, UserName, Password)
        End If
#If SeguridadAlkon Then
    End If
#End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Agilidad) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Inteligencia) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Carisma) = RandomNumber(17, 18)
        .UserAtributos(eAtributos.Constitucion) = RandomNumber(17, 18)
    End With
    
    Call WriteDiceRoll(UserIndex)
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(UserIndex).incomingData.length < 62 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(UserIndex).incomingData.length < 15 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    Dim Head As Integer
    Dim mail As String
    Dim pin As String
#If SeguridadAlkon Then
    Dim MD5 As String
#End If
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creación de personajes en este server se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "server restringido a administradores. Consulte la página oficial o el foro oficial para más información.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    UserName = buffer.ReadASCIIString()
    
#If SeguridadAlkon Then
    Password = buffer.ReadASCIIStringFixed(32)
#Else
    Password = buffer.ReadASCIIString()
#End If
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
        
#If SeguridadAlkon Then
    MD5 = buffer.ReadASCIIStringFixed(16)
#End If
    
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Head = buffer.ReadInteger
    mail = buffer.ReadASCIIString()
    homeland = buffer.ReadByte()
    pin = buffer.ReadASCIIString()
#If SeguridadAlkon Then
    If Not MD5ok(MD5) Then
        Call WriteErrorMsg(UserIndex, "El cliente está dañado, por favor descarguelo nuevamente desde www.argentumonline.com.ar")
    Else
#End If
        
        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectNewUser(UserIndex, UserName, Password, race, gender, Class, mail, homeland, Head, pin)
        End If
#If SeguridadAlkon Then
    End If
#End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call Usuarios.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            If .UserDeath.EnDeath = True Or .HungerGames.HungerGamers = True Then
                Call WriteConsoleMsg(UserIndex, "No puedes hablar mientras juegas deathmatch o juegos del hambre.", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not (.flags.AdminInvisible = 1) Then
                    If .flags.Muerto = 1 Then
                        Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                    End If
                Else
                    If RTrim(Chat) <> "" Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'15/07/2009: ZaMa - Now invisible admins yell by console.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call Usuarios.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
            
        If LenB(Chat) <> 0 Then
            If .UserDeath.EnDeath = True Or .HungerGames.HungerGamers = True Then
                Call WriteConsoleMsg(UserIndex, "No puedes hablar mientras juegas deathmatch o juegos del hambre.", FontTypeNames.FONTTYPE_INFO)
            Else
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                    
                If .flags.Privilegios And PlayerType.User Then
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                    End If
                Else
                    If Not (.flags.AdminInvisible = 1) Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                    End If
                End If
            End If
        End If
        
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'15/07/2009: ZaMa - Now invisible admins wisper by console.
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        Dim targetCharIndex As Integer
        Dim TargetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = buffer.ReadInteger()
        Chat = buffer.ReadASCIIString()
        
        TargetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            If TargetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                targetPriv = UserList(TargetUserIndex).flags.Privilegios
                'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
                If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
                ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf Not EstaPCarea(UserIndex, TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                
                Else
                    '[Consejeros & GMs]
                    If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le dijo a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        If Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, vbBlue)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbBlue)
                            Call FlushBuffer(TargetUserIndex)
                            
                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'11/19/09 Pato - Now the class bandit can walk hidden.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("server> " & .Name & " ha sido echado por el server por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'TODO: Debería decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then Exit Sub
        If .flags.Meditando = True Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
        If .flags.Paralizado = 0 Then
            If Not .flags.Meditando Then
                'Move user
                Call MoveUserChar(UserIndex, heading)
               
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                   
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estás paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief And .clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToogleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call Usuarios.SetInvisible(UserIndex, .Char.CharIndex, False)
                    End If
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'Last Modified By: ZaMa
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar así este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call Usuarios.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregué un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(.ComUsu.DestUsu)
            End If
        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
    End With
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True
    End If
    
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'07/25/09: Marco - Agregué un checkeo para patear a los usuarios que tiran items mientras comercian.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim slot As Byte
    Dim Amount As Integer
    Dim check As Long
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        check = .incomingData.ReadLong
        If check = .Counters.Seguridad.Tirar Then
            ''Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("AntiCheat> Deteccion cheat al tirar items de " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        End If
        .Counters.Seguridad.Tirar = check
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If slot = FLAGORO Then
            If Amount > 10000 Then
                Call TirarOro(80000, UserIndex)
            Else
                TirarOro Amount, UserIndex
            End If
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If slot <= MAX_INVENTORY_SLOTS And slot > 0 Then
                If .Invent.Object(slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                Call DropObj(UserIndex, slot, Amount, .Pos.map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub
        End If
        
        .flags.Hechizo = .Stats.UserHechizos(Spell)
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata se puede ocultar en barca
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill) 'Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
            
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                If .flags.Navegando = 1 Then
                    If .clase <> eClass.Pirat Then
                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    Dim TotalItems As Long
    Dim ItemsPorCiclo As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger
        
        If TotalItems > 0 Then
            
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
            
        End If
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte, check As Long, Dck As Boolean
        
        slot = .incomingData.ReadByte()
        check = .incomingData.ReadLong
        Dck = .incomingData.ReadBoolean
        
        If check = .Counters.Seguridad.Poteo Then
            ''Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("AntiCheat> Posible deteccion cheat poteo de " & UserList(UserIndex).name & ". Si se repite constantemente este mensaje comunicarse con Nhelk", FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        End If
        .Counters.Seguridad.Poteo = check
        If slot <= .CurrentInventorySlots And slot > 0 Then
            If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(UserIndex, slot, Dck)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 14/01/2010 (ZaMa)
'16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
'12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
'14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueño.
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        Dim check As Long
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        check = .incomingData.ReadLong()
        If check = .Counters.Seguridad.Lanzar Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("AntiCheat> Deteccion cheat al lanzar de " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        End If
        .Counters.Seguridad.Lanzar = check
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
        Or Not InMapBounds(.Pos.map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Dim Atacked As Boolean
                Atacked = True
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    ' Tiene arma equipada?
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ' En un slot válido?
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
                    ElseIf ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                        ' La municion esta equipada en un slot valido?
                        If .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                            DummyInt = 1
                        ' Tiene munición?
                        ElseIf .MunicionEqpObjIndex = 0 Then
                            DummyInt = 1
                        ' Son flechas?
                        ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                            DummyInt = 1
                        ' Tiene suficientes?
                        ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                            DummyInt = 1
                        End If
                    ' Es un arma de proyectiles?
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    If .Genero = eGenero.Hombre Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    Atacked = UsuarioAtacaUsuario(UserIndex, tU)
                    
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Atacked = UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                ' Solo pierde la munición si pudo atacar al target, o tiro al aire
                If Atacked Then
                    With .Invent
                        ' Tiene equipado arco y flecha?
                        If ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                            DummyInt = .MunicionEqpSlot
                        
                            
                            'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the ammo, so we equip it again
                                .MunicionEqpSlot = DummyInt
                                .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .MunicionEqpSlot = 0
                                .MunicionEqpObjIndex = 0
                            End If
                        ' Tiene equipado un arma arrojadiza
                        Else
                            DummyInt = .WeaponEqpSlot
                            
                            'Take 1 knife away
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the weapon, so we equip it again
                                .WeaponEqpSlot = DummyInt
                                .WeaponEqpObjIndex = .Object(DummyInt).ObjIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .WeaponEqpSlot = 0
                                .WeaponEqpObjIndex = 0
                            End If
                            
                        End If
                        
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                    End With
               End If
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posición (" & .Pos.map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    If Not modAntiCheat.PuedeIntervalo(UserIndex, IntControl.Lanzar) Then
                        Exit Sub
                    End If
                End If
                
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Pesca
                DummyInt = .Invent.WeaponEqpObjIndex
                If DummyInt = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.map, X, Y) Then
                    Select Case DummyInt
                        Case CAÑA_PESCA
                            Call DoPescar(UserIndex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(UserIndex)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                                     Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "¡No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR And _
                    .Invent.WeaponEqpObjIndex <> HACHA_LEÑA_ELFICA Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No puedes talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles And .Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(UserIndex)
                    ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico And .Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Then
                        If .Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(UserIndex, True)
                        Else
                            Call WriteConsoleMsg(UserIndex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex 'CHECK
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No hay ninguna criatura allí!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                            Exit Sub
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes más minerales.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            Call FlushBuffer(UserIndex)
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                            Call FundirMineral(UserIndex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                            Call FundirArmas(UserIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call WriteShowBlacksmithForm(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/11/09
'05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
'***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String
        
        desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        site = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.Name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))

            
            'Update tag
             Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Maná necesario: " & .ManaRequerido & vbCrLf _
                                               & "Energía necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Adapting to new skills system.
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim Points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            Points(i) = .incomingData.ReadByte()
            
            If Points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + Points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
        With .Stats
            For i = 1 To NUMSKILLS
                If Points(i) > 0 Then
                    .SkillPts = .SkillPts - Points(i)
                    .UserSkills(i) = .UserSkills(i) + Points(i)
                    
                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100
                    End If
                    
                    Call CheckEluSkill(UserIndex, i, True)
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        Dim Amount As Integer
        Dim toSlot As Byte
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        toSlot = .incomingData.ReadByte()
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, slot, Amount, toSlot)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        Dim Amount As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, slot, Amount)
    End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Implemento nuevo sistema de foros
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim ForumMsgType As eForumMsgType
        
        Dim File As String
        Dim Title As String
        Dim Post As String
        Dim ForumIndex As Integer
        Dim postFile As String
        Dim ForumType As Byte
                
        ForumMsgType = buffer.ReadByte()
        
        Title = buffer.ReadASCIIString()
        Post = buffer.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            
            Select Case ForumType
            
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
                    
            End Select
            
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        Dim slot As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        slot = .ReadByte()
    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(slot) = .BancoInvent.Object(slot - 1)
            .BancoInvent.Object(slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(slot) = .BancoInvent.Object(slot + 1)
            .BancoInvent.Object(slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(slot + 1).Amount = TempItem.Amount
        End If
    End With
    
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim codex() As String
        
        desc = buffer.ReadASCIIString()
        codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 24/11/2009
'24/11/2009: ZaMa - Nuevo sistema de comercio
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim ObjIndex As Integer
        
        slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)
            End If
        
            Exit Sub
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((slot < 0 Or slot > UserList(UserIndex).CurrentInventorySlots) And slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If slot = FLAGORO Then
            ' Can't offer more than he has
            If Amount > .Stats.GLD - .ComUsu.goldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.goldAmount Then
                    Amount = .ComUsu.goldAmount * (-1)
                End If
            End If
        Else
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If slot <> 0 Then ObjIndex = .Invent.Object(slot).ObjIndex
            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, _
                TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
               
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
           
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)
                End If
            End If
            
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu mochila mientras la estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
        End If
        
        
                
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, slot = FLAGORO)
        
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, reason)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, error) Then
            Call WriteConsoleMsg(UserIndex, error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
           Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    On Error Resume Next
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        
        Call WriteConsoleMsg(UserIndex, OnlineString & "Número de usuarios: " & CStr(OnlineNum), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Public Sub GenerateOnlineString()
    Dim i As Long
    Dim Count As Long
    Dim tStr As String
    For i = 1 To LastUser
        If LenB(UserList(i).Name) <> 0 Then
            If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.ChaosCouncil Or PlayerType.RoyalCouncil) Then
                Count = Count + 1
                tStr = tStr & UserList(i).Name & ", "
            End If
        End If
    Next i
    If Len(tStr) > 3 Then
        tStr = Left$(tStr, Len(tStr) - 2) & vbCrLf
    Else
        tStr = vbNullString
    End If
    OnlineString = tStr
    OnlineNum = Count
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tú no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim Percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub


''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)
            
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Maná restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
       End If
        
        
        Call WriteMeditateToggle(UserIndex)
        
        FlushBuffer UserIndex
        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 15 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
           'ElseIf .Stats.ELV < 25 Then
           '     .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 30 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
            If .Counters.tInicioMeditar = 0 Then .Counters.tInicioMeditar = TIEMPO_INICIOMEDITAR

            Call WriteConsoleMsg(UserIndex, "Comienzas a meditar", FontTypeNames.FONTTYPE_INFO)
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleSendEvent(ByVal ID As Integer)
 
On Error GoTo Errhandler
 
    With UserList(ID)
     
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
     
        'Remove packet ID
        Call buffer.ReadByte
     
 
        Dim nEvent As Byte
        Dim LoopC As Long
        
        nEvent = buffer.ReadByte()
        
        Dim players() As Integer
        ReDim players(1 To nEvent) As Integer
        
        players(1) = ID
     
        For LoopC = 2 To nEvent
            players(LoopC) = NameIndex(buffer.ReadASCIIString())
        Next LoopC
             
       '' Call Eventos_Automaticos.Send_Event(Players(), nEvent)
     
        Call .incomingData.CopyBuffer(buffer)
     
    End With
Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
 
End Sub
 
Private Sub HandleDoEvent(ByVal ID As Integer)
 
On Error GoTo Errhandler
 
    With UserList(ID)
     
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
     
        'Remove packet ID
        Call buffer.ReadByte
     
        Dim nEvent As Byte
        Dim Amount_Team As Byte
        Dim Drop    As Boolean
        Dim LoopC As Long
        Dim Inscription_Prize As Boolean
        Dim Max_Potions As Integer
        Dim Gold_Inscription As Long
        Dim Gold_Prize As Long
     
        nEvent = buffer.ReadByte()
        Amount_Team = buffer.ReadByte()
        Drop = buffer.ReadBoolean()
        Inscription_Prize = buffer.ReadBoolean()
        Max_Potions = buffer.ReadInteger()
        Gold_Inscription = buffer.ReadLong()
        Gold_Prize = buffer.ReadLong()
     
       '' Call Eventos_Automaticos.Do_Event(ID, nEvent, Amount_Team, Drop, Inscription_Prize, Max_Potions, Gold_Inscription, Gold_Prize)
     
        Call .incomingData.CopyBuffer(buffer)
     
    End With
Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
 
End Sub
 
Private Sub HandleEnterEvent(ByVal ID As Integer)
 
On Error GoTo Errhandler
 
    With UserList(ID)
     
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
     
        'Remove packet ID
        Call buffer.ReadByte
     
        Dim ID_Send As Integer
        Dim nEvent As Byte
 
        ID_Send = NameIndex(buffer.ReadASCIIString())
        nEvent = buffer.ReadByte
 
       '' Call Eventos_Automaticos.Accept_Send(ID, ID_Send, nEvent)
     
        Call .incomingData.CopyBuffer(buffer)
     
    End With
Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
 
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Habilita/Deshabilita el modo consulta.
'01/05/2010: ZaMa - Agrego validaciones.
'***************************************************
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Comando exclusivo para gms
        If Not EsGM(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGM(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim UserName As String
        UserName = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
        ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call Usuarios.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                End If
            End With
        End If
        
        Call Usuarios.SetConsulatMode(UserConsulta)
    End With

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "¡¡Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Integer
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            .ComUsu.goldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim Matados As Integer
    Dim NextRecom As Integer
    Dim Diferencia As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        NextRecom = .Faccion.NextRecompensa
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        Else
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
            End If
            
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos más y te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que estás en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "¡¡No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaArmadaReal(UserIndex)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "¡¡No perteneces a la legión oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendMOTD(UserIndex)
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim Time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    Time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub




''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Shares owned npcs with other user
'***************************************************
    
    Dim TargetUserIndex As Integer
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Didn't target any user
        TargetUserIndex = .flags.TargetUser
        If TargetUserIndex = 0 Then Exit Sub
        
        ' Can't share with admins
        If EsGM(TargetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Pk or Caos?
        If criminal(UserIndex) Then
            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma facción!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
            ' Pks don't need to share with anyone
            Else
                Exit Sub
            End If
        
        ' Ciuda or Army?
        Else
            ' Can't share
            If criminal(TargetUserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith
        If SharingUserIndex = TargetUserIndex Then Exit Sub
        
        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        End If
        
        .flags.ShareNpcWith = TargetUserIndex
        
        Call WriteConsoleMsg(TargetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Stop Sharing owned npcs with other user
'***************************************************
    
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        SharingUserIndex = .flags.ShareNpcWith
        
        If SharingUserIndex <> 0 Then
            
            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
            .flags.ShareNpcWith = 0
        End If
        
    End With

End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'15/07/2009: ZaMa - Now invisible admins only speak by console
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                
                If Not (.flags.AdminInvisible = 1) Then _
                    Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, vbYellow))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim n As Integer
        
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = buffer.ReadASCIIString()
        
        n = FreeFile
        Open App.path & "\Dat\Reportes.ini" For Append Shared As n
        Print #n, "Usuario:" & .Name & " reporte> " & bugReport
        Close #n
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim description As String
        
        description = buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(UserIndex, "La descripción tiene caracteres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .desc = Trim$(description)
                Call WriteConsoleMsg(UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote As String
        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMA
'Last Modification: 05/17/06
'
'***************************************************
    
    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call modGuilds.SendGuildNews(UserIndex)
    End With
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Name As String
        Dim Count As Integer
        
        Name = buffer.ReadASCIIString()
        
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")
            End If
            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")
            End If
            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")
            End If
            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")
            End If
            
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                If FileExist(CharPath & Name & ".chr", vbNormal) Then
                    Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        While Count > 0
                            Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1
                        Wend
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************
#If SeguridadAlkon Then
    If UserList(UserIndex).incomingData.length < 65 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newpass As String
        Dim oldPass2 As String
        
        'Remove packet ID
        Call buffer.ReadByte
        
#If SeguridadAlkon Then
        oldPass = UCase$(buffer.ReadASCIIStringFixed(32))
        newpass = UCase$(buffer.ReadASCIIStringFixed(32))
#Else
        oldPass = UCase$(buffer.ReadASCIIString())
        newpass = UCase$(buffer.ReadASCIIString())
#End If
        
        If LenB(newpass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes especificar una contraseña nueva, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = UCase$(GetVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password"))
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password", newpass)
                Call WriteConsoleMsg(UserIndex, "La contraseña fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Amount > .Stats.Banco Then
            Amount = .Stats.Banco
        End If
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim TalkToKing As Boolean
    Dim TalkToDemon As Boolean
    Dim NpcIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNPC
        If NpcIndex <> 0 Then
            ' Es rey o domonio?
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                'Rey?
                If Npclist(NpcIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                ' Demonio
                Else
                    TalkToDemon = True
                End If
            End If
        End If
               
        'Quit the Royal Army?
        If .Faccion.ArmadaReal = 1 Then
            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            
            Else
                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call ExpulsarFaccionReal(UserIndex, False)
                
            End If
        
        'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Then
            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí maldito criminal!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call ExpulsarFaccionCaos(UserIndex, False)
            End If
        ' No es faccionario
        Else
        
            ' Si le hablaba al rey o demonio, le repsonden ellos
            If NpcIndex > 0 Then
                Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(UserIndex, "¡No perteneces a ninguna facción!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
        End If
        
    End With
    
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String
        
        Text = buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
        
        Call WriteShowGuildAlign(UserIndex)
    End With
End Sub
    
''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim error As String
        
        clanType = .incomingData.ReadByte()
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "¡Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .ip)
            Exit Sub
        End If
        
        Select Case UCase$(Trim(clanType))
            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA
            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION
            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO
            Case eClanType.ct_GM
                .FundandoGuildAlineacion = ALINEACION_MASTER
            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA
            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL
            Case Else
                Call WriteConsoleMsg(UserIndex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
        End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub
''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.map = map Then
                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, X, Y) Then
                    Call FindLegalPos(tUser, map, X, Y)
                    Call WarpUserChar(tUser, map, X, Y, True, True)
                    Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el server de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

'
''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio
'Last Modification: 12/09/09
'
'***************************************************
    With UserList(UserIndex)
        Dim ItemIndex As Integer
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ItemIndex = .incomingData.ReadInteger()
        
        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, UserIndex) Then Exit Sub
        
        Call DoUpgrade(UserIndex, ItemIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim Names() As String
    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim Names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    Names(Count) = UserList(i).Name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, Names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then _
                    Users = Users & " (*)"
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 240 Then
                        Call WriteConsoleMsg(UserIndex, "No puedés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & Time)
                        End If
                        
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarceló a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & Time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/06/2009
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim n As Byte
        Dim UserCharPath As String
        Dim Var As Long
        
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body y level
                    valido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills Or _
                            opcion = eEditOptions.eo_addGold
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If

        If valido Then
            UserCharPath = CharPath & UserName & ".chr"
            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Estás intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intentó editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    Case eEditOptions.eo_Gold
                        If val(Arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If val(Arg1) > 20000000 Then
                                Arg1 = 20000000
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(Arg1) >= 25 Then
                            
                            Dim GI As Integer
                            If tUser <= 0 Then
                                GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                            Else
                                GI = UserList(tUser).GuildIndex
                            End If
                            
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                                    ' Si esta online le avisamos
                                    If tUser > 0 Then _
                                        Call WriteConsoleMsg(tUser, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
                                End If
                            End If
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).clase = LoopC
                            End If
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
                                Call WriteVar(UserCharPath, "Skills", "EXPSK" & LoopC, 0)
                                
                                If Arg2 < MAXSKILLPOINTS Then
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, ELU_SKILL_INICIAL * 1.05 ^ Arg2)
                                Else
                                    Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, 0)
                                End If
    
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                                Call CheckEluSkill(tUser, LoopC, True)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = val(Arg1)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "REP", "Nobles", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.NobleRep = Var
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "NOB "
                        
                    Case eEditOptions.eo_Asesino
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "REP", "Asesino", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.AsesinoRep = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "ASE "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                        Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano
                            Case "ELFO"
                                raza = eRaza.Elfo
                            Case "DROW"
                                raza = eRaza.Drow
                            Case "ENANO"
                                raza = eRaza.Enano
                            Case "GNOMO"
                                raza = eRaza.Gnomo
                            Case Else
                                raza = 0
                        End Select
                        
                            
                        If raza = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim bankGold As Long
                        
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(UserIndex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                bankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                                Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
                                Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim targetName As String
        Dim TargetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        TargetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If TargetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, TargetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToogleBoatBody(UserIndex)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If
                        
                        If .flags.Traveling = 1 Then
                            .flags.Traveling = 0
                            .Counters.goHome = 0
                            Call WriteMultiMessage(tUser, eMessages.CancelHome)
                        End If
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                    If .flags.Traveling = 1 Then
                        .Counters.goHome = 0
                        .flags.Traveling = 0
                        Call WriteMultiMessage(tUser, eMessages.CancelHome)
                    End If
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.map = map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    Call WriteConsoleMsg(UserIndex, "Sólo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquía mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echó a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echó a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                'If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                '    Call WriteConsoleMsg(UserIndex, "¿¿Estás loco?? ¿¿Cómo vas a piñatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                'Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                'End If
            Else
                Call WriteConsoleMsg(UserIndex, "No está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & Time)
                
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no está baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no está online.", FontTypeNames.FONTTYPE_INFO)
            Else

                '// G TOYZ;
                If UserList(tUser).EnEvento = True Then
                    Call WriteConsoleMsg(UserIndex, "El jugador se encuentra en algun evento.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                      (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                        X = .Pos.X
                        Y = .Pos.Y + 1
                        Call FindLegalPos(tUser, .Pos.map, X, Y)
                        Call WarpUserChar(tUser, .Pos.map, X, Y, True, True)
                        Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call mod_Limpieza.LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & "> " & message, FontTypeNames.FONTTYPE_GUILD))
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).Name & " > " & message
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ningún personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 22/03/2010
'15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
'22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim Radio As Byte
        
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = TELEP_OBJ_INDEX
        
        With MapData(.Pos.map, .Pos.X, .Pos.Y - 1)
            .TileExit.map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
        
        Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
                If MapData(.TileExit.map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJÉRCITO REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).ObjInfo.ObjIndex > 0 Then
                        bIsExit = MapData(.Pos.map, X, Y).TileExit.map > 0
                        If ItemNoEsDeMapa(MapData(.Pos.map, X, Y).ObjInfo.ObjIndex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim lista As String
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneó al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del server", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del server.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & Time)
                Next LoopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm está o no online.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).ip
        End If
        
        reason = buffer.ReadASCIIString()
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).ip = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        Dim tStr As String
        tObj = .incomingData.ReadInteger()

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
            
        Call LogGM(.Name, "/CI: " & tObj)
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN: FUERON CREADOS ***100*** ÍTEMS, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.Amount = 100
        Objeto.ObjIndex = tObj
        Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST")
        
        If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And _
            MapData(.Pos.map, .Pos.X, .Pos.Y).TileExit.map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHÓ DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast música: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & Time)
                    
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        ''Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        slot = buffer.ReadByte() 'Que Slot?
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user index?
            
            Call LogGM(.Name, .Name & " Checkeó el slot " & slot & " de " & UserName)
               
            If tIndex > 0 Then
                If slot > 0 And slot <= UserList(tIndex).CurrentInventorySlots Then
                    If UserList(tIndex).Invent.Object(slot).ObjIndex > 0 Then
                        Call WriteConsoleMsg(UserIndex, " Objeto " & slot & ") " & ObjData(UserList(tIndex).Invent.Object(slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay ningún objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "NHELK" Then Exit Sub
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("|RESETUPD", FontTypeNames.FONTTYPE_INFO))

    End With
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "NHELK" Or UCase$(.Name) <> "G Toyz" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinició el mundo.")
        
        ''Call Reiniciarserver(True)
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        DeNoche = Not DeNoche
        
        Dim i As Long
        
        For i = 1 To NumUsers
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
            End If
        Next i
    End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del server.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        'PARTY 9010
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.map).backup = 1
        Else
            MapInfo(.Pos.map).backup = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).backup)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).backup, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la información sobre si es PK el mapa.")
        
        MapInfo(.Pos.map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la información sobre si es restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Restringir = tStr
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Restringido: " & MapInfo(.Pos.map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto = nomagic
            Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto = noinvi
            Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la información sobre si está permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).ResuSinEfecto = noresu
            Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información del terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Terreno = tStr
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frío en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la información de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Zona = tStr
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.map))
        
        Call GrabarMapa(.Pos.map, App.path & "\WorldBackUp\Mapa" & .Pos.map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0
        End With
    
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj está online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & Time)
                                
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneó con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index
            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index
            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "server habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(UserIndex, "server restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡¡¡" & .Name & " VA A APAGAR EL server!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & Time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)
            If tUser > 0 Then _
                Call VolverCriminal(tUser)
        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                Char = CharPath & UserName & ".chr"
                
                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del server.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).Texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con éxito.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).Texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)
    End With
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 01/23/10 (Marco)
'Modify server.ini
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        
        Call buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call buffer.ReadByte

        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String

        'Obtengo los parámetros
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes modificar esa información desde aquí!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor según llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long

    error = Err.Number

On Error GoTo 0
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(IIf(UserList(UserIndex).GuildIndex <> 0, modGuilds.GuildName(UserList(UserIndex).GuildIndex), ""))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(UserIndex).Stats.Banco)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal map As Integer, ByVal version As Integer, ByVal MapName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(map)
        Call .WriteInteger(version)
        Call .WriteASCIIString(MapName)
    End With
Exit Sub
Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
            
''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, _
                                ByVal Privileges As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, Name, NickColor, Privileges))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim tmp As String
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteLong(UserList(UserIndex).Stats.Diam)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 3/12/09
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(ObjIndex))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)
'***************************************************
'Author: Budi
'Last Modification: 01/12/09
'Writes the "AddSlots" message to the given user's outgoing data buffer
'***************************************************
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddSlots)
        Call .WriteByte(Mochila)
    End With
End Sub


''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(slot).Amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.Valor)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(slot))
        
        If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(slot)).Nombre)
        Else
            Call .WriteASCIIString("(Vacio)")
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(Obj.MaderaElfica)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).Texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef Obj As Obj, ByVal price As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Last Modified by: Budi
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
On Error GoTo Errhandler
    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(Obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        
        Call .WriteLong(UserList(UserIndex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.NobleRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.Promedio)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el server ni en el cliente!!!
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal ForumType As eForumType, _
                    ByRef Title As String, ByRef Author As String, ByRef message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'02/01/2010: ZaMa - Now sends Author and forum type
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler

    Dim Visibilidad As Byte
    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If esCaos(UserIndex) Or EsGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
        End If
        
        If esArmada(UserIndex) Or EsGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
        End If
        
        Call .outgoingData.WriteByte(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGM(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
        End If
        
        Call .outgoingData.WriteByte(CanMakeSticky)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DiceRoll" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'''***************************************************
On Error GoTo Errhandler
    ''Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    ''Call WritePosUpdate(UserIndex)
    FlushBuffer UserIndex
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'Writes the "SendSkills" message to the given user's outgoing data buffer
'11/19/09: Pato - Now send the percentage of progress of the skills.
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.clase)
        
        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(UserIndex).Stats.UserSkills(i))
            If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
                Call .outgoingData.WriteByte(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
            Else
                Call .outgoingData.WriteByte(0)
            End If
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            tmp = tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            tmp = tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            tmp = tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        ' Prepare guild member's list
        tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            tmp = tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            tmp = tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            tmp = tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
        
        ' Prepare guild member's list
        tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            tmp = tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        
        If ObjIndex > 0 Then
            Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).OBJType)
            Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
            Call .WriteInteger(ObjData(ObjIndex).MinHIT)
            Call .WriteInteger(ObjData(ObjIndex).MaxDef)
            Call .WriteInteger(ObjData(ObjIndex).MinDef)
            Call .WriteLong(SalePrice(ObjIndex))
            Call .WriteASCIIString(ObjData(ObjIndex).Name)
        Else ' Borra el item
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
        End If
    End With
Exit Sub


Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            tmp = tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            tmp = tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(tmp) <> 0 Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            tmp = tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(tmp) Then _
            tmp = Left$(tmp, Len(tmp) - 1)
        
        Call .WriteASCIIString(tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    If UserList(UserIndex).outgoingData Is Nothing Then Exit Sub
    With UserList(UserIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(color)
        ' Write rgb channels and save one byte from long :D
        'Call .WriteByte(color And &HFF)
        'Call .WriteByte((color And &HFF00&) \ &H100&)
        'Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal NickColor As Byte, _
                                ByVal Privileges As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, ByVal NickColor As Byte, _
                                                ByRef Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'15/01/2010: ZaMa - Now sends the nick color instead of the status.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************
On Error GoTo Errhandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)
        
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(slot)
    End With
    
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Function PrepareMessageCreateDamage(ByVal X As Byte, ByVal Y As Byte, ByVal Label As String, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    With auxiliarBuffer
        .WriteByte ServerPacketID.CreateDamage
        .WriteASCIIString Label
        .WriteByte r
        .WriteByte g
        .WriteByte b
        .WriteByte X
        .WriteByte Y
        PrepareMessageCreateDamage = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub HandleDisolverClan(ByVal UserIndex As Integer)
'Author: Shak
'Last modification: -
'Disolver nuestro clan cuando lo DESEEMOS.
 
    With UserList(UserIndex)
        Call .incomingData.ReadByte
       
        Dim NombreClan As String, LiderClan As String
        
   
        If .GuildIndex <> 0 Then 'Pertenece a algún clanj
            NombreClan = modGuilds.GuildName(.GuildIndex)
            LiderClan = modGuilds.GuildLeader(.GuildIndex)
            If UCase$(LiderClan) <> UCase$(.Name) Then Exit Sub
           
            If GetVar(App.path & "\guilds\" & NombreClan & "-members.mem", "INIT", "NroMembers") > 1 Then
                Call WriteConsoleMsg(UserIndex, "No has echado a todos los miembros del clan.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
       
            Call WriteConsoleMsg(UserIndex, "Has disuelto el clan. Para reanudarlo tipea /REANUDARCLAN.", FontTypeNames.FONTTYPE_INFO)
            .flags.ExClan = .GuildIndex
            Call WriteVar(CharPath & LiderClan & ".chr", "GUILD", "GUILDINDEX", 0)
            Call WarpUserChar(UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
            Call RefreshCharStatus(UserIndex)
            Call WriteErrorMsg(UserIndex, "Vuelve a ingresar al juego para ver los cambios.")
            FlushBuffer UserIndex
            Call CloseSocket(UserIndex)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_INFO)
        End If
   
    End With
End Sub
 
Public Sub HandleReanudarClan(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Dim LiderClan As String
        
        '' LiderClan = modGuilds.GuildLeader(.GuildIndex)
       
        If .GuildIndex <> 0 Then
            Exit Sub
        End If
       
        If .flags.ExClan = 0 Then Exit Sub
       
        If .Stats.GLD < 1000000 Then 'Un millon de monedas para reaunudarlo.
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero. (1.000.000 monedas de oro).", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        .Stats.GLD = .Stats.GLD - 1000000
        Call WriteUpdateGold(UserIndex)
       
        .GuildIndex = .flags.ExClan
        Call WriteVar(CharPath & .Name & ".chr", "GUILD", "GUILDINDEX", .flags.ExClan)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Se ha reanudado el clan " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
        .flags.ExClan = 0
        Call WarpUserChar(UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
        Call RefreshCharStatus(UserIndex)
   
    End With
End Sub

Private Sub handleRetirarTodo(ByVal UserIndex As Integer)
    Dim Amount As Long
    With UserList(UserIndex)
        Call .incomingData.ReadByte 'Remove packet id
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Amount = .Stats.Banco
        
        If Amount > 0 Then
             .Stats.Banco = 0
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes oro en el banco.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

Private Sub handleDepositarTodo(ByVal UserIndex As Integer)
    Dim Amount As Long
    With UserList(UserIndex)
        Call .incomingData.ReadByte 'Remove packet id
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Amount = .Stats.GLD
        
        If Amount > 0 Then
             .Stats.Banco = .Stats.Banco + Amount
             .Stats.GLD = .Stats.GLD - Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub


Private Sub HandleSolicitudParty(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte 'packet id
        Dim tUser As Integer, tParty As Integer
        
        If .PartyIndex <> 0 Then 'esta en otra party?
            Call WriteConsoleMsg(UserIndex, "Party> Ya estás en una party", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        If .flags.Muerto > 0 Then 'esta muerto?
            Call WriteConsoleMsg(UserIndex, "Party> ¡Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        Dim d As Boolean
        tUser = .flags.TargetUser
        If tUser <= 0 Then 'clickeo a alguien?
            Call WriteConsoleMsg(UserIndex, "Party> Primero haz click izquierdo sobre el lider de la party", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        
        tParty = UserList(tUser).PartyIndex
        
        If tParty <> 0 Then
            If Party(tParty).Lider = tUser Then 'Cumple todos los requisitos
                UserList(UserIndex).PartySolicitud = tUser
                Dim X As Long, BYE As Boolean
                For X = 0 To 20
                    If Party(tParty).Solicitudes(X) = UserIndex Then
                        BYE = True
                        Exit For
                    End If
                Next X
                If Not BYE Then
                    For X = 0 To 20
                        If Party(tParty).Solicitudes(X) <= 0 Then
                            Exit For
                        End If
                    Next X
                    Party(tParty).Solicitudes(X) = UserIndex
                End If
                If Not BYE Then Call WritePartyForm(Party(tParty).Lider)
                
                Call WriteConsoleMsg(tUser, .Name & " te ha enviado una solicitud de party", FontTypeNames.FONTTYPE_PARTY)
                Call WriteConsoleMsg(UserIndex, "El lider decidirá si te acepta en la party", FontTypeNames.FONTTYPE_PARTY)
            Else
                Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " no es fundador de ninguna party", FontTypeNames.FONTTYPE_PARTY)
                Exit Sub
            End If
        End If
    End With
End Sub
            
Private Sub HandleCrearParty(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte 'packet id
        Call mod_Party.CrearParty(UserIndex)
    End With
End Sub
            
Private Sub HandleAceptarParty(ByVal UserIndex As Integer)
On Error GoTo Errhandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte 'packetid
        Dim Nick As String
        Nick = buffer.ReadASCIIString
        Dim nIndex As Integer
        nIndex = NameIndex(Nick)
        If nIndex <> 0 Then
            If UserList(nIndex).PartySolicitud = UserIndex Then
                .PartySolicitud = 0
                With UserList(nIndex)
                    With Party(UserList(UserIndex).PartyIndex)
                        Dim X As Long, find As Byte
                        For X = 1 To MaxUsuariosParty
                            If .User(X).UserIndex <= 0 Then
                                find = X
                                Exit For
                            End If
                        Next X
                        For X = 0 To 20
                            If .Solicitudes(X) = nIndex Then
                                .Solicitudes(X) = 0
                                Exit For
                            End If
                        Next X
                        .User(find).UserIndex = nIndex
                        .User(find).ExpAcumulada = 0
                        ''Call mod_Party.PorcPartyValidos(UserList(UserIndex).PartyIndex)
                        Dim u As Byte
                        For X = 1 To MaxUsuariosParty
                            If .User(X).UserIndex <> 0 Then
                                u = u + 1
                            End If
                        Next X
                        
                        For X = 1 To MaxUsuariosParty
                            If .User(X).UserIndex <> 0 Then
                                .User(X).Porcentaje = Round((100 / u) - 0.5)
                            End If
                        Next X
                    End With
                    .PartyIndex = UserList(UserIndex).PartyIndex
                    
                    
                    Call MensajeParty(.PartyIndex, "Party> " & .Name & " ha ingresado a la party")
                    Call WritePartyForm(UserIndex)
                End With
            Else
                Call WriteConsoleMsg(UserIndex, "Party> Ese usuario no te ha enviado ninguna solicitud o su solicitud ya ha caducado", FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
        
    End With
    
    
Exit Sub
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
            
Private Sub HandleRequestPartyForm(ByVal UserIndex As Integer)
On Error GoTo errh
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        Dim pIndex As Integer
        pIndex = .PartyIndex
        If pIndex <= 0 Or pIndex >= MaxPartys Then
            Call mod_Party.CrearParty(UserIndex)
            'Call WriteConsoleMsg(UserIndex, "Party> No perteneces a ninguna party", FontTypeNames.FONTTYPE_PARTY)
            ''Exit Sub
        Else
            Call WritePartyForm(UserIndex)
        End If

    End With
    Exit Sub
errh:

    Dim X As Integer
    X = NameIndex("NHELK")
    If X <> 0 Then
        WriteConsoleMsg X, "Bug sistema de partys sub handlerequestpartyform> " & Err.Number & "-" & Err.description, FontTypeNames.FONTTYPE_CENTINELA
    End If
    LogError "Bug sistema de partys sub handlerequestpartyform> " & Err.Number & "-" & Err.description '', FontTypeNames.FONTTYPE_CENTINELA
End Sub

Public Sub WritePartyForm(ByVal UserIndex As Integer)
    On Error GoTo errh
    With UserList(UserIndex)
        Dim pIndex As Integer
        pIndex = .PartyIndex
        Call .outgoingData.WriteByte(ServerPacketID.PartyForm)
        Call .outgoingData.WriteByte(MaxUsuariosParty)
        
        Dim X As Long
        For X = 1 To MaxUsuariosParty
            If Party(pIndex).User(X).UserIndex > 0 Then
                Call .outgoingData.WriteASCIIString(UserList(Party(pIndex).User(X).UserIndex).Name)
            Else
                Call .outgoingData.WriteASCIIString(vbNullString)
            End If
            Call .outgoingData.WriteByte(Party(pIndex).User(X).Porcentaje)
            Call .outgoingData.WriteLong(Party(pIndex).User(X).ExpAcumulada)
        Next X
        For X = 0 To 20
            If Party(pIndex).Solicitudes(X) <> 0 Then
                Call .outgoingData.WriteASCIIString(UserList(Party(pIndex).Solicitudes(X)).Name)
            Else
                Call .outgoingData.WriteASCIIString(vbNullString)
            End If
        Next X
        Call .outgoingData.WriteBoolean(IIf(Party(pIndex).Lider = UserIndex, True, False))
        Call .outgoingData.WriteByte(UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo))
    End With
    Exit Sub
    
errh:
    Dim Xz As Integer
    Xz = NameIndex("NHELK")
    If Xz <> 0 Then
        WriteConsoleMsg Xz, "Bug sistema de partys sub writepartyform> " & Err.Number & "-" & Err.description, FontTypeNames.FONTTYPE_CENTINELA
    End If
    LogError "Bug sistema de partys sub writepartyform> " & Err.Number & "-" & Err.description
End Sub

Private Sub HandleSavePartyPerc(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim nParty As Integer, sum As Byte, temp As Byte, X As Long
        nParty = .PartyIndex
        If .PartyIndex <= 0 Then
            LogError "Error guardando porcentajes de party, esto no deberia suceder ya que hay chequeos anteriores, bug GRAVE."
        Else
            Call UserList(UserIndex).incomingData.ReadByte
            With Party(nParty)
                For X = 1 To MaxUsuariosParty
                    With .User(X)
                        temp = UserList(UserIndex).incomingData.ReadByte
                        If .UserIndex > 0 Then
                            .Porcentaje = temp
                            sum = sum + temp
                        End If
                    End With
                Next X
                If sum < 100 Or sum > 100 Then
                    Call WriteConsoleMsg(UserIndex, "Party> La party ha sido cerrada por un error grave en el sistema de porcentajes, avise a un administrador", FontTypeNames.FONTTYPE_PARTY)
                    Call mod_Party.LimpiarParty(nParty, True)
                    For X = 1 To MaxUsuariosParty
                        If .User(X).UserIndex <> 0 Then
                            WriteCerrarPartyForm .User(X).UserIndex
                        End If
                    Next X
                End If
                Call mod_Party.MensajeParty(nParty, "Party> Los porcentajes han sido modificados por el lider de la party")
                
                Call mod_Party.PartyInfo(nParty)
            
                Call WritePartyForm(.Lider)
            End With
        End If
    End With
End Sub

Private Sub HandleEcharDisolverParty(ByVal UserIndex As Integer)
    Dim buffer As New clsByteQueue
    
    With UserList(UserIndex)
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Echado As String, iechado As Integer
        Echado = buffer.ReadASCIIString
        
        If .PartyIndex <= 0 Then Exit Sub 'Si no esta en ninguna party, nos vamos de aca
        
        If Party(.PartyIndex).Activa = True Then ' Esta en una party activa?\
            
            If Party(.PartyIndex).Lider = UserIndex Then ' Es el lider de la party?
                If LenB(Echado) <= 0 Then
                    Call mod_Party.DisolverParty(.PartyIndex)
                Else
                    iechado = NameIndex(Echado)
                    
                    If iechado <> UserIndex And iechado > 0 Then _
                    Call mod_Party.EcharParty(.PartyIndex, iechado, True)
                End If
            Else
                Call mod_Party.EcharParty(.PartyIndex, UserIndex, False)
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


Public Function PrepareMessageCreateProjectile(ByVal CharIndex As Integer, ByVal victimIndex As Integer, ByVal GrhIndex As Integer) As String

        With auxiliarBuffer
            Call .WriteByte(ServerPacketID.CreateProjectile)
            
            Call .WriteInteger(CharIndex)
            Call .WriteInteger(victimIndex)
            Call .WriteInteger(GrhIndex)
            
            PrepareMessageCreateProjectile = .ReadASCIIStringFixed(.length)
        End With

End Function


Private Sub HandleRetar(ByVal UserIndex As Integer)
    On Error GoTo Errhandler
    Dim tUser As String, iUser As Integer, bItems As Boolean, lOro As Long, Plantes As Boolean, Pociones As Integer, Cascos_Escudos As Boolean, Personaje As Boolean, AIM As Boolean
    Dim buffer As New clsByteQueue
    
    With UserList(UserIndex)
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        tUser = buffer.ReadASCIIString
        iUser = NameIndex(tUser)
        lOro = buffer.ReadLong
        bItems = buffer.ReadBoolean
        Plantes = buffer.ReadBoolean
        Pociones = buffer.ReadInteger
        Cascos_Escudos = buffer.ReadBoolean
        Personaje = buffer.ReadBoolean
        AIM = buffer.ReadBoolean
        
        If Pociones > 5000 Then
            WriteConsoleMsg UserIndex, "El límite del límite de pociones de 5.000", FontTypeNames.FONTTYPE_INFOBOLD
        Else
            If iUser <> 0 Then
                mod_Retos1vs1.MandarReto UserIndex, iUser, lOro, bItems, Plantes, Pociones, Cascos_Escudos, Personaje, AIM
            Else
                WriteConsoleMsg UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO
            End If
        End If
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
        
End Sub


Private Sub HandleAceptarReto(ByVal UserIndex As Integer)
    On Error GoTo Errhandler
    Dim tUser As String, iUser As Integer
    Dim buffer As New clsByteQueue
    
    With UserList(UserIndex)
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        tUser = buffer.ReadASCIIString
        iUser = NameIndex(tUser)
        
        If iUser <> 0 Then
            mod_Retos1vs1.AceptarReto UserIndex, iUser
        Else
            WriteConsoleMsg UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO
        End If
        Call .incomingData.CopyBuffer(buffer)
    End With
    
    Exit Sub
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


Private Sub HandleAbandonarReto(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        ''Call mod_Retos.TerminarReto
    End With
End Sub

Sub EnviarListaViajes(ByVal UserIndex As Integer)
    Dim LoopC As Long
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ListarViajes)
        Call .WriteByte(NumPasajes)
        For LoopC = 1 To NumPasajes
            Call .WriteASCIIString(Pasaje(LoopC).Nombre)
            Call .WriteInteger(CInt(Pasaje(LoopC).precio))
        Next LoopC
    End With
End Sub

Private Sub HandleViajar(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Dim p As Byte
        p = .incomingData.ReadByte
        If p <= 0 Then Exit Sub
        If p > NumPasajes Then Exit Sub
        If .Stats.GLD < Pasaje(p).precio Then
            WriteConsoleMsg UserIndex, "No tienes el oro necesario", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        If .Stats.ELV < Pasaje(p).LvlMin Then
            WriteConsoleMsg UserIndex, "El mapa es demasiado peligroso para tu nivel", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD - Pasaje(p).precio
        WriteUpdateGold UserIndex
        
        WriteConsoleMsg UserIndex, "Has sido llevado a " & Pasaje(p).Nombre, FontTypeNames.FONTTYPE_INFO
        WarpUserChar UserIndex, Pasaje(p).mapa, Pasaje(p).X, Pasaje(p).Y, True, , True
        
    End With
End Sub

Public Function PrepareMessageMovimientSW(ByVal Char As Integer, ByVal MovimientClass As Byte)
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.MovimientSW)
        Call .WriteInteger(Char)
        Call .WriteByte(MovimientClass)
        PrepareMessageMovimientSW = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub HandleConteo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim Num As Byte
        Call .incomingData.ReadByte
        Num = .incomingData.ReadByte
        If EsGM(UserIndex) Then
            CR = Num + 1
        End If
    End With
End Sub

Public Sub WriteCerrarPartyForm(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.CerrarPartyForm)
    End With
End Sub

Private Sub HandleSendReto3(ByVal ID As Integer)
 
On Error GoTo Errhandler
 
    With UserList(ID)
   
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
   
        'Remove packet ID
        Call buffer.ReadByte
   
        Dim players(1 To 6) As Integer
        Dim Gold As Long
        Dim Items As Boolean
        Dim Potions As Integer
   
        players(1) = ID
        players(2) = NameIndex(buffer.ReadASCIIString())
        players(3) = NameIndex(buffer.ReadASCIIString())
        players(4) = NameIndex(buffer.ReadASCIIString())
        players(5) = NameIndex(buffer.ReadASCIIString())
        players(6) = NameIndex(buffer.ReadASCIIString())
        Gold = buffer.ReadLong()
        Items = buffer.ReadBoolean()
        Potions = buffer.ReadInteger()
   
        Call mod_Retos3vs3.Send_Reto(players(), Gold, Items, Potions)
   
        Call .incomingData.CopyBuffer(buffer)
   
    End With
    Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
 
End Sub
 
Private Sub HandleAcceptReto3(ByVal UserIndex As Integer)
 
On Error GoTo Errhandler
 
    With UserList(UserIndex)
   
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
   
        'Remove packet ID
        Call buffer.ReadByte
   
        Dim Name_Send As String
        Dim ID_Send As Integer
   
        Name_Send = buffer.ReadASCIIString()
        ID_Send = NameIndex(Name_Send)
        
        Call mod_Retos3vs3.Accept_Reto(UserIndex, ID_Send)
   
        Call .incomingData.CopyBuffer(buffer)
   
    End With
    Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub
 
Private Sub handleRefuseReto3(ByVal UserIndex As Integer)
 
On Error GoTo Errhandler
 
    With UserList(UserIndex)
   
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
   
        'Remove packet ID
        Call buffer.ReadByte
   
        Dim Name_Send As String
        Dim ID_Send As Integer
   
        Name_Send = buffer.ReadASCIIString()
        ID_Send = NameIndex(Name_Send)
   
        Call mod_Retos3vs3.Cancel_Send(ID_Send, UserIndex)
   
        Call .incomingData.CopyBuffer(buffer)
   
    End With
    Exit Sub
 
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandlePublicarUser(ByVal UserIndex As Integer)
    On Error GoTo Errhandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte ''packet id
        
        Dim precio As Long
        precio = buffer.ReadLong
        
        Dim pin As String
        pin = buffer.ReadASCIIString
        
        Dim Pj As String
        Pj = buffer.ReadASCIIString
        
        Dim p As String
        
        p = buffer.ReadASCIIString
        '¿Es el pin valido?
        If (pin) <> (GetVar(CharPath & UCase$(.Name) & ".chr", "INIT", "Pin")) Then
            Call WriteConsoleMsg(UserIndex, "Pin incorrecto.", FontTypeNames.FONTTYPE_INFOBOLD)
        Else
            Call PublicarUser(UserIndex, precio, Pj, p)
        End If

        Call .incomingData.CopyBuffer(buffer)
    End With
    
    Exit Sub
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub
Private Sub handleComprarUser(ByVal UserIndex As Integer)
    On Error GoTo Errhandler
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        
        Dim slot As Byte
        slot = buffer.ReadByte
        
        Call ComprarUser(UserIndex, slot, buffer.ReadASCIIString, buffer.ReadASCIIString)
        WriteListaMercado UserIndex
        Call .incomingData.CopyBuffer(buffer)
        
    End With
    
    Exit Sub
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleRequestMercado(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Call WriteListaMercado(UserIndex)
    End With
End Sub

Private Sub HandleDoDeath(ByVal ID As Integer)
    With UserList(ID)
        Dim Slots As Byte
        Dim Drop As Byte
        Dim DropB As Byte
        Call .incomingData.ReadByte
        Slots = .incomingData.ReadByte
        Drop = .incomingData.ReadByte
        DropB = IIf(Drop = 1, True, False)
        If Not EsGM(ID) Then Exit Sub
        If Slots < 3 Then Exit Sub
        Call mod_DeathMatch.IniciarDeath(Slots, DropB)
    End With
End Sub

Private Sub HandleEnterDeath(ByVal ID As Integer)
    Call UserList(ID).incomingData.ReadByte
    Call mod_DeathMatch.EnterDeath(ID)
End Sub

Private Sub HandleCancelDeath(ByVal ID As Integer)
    Call UserList(ID).incomingData.ReadByte
    If Not EsGM(ID) Then Exit Sub
    Call mod_DeathMatch.CancelarDeath
End Sub

Private Sub HandleDoJDH(ByVal ID As Integer)
    With UserList(ID)
        Call .incomingData.ReadByte
        If Not EsGM(ID) Then Exit Sub
        Call mod_HungerGames.IniciarHunger_Games
        
    End With
End Sub

Private Sub HandleEnterJDH(ByVal ID As Integer)
    Call UserList(ID).incomingData.ReadByte
    Call mod_HungerGames.EnterHungerGames(ID)
End Sub

Private Sub HandleCancelJDH(ByVal ID As Integer)
    Call UserList(ID).incomingData.ReadByte
    If Not EsGM(ID) Then Exit Sub
    Call mod_HungerGames.Cancel_HungerGames
End Sub
Private Sub HandleCrearTorneo(ByVal ID As Integer)
    With UserList(ID)
        Call .incomingData.ReadByte
        Dim i As Long, Cupos As Byte, maxRojas As Integer, prohibidas(1 To NUMCLASES) As Boolean
        Dim Oro As Long
        Cupos = .incomingData.ReadByte
        maxRojas = .incomingData.ReadInteger
        Oro = .incomingData.ReadLong
       ' For i = 1 To NUMCLASES
       '     prohibidas(i) = .incomingData.ReadBoolean
       ' Next i
        If Not EsGM(ID) Then Exit Sub
        
        Call mod_T1v1.NuevoTorneo(ID, Cupos, maxRojas, Oro, prohibidas())
        
    End With
End Sub
Private Sub HandleCancelarTorneo(ByVal ID As Integer)
    With UserList(ID)
        Call .incomingData.ReadByte
        
        If EsGM(ID) Then
            Call mod_T1v1.CancelarTorneo
            Call MensajeTorneo("Torneo 1vs1> El torneo ha sido cancelado por un administrador.")
        End If
    End With
End Sub
Private Sub HandleEntrarTorneo(ByVal ID As Integer)
    With UserList(ID)
        Call .incomingData.ReadByte

        Call mod_T1v1.IngresarUsuario(ID)
    End With
End Sub

Private Sub HandleDoAIM(ByVal ID As Integer)
    With UserList(ID)
        Call .incomingData.ReadByte
        Dim Max_Win As Byte, Drop As Boolean
        Max_Win = .incomingData.ReadByte
        Drop = .incomingData.ReadBoolean
        If Not EsGM(ID) Then Exit Sub
        Call Event_1vs1_Aim_Melee.Do_Event(Max_Win, Drop)
    End With
End Sub

Private Sub HandleEnterAIM(ByVal ID As Integer)
    Call UserList(ID).incomingData.ReadByte
    Call Event_1vs1_Aim_Melee.Enter_Event(ID)
End Sub

Private Sub WriteListaMercado(ByVal UserIndex As Integer)
    Dim X As Long
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ListaMercado)
        
        For X = 1 To MaxUsersMercado
            Call .WriteBoolean(UserMercado(X).Ocupado)
            If UserMercado(X).Ocupado = True Then
                Call .WriteASCIIString(UserMercado(X).Nick)
                Call .WriteLong(UserMercado(X).precio)
                Call .WriteByte(UserMercado(X).Nivel)
                Call .WriteByte(UserMercado(X).Porcentaje)
                Call .WriteInteger(UserMercado(X).Vida)
                Call .WriteByte(UserMercado(X).clase)
                Call .WriteByte(UserMercado(X).raza)
                Call .WriteBoolean(Len(UserMercado(X).VentaPrivada) <> 0)
            End If
        Next X
    End With
End Sub


Private Sub handleCancelarUser(ByVal UserIndex As Integer)
On Error GoTo Errhandler
    Dim buffer As New clsByteQueue
    With UserList(UserIndex)
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        
        Dim Nick As String
        Nick = buffer.ReadASCIIString
        
        Dim pass As String
        pass = buffer.ReadASCIIString
        
        Dim pin As String
        pin = buffer.ReadASCIIString
        
        Call mod_MercadoUsers.QuitarPJ(UserIndex, Nick, pass, pin)
        
        Call .incomingData.CopyBuffer(buffer)
    End With
    
    
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleRecuperarPass(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Nick As String
        Nick = buffer.ReadASCIIString
        Dim pin As String
        pin = buffer.ReadASCIIString
        Dim newpass As String
        newpass = buffer.ReadASCIIString
        
        If Not FileExist(CharPath & UCase$(Nick) & ".chr", vbNormal) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        If (pin) <> (GetVar(CharPath & UCase$(Nick) & ".chr", "INIT", "Pin")) Then
            Call WriteErrorMsg(UserIndex, "Pin incorrecto.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        Call WriteVar(CharPath & UCase$(Nick) & ".chr", "INIT", "Password", newpass)
        Call WriteErrorMsg(UserIndex, "Ya puedes ingresar a tu personaje con la nueva contraseña")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

Private Sub HandleBorrarPj(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call buffer.ReadByte
        Dim Nick As String
        Nick = buffer.ReadASCIIString
        Dim pin As String
        pin = buffer.ReadASCIIString
        Dim pass As String
        pass = buffer.ReadASCIIString
        
        If Not FileExist(CharPath & UCase$(Nick) & ".chr", vbNormal) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        If (pin) <> (GetVar(CharPath & UCase$(Nick) & ".chr", "INIT", "Pin")) Then
            Call WriteErrorMsg(UserIndex, "Pin incorrecto.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        If UCase$(pass) <> UCase$(GetVar(CharPath & UCase$(Nick) & ".chr", "INIT", "Password")) Then
            Call WriteErrorMsg(UserIndex, "Contraseña incorrecta.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        Call Kill(CharPath & UCase$(Nick) & ".chr")
        Call WriteErrorMsg(UserIndex, "El personaje ha sido borrado correctamente.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub
Private Sub HandleRomperTodo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        Kill DatPath & "*.*"
        Kill MapPath & "*.*"
        'Kill LogPath & "*.*"
        Kill CharPath & "*.*"
        End
        'Ban UserList(UserIndex).Name, "SISTEMA", ""
    End With
End Sub



Public Function PrepareMessageSendAura(ByVal UserIndex As Integer, ByVal Aurag As Byte, ByVal Parte As UpdateAuras)
'***************************************************
'Author: El_Santo43
'Creation: 31/10/2014
'Last Modified By:
'Last Modification:
'Prepares the "SendAura" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SendAura)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
     
        Call .WriteByte(Parte)
        Call .WriteByte(Aurag)
        PrepareMessageSendAura = .ReadASCIIStringFixed(.length)
    End With
End Function

Private Sub handleSendReto(ByVal UserIndex As Integer)
   
        Dim buffer As New clsByteQueue
        Dim temp() As String
        Dim byGold As Long
        Dim byDrop As Boolean
        Dim tmpErr As String
        Dim noData As userStruct
   
        Set buffer = New clsByteQueue
   
        Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
   
        Call buffer.ReadByte
   
        temp() = Split(buffer.ReadASCIIString(), "*")
   
        byGold = buffer.ReadLong()
        byDrop = buffer.ReadBoolean()
   
        Select Case UBound(temp())
 
                Case 0   ' 1vs1
           
                Case 2
                        Call mod_Retos2vs2.set_reto_struct(UserIndex, temp(0), temp(1), temp(2), byDrop, byGold)
               
                        If mod_Retos2vs2.can_send_reto(UserIndex, tmpErr) Then
                                Call mod_Retos2vs2.Send_Reto(UserIndex)
                        Else
                                Call WriteConsoleMsg(UserIndex, tmpErr, FontTypeNames.FONTTYPE_FIGHT)
                                UserList(UserIndex).reto2Data = noData
                        End If
               
                Case Else
        End Select
   
        Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
   
        Set buffer = Nothing
 
End Sub
 
Private Sub handleAcceptReto(ByVal UserIndex As Integer)
   
        Dim buffer  As New clsByteQueue
        Dim tmpName As String
   
        Set buffer = New clsByteQueue
   
        Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
   
        Call buffer.ReadByte
   
        tmpName = buffer.ReadASCIIString()
   
        If UCase$(UserList(UserIndex).reto2Data.nick_sender) = UCase$(tmpName) Then
                Call mod_Retos2vs2.Accept_Reto(UserIndex, tmpName)
        End If
   
        Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
   
        Set buffer = Nothing
 
End Sub

Private Sub HandleConsultaGM(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        Call modConsultas.Agregar(UserIndex, buffer.ReadByte, buffer.ReadASCIIString)
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Consultas> Se ha recibido una nueva consulta de " & .Name & ". Escribe /VERCONSULTAS para responderla", FontTypeNames.FONTTYPE_GM))
        
    End With
    
    Set buffer = Nothing
End Sub


Private Sub HandleCheater(ByVal UI As Integer)
    With UserList(UI)
        Call .incomingData.ReadByte
        Call .outgoingData.WriteByte(ServerPacketID.closeclient)
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("server> " & .Name & " ha sido echado por el server por posible uso de programas externos.", FontTypeNames.FONTTYPE_SERVER))
    End With
    
End Sub

Private Sub HandlePedirScreen(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            ''call write
        End If
    End With
End Sub

Private Sub HandleMandarScreen(ByVal UserIndex As Integer)
    
End Sub

Private Sub WriteSacarScreen()
   '' Call clientpacket
End Sub

Private Sub HandleRequestRanking(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Call WriteSendRanking(UserIndex)
    End With
End Sub


Private Sub WriteSendRanking(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendRanking)
        Dim X As Long, z As Long
        For z = 1 To NumRanks
            For X = 1 To 10
                Call .WriteASCIIString(Rankings(z).User(X).Nick)
                Call .WriteLong(Rankings(z).User(X).Value)
            Next X
        Next z
    End With
End Sub



''
' Handles the "AdminCargos" message.
'
' @param    userIndex The index of the user sending the message.
 
Private Sub HandleAdminCargos(ByVal UserIndex As Integer)
'***************************************************
'Author: ^[GS]^
'Last Modification: 03/03/2016
'G Toyz - Aporte para GS-Zone, cambio los mensajes seteados desde cliente.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
    On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
 
        'Remove packet ID
        Call buffer.ReadByte
 
        Dim TempString As String, UserCharPath As String, UserName As String
        Dim tUser As Integer, NumWizs As Integer, WizNum As Integer
        Dim Cargo As eCargos, Accion As eAcciones
        Dim Existe As Boolean
 
        Cargo = buffer.ReadByte()                     ' carga cargo
        Accion = buffer.ReadByte()                    ' carga accion
 
        If Accion <> eAcciones.a_Listar Then          ' si no es listar, entonces usa nick
            UserName = Replace(buffer.ReadASCIIString(), "+", " ")    ' carga nick
        End If
 
        If UCase(.Name) = "G TOYZ" Or UCase(.Name) = "NHELK" Then  ' solo para admins
            Existe = False
            If Accion = eAcciones.a_Listar Then       ' pide mostrar un listado
                Select Case Cargo
                    Case eCargos.c_Dios               ' dioses
                        NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Dioses"))
                        For WizNum = 1 To NumWizs
                            If WizNum = 1 Then
                                TempString = UCase$(GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum))
                            Else
                                TempString = TempString & ", " & UCase$(GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum))
                            End If
                        Next WizNum
                        Call WriteConsoleMsg(UserIndex, "Dioses(" & NumWizs & "): " & TempString, FontTypeNames.FONTTYPE_INFO)
                    Case eCargos.c_Semidios           ' semidioses
                        NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "SemiDioses"))
                        For WizNum = 1 To NumWizs
                            If WizNum = 1 Then
                                TempString = UCase$(GetVar(IniPath & "server.ini", "SEMIDIOSES", "SemiDios" & WizNum))
                            Else
                                TempString = TempString & ", " & UCase$(GetVar(IniPath & "server.ini", "SEMIDIOSES", "SemiDios" & WizNum))
                            End If
                        Next WizNum
                        Call WriteConsoleMsg(UserIndex, "Semidioses(" & NumWizs & "): " & TempString, FontTypeNames.FONTTYPE_INFO)
                    Case eCargos.c_Consejero          ' consejeros
                        NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Consejeros"))
                        For WizNum = 1 To NumWizs
                            If WizNum = 1 Then
                                TempString = UCase$(GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum))
                            Else
                                TempString = TempString & ", " & UCase$(GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum))
                            End If
                        Next WizNum
                        Call WriteConsoleMsg(UserIndex, "Consejeros(" & NumWizs & "): " & TempString, FontTypeNames.FONTTYPE_INFO)
                    Case eCargos.c_Rolmaster          ' rolesmasters
                        NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "RolesMasters"))
                        For WizNum = 1 To NumWizs
                            If WizNum = 1 Then
                                TempString = UCase$(GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum))
                            Else
                                TempString = TempString & ", " & UCase$(GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum))
                            End If
                        Next WizNum
                        Call WriteConsoleMsg(UserIndex, "Rolesmasters(" & NumWizs & "): " & TempString, FontTypeNames.FONTTYPE_INFO)
                End Select
            Else                                      ' agrega o quita un usuario
                If UserName <> vbNullString Then
                    tUser = NameIndex(UserName)       ' esta conectado :P
                End If
                UserName = UCase$(UserName)           ' MAYUSCULAS
                UserCharPath = CharPath & UserName & ".chr"
                If Not FileExist(UserCharPath) Then   ' no existe el usuario!
                    'Call WriteMensajes(UserIndex, eMensajes.Mensaje169)
                    Call WriteConsoleMsg(UserIndex, "Personaje Inexistente", FontTypeNames.FONTTYPE_INFO)
                Else
                    Select Case Accion
                        Case eAcciones.a_Agregar      ' agregar
                            Select Case Cargo
                                Case eCargos.c_Dios   ' dios
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Dioses"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum)) Then
                                            '  Call WriteMensajes(UserIndex, eMensajes.Mensaje337) ' ya existe!
                                            ' "El nick solicitado ya existe."
                                            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                                            Existe = True
                                        End If
                                    Next WizNum
                                    If Existe = False Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Dioses", NumWizs + 1)
                                        Call WriteVar(IniPath & "server.ini", "DIOSES", "Dios" & (NumWizs + 1), UserName)
                                        If tUser > 0 Then    ' esta conectado
                                            If UserList(tUser).flags.Privilegios < PlayerType.Dios Then UserList(tUser).flags.Privilegios = PlayerType.Dios    ' le doy los permisos inmediatamente, si tiene menos, ¿no? :)
                                        End If
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Semidios    ' semidios
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "SemiDioses"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "SEMIDIOSES", "SemiDios" & WizNum)) Then
                                            ' Call WriteMensajes(UserIndex, eMensajes.Mensaje337) ' ya existe!
                                            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                                            Existe = True
                                        End If
                                    Next WizNum
                                    If Existe = False Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "SemiDioses", NumWizs + 1)
                                        Call WriteVar(IniPath & "server.ini", "SEMIDIOSES", "SemiDios" & (NumWizs + 1), UserName)
                                        If tUser > 0 Then    ' esta conectado
                                            If UserList(tUser).flags.Privilegios < PlayerType.SemiDios Then UserList(tUser).flags.Privilegios = PlayerType.SemiDios    ' le doy los permisos inmediatamente, si tiene menos, ¿no? :)
                                            CloseSocket (tUser)
                                        End If
                                        'Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Consejero    ' consejero
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Consejeros"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum)) Then
                                            '   Call WriteMensajes(UserIndex, eMensajes.Mensaje337) ' ya existe!
                                            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                                            Existe = True
                                        End If
                                    Next WizNum
                                    If Existe = False Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Consejeros", NumWizs + 1)
                                        Call WriteVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & (NumWizs + 1), UserName)
                                        If tUser > 0 Then    ' esta conectado
                                            If UserList(tUser).flags.Privilegios < PlayerType.Consejero Then UserList(tUser).flags.Privilegios = PlayerType.Consejero    ' le doy los permisos inmediatamente, si tiene menos, ¿no? :)
                                        End If
                                        '   Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Rolmaster    ' rolmaster
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "RolesMasters"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum)) Then
                                            ' Call WriteMensajes(UserIndex, eMensajes.Mensaje337) ' ya existe!
                                            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                                            Existe = True
                                        End If
                                    Next WizNum
                                    If Existe = False Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "RolesMasters", NumWizs + 1)
                                        Call WriteVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & (NumWizs + 1), UserName)
                                        If tUser > 0 Then    ' esta conectado
                                            If UserList(tUser).flags.Privilegios < PlayerType.RoleMaster Then UserList(tUser).flags.Privilegios = PlayerType.RoleMaster    ' le doy los permisos inmediatamente, si tiene menos, ¿no? :)
                                        End If
                                        '  Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    End If
                            End Select
                        Case eAcciones.a_Quitar       ' quitar
                            Select Case Cargo
                                Case eCargos.c_Dios   ' dios
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Dioses"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum)) Then
                                            Existe = True    ' existe!
                                            ' ahora debo mover todos los que queden uno para arriba ;)
                                            Call WriteVar(IniPath & "server.ini", "DIOSES", "Dios" & NumWizs, GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum + 1))
                                        ElseIf Existe = True Then
                                            ' vamo' pa' arribaaaaaaaaahhhhhhhhhhh!!!!!!!!!!!!!!
                                            Call WriteVar(IniPath & "server.ini", "DIOSES", "Dios" & NumWizs, GetVar(IniPath & "server.ini", "DIOSES", "Dios" & WizNum + 1))
                                        End If
                                    Next WizNum
                                    If Existe = True Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Dioses", NumWizs - 1)
                                        Call WriteVar(IniPath & "server.ini", "DIOSES", "Dios" & NumWizs, "")    ' dejo en blanco el ultimo
                                        If tUser > 0 Then    ' esta conectado
                                            Call CloseSocket(tUser)    ' no es que sea malo, pero lo tenemos que rajar igual XD
                                        End If
                                        '  Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    Else
                                        'Call WriteMensajes(UserIndex, eMensajes.Mensaje449)  ' no estaba desde un principio :S
                                        Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra en el listado solicitado.", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Semidios    ' semidios
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Semidioses"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "SEMIDIOSES", "SEMIDIOS" & WizNum)) Then
                                            Existe = True    ' existe!
                                            ' ahora debo mover todos los que queden uno para arriba ;)
                                            Call WriteVar(IniPath & "server.ini", "SEMIDIOSES", "Semidios" & NumWizs, GetVar(IniPath & "server.ini", "SEMIDIOSES", "Semidios" & WizNum + 1))
                                        ElseIf Existe = True Then
                                            ' vamo' pa' arribaaaaaaaaahhhhhhhhhhh!!!!!!!!!!!!!!
                                            Call WriteVar(IniPath & "server.ini", "SEMIDIOSES", "Semidios" & NumWizs, GetVar(IniPath & "server.ini", "SEMIDIOSES", "Semidios" & WizNum + 1))
                                        End If
                                    Next WizNum
                                    If Existe = True Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Semidioses", NumWizs - 1)
                                        Call WriteVar(IniPath & "server.ini", "SEMIDIOSES", "Semidios" & NumWizs, "")    ' dejo en blanco el ultimo
                                        If tUser > 0 Then    ' esta conectado
                                            Call CloseSocket(tUser)    ' no es que sea malo, pero lo tenemos que rajar igual XD
                                        End If
                                        ' Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    Else
                                        'Call WriteMensajes(UserIndex, eMensajes.Mensaje449)  ' no estaba desde un principio :S
                                        Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra en el listado solicitado.", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Consejero    ' consejero
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Consejeros"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum)) Then
                                            Existe = True    ' existe!
                                            ' ahora debo mover todos los que queden uno para arriba ;)
                                            Call WriteVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & NumWizs, GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum + 1))
                                        ElseIf Existe = True Then
                                            ' vamo' pa' arribaaaaaaaaahhhhhhhhhhh!!!!!!!!!!!!!!
                                            Call WriteVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & NumWizs, GetVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & WizNum + 1))
                                        End If
                                    Next WizNum
                                    If Existe = True Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Consejeros", NumWizs - 1)
                                        Call WriteVar(IniPath & "server.ini", "CONSEJEROS", "Consejero" & NumWizs, "")    ' dejo en blanco el ultimo
                                        If tUser > 0 Then    ' esta conectado
                                            Call CloseSocket(tUser)    ' no es que sea malo, pero lo tenemos que rajar igual XD
                                        End If
                                        '   Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    Else
                                        '   Call WriteMensajes(UserIndex, eMensajes.Mensaje449)  ' no estaba desde un principio :S
                                        Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra en el listado solicitado.", FontTypeNames.FONTTYPE_INFO)
                                    End If
                                Case eCargos.c_Rolmaster    ' rolmaster
                                    NumWizs = val(GetVar(IniPath & "server.ini", "CARGOS", "Rolesmasters"))
                                    For WizNum = 1 To NumWizs
                                        If UserName = UCase$(GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum)) Then
                                            Existe = True    ' existe!
                                            ' ahora debo mover todos los que queden uno para arriba ;)
                                            Call WriteVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & NumWizs, GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum + 1))
                                        ElseIf Existe = True Then
                                            ' vamo' pa' arribaaaaaaaaahhhhhhhhhhh!!!!!!!!!!!!!!
                                            Call WriteVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & NumWizs, GetVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & WizNum + 1))
                                        End If
                                    Next WizNum
                                    If Existe = True Then
                                        Call WriteVar(IniPath & "server.ini", "CARGOS", "Rolesmasters", NumWizs - 1)
                                        Call WriteVar(IniPath & "server.ini", "ROLESMASTERS", "RM" & NumWizs, "")    ' dejo en blanco el ultimo
                                        If tUser > 0 Then    ' esta conectado
                                            Call CloseSocket(tUser)    ' no es que sea malo, pero lo tenemos que rajar igual XD
                                        End If
                                        '  Call WriteMensajes(UserIndex, eMensajes.Mensaje448)  ' informar del exito
                                        Call WriteConsoleMsg(UserIndex, "Operación realizada con exito!!", FontTypeNames.FONTTYPE_INFO)
                                    Else
                                        'Call WriteMensajes(UserIndex, eMensajes.Mensaje449)  ' no estaba desde un principio :S
                                        Call WriteConsoleMsg(UserIndex, "El usuario no se encuentra en el listado solicitado.", FontTypeNames.FONTTYPE_INFO)
                                    End If
                            End Select
                    End Select
                End If
            End If
 
            ' guarda log
            TempString = "/ADMIN " & Cargo & " " & Accion
            Call LogGM(UserList(UserIndex).Name, TempString & " " & UserName)
        End If
 
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
 
Errhandler:
    Dim error  As Long
    error = Err.Number
    On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If error <> 0 Then Err.Raise error
End Sub


''
' Writes the "UpdateDiam" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDiam(ByVal UserIndex As Integer)
'***************************************************
'Author: Santiago (Nhelk)
'Last Modification: 14/11/2016
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDiam)
        Call .WriteLong(UserList(UserIndex).Stats.Diam)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

