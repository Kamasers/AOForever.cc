Attribute VB_Name = "mod_DragDrop"
Option Explicit

Public Sub DragToUser(ByVal UserIndex As Integer, _
                          ByVal tIndex As Integer, _
                          ByVal Slot As Byte, _
                          ByVal Amount As Integer)
     

     
            Dim tObj    As Obj
            Dim tString As String
            Dim Espacio As Boolean
            
            If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then
                Amount = UserList(UserIndex).Invent.Object(Slot).Amount
                'Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                'Exit Sub
            End If
            
            If UserIndex = tIndex Then Exit Sub
            
            'Preparo el objeto.
           tObj.Amount = Amount
           tObj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            
            If tObj.ObjIndex = 0 Then Exit Sub
            
           Espacio = MeterItemEnInventario(tIndex, tObj)
     
           'No tiene espacio.
     
            If Not Espacio Then
            WriteConsoleMsg UserIndex, "El usuario no tiene espacio en su inventario.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
            End If
     
            'Quito el objeto.
           QuitarUserInvItem UserIndex, Slot, Amount
     
           'Hago un update de su inventario.
            UpdateUserInv False, UserIndex, Slot
     
            'Preparo el mensaje para userINdex (quien dragea)
     
           tString = "Le has arrojado"
     
           If tObj.Amount <> 1 Then
                   tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.ObjIndex).name
           Else
                   tString = tString & " tu " & ObjData(tObj.ObjIndex).name
           End If
     
           tString = tString & " a " & UserList(tIndex).name
     
           'Envio el mensaje
            WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO
     
            'Preparo el mensaje para el otro usuario (quien recibe)
           tString = UserList(UserIndex).name & " te ha arrojado"
     
           If tObj.Amount <> 1 Then
                   tString = tString & " " & tObj.Amount & " - " & ObjData(tObj.ObjIndex).name
           Else
                   tString = tString & " su " & ObjData(tObj.ObjIndex).name
           End If
     
           'Envio el mensaje al otro usuario
            WriteConsoleMsg tIndex, tString, FontTypeNames.FONTTYPE_INFO
     
End Sub
    
Public Sub DragToNPC(ByVal UserIndex As Integer, _
                         ByVal tNPC As Integer, _
                         ByVal Slot As Byte, _
                         ByVal Amount As Integer)
     
            ' @ Author : maTih.-
           '            Drag un slot a un npc.
     
            On Error GoTo Errhandler
     
            Dim TeniaOro As Long
            Dim teniaObj As Integer
            Dim tmpIndex As Integer
            
            If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then
                Amount = UserList(UserIndex).Invent.Object(Slot).Amount
                'Exit Sub
            End If


     
            tmpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            TeniaOro = UserList(UserIndex).Stats.GLD
            teniaObj = UserList(UserIndex).Invent.Object(Slot).Amount
     
            'Es un banquero?
     
           If Npclist(tNPC).NPCtype = eNPCType.Banquero Then
                   Call UserDejaObj(UserIndex, Slot, Amount)
                   'No tiene más el mismo amount que antes? entonces depositó.
     
                    If teniaObj <> UserList(UserIndex).Invent.Object(Slot).Amount Then
                            WriteConsoleMsg UserIndex, "Has depositado un objeto.", FontTypeNames.FONTTYPE_INFO ' & Amount & " - " & ObjData(tmpIndex).name,
                            UpdateUserInv False, UserIndex, Slot
                    End If
     
                    'Es un npc comerciante?
           ElseIf Npclist(tNPC).Comercia = 1 Then
                   'El npc compra cualquier tipo de items?
     
                    If Not Npclist(tNPC).TipoItems <> eOBJType.otCualquiera Or Npclist(tNPC).TipoItems = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType Then
                            Call Comercio(eModoComercio.Venta, UserIndex, tNPC, Slot, Amount)
                            'Ganó oro? si es así es porque lo vendió.
     
                           If TeniaOro <> UserList(UserIndex).Stats.GLD Then
                                   WriteConsoleMsg UserIndex, "Le has vendido (" & Amount & " - " & ObjData(tmpIndex).name & ") a " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO
                           End If
     
                   Else
                           WriteConsoleMsg UserIndex, "El npc no está interesado en comprar este tipo de objetos.", FontTypeNames.FONTTYPE_INFO
                   End If
            ElseIf Npclist(tNPC).Comercia = 0 And Npclist(tNPC).NPCtype <> eNPCType.Banquero Then
                WriteConsoleMsg UserIndex, "El npc no esta interesado en comerciar contigo.", FontTypeNames.FONTTYPE_INFO
           End If
     
           Exit Sub
     
Errhandler:
     
    End Sub
Public Sub DragToPos(ByVal UserIndex As Integer, _
                        ByVal X As Byte, _
                        ByVal Y As Byte, _
                        ByVal Slot As Byte, _
                        ByVal Amount As Integer)
     
         
     
           Dim errorFound As String
           Dim tObj       As Obj
           Dim tString    As String
     
           'No puede dragear en esa pos?
     
            If Not CanDragToPos(UserList(UserIndex).Pos.map, X, Y, errorFound) Then
                    WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFO
     
                    Exit Sub
     
            End If
            
            If Not CanDragObj(UserList(UserIndex).Invent.Object(Slot).ObjIndex, (UserList(UserIndex).flags.Navegando = 1), errorFound, UserIndex) Then
                WriteConsoleMsg UserIndex, errorFound, FontTypeNames.FONTTYPE_INFO
     
                    Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then
                Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then Exit Sub
            
            'Creo el objeto.
           tObj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
           tObj.Amount = Amount
           Dim tmpPos As WorldPos
           tmpPos.map = UserList(UserIndex).Pos.map
           tmpPos.X = X
           tmpPos.Y = Y
           
           If Distancia(UserList(UserIndex).Pos, tmpPos) > 9 Then
                Call WriteConsoleMsg(UserIndex, "¡Lanzas imprecisamente!", FontTypeNames.FONTTYPE_INFO)
                Dim Err As String
                
                X = RandomNumber(X - 3, X + 3)
                Y = RandomNumber(Y - 3, Y + 3)
                Do While (CanDragToPos(UserList(UserIndex).Pos.map, X, Y, Err) = False)
                    X = RandomNumber(X - 3, X + 3)
                    Y = RandomNumber(Y - 3, Y + 3)
                Loop
                            
            End If
           'Agrego el objeto a la posición.
            MakeObj tObj, UserList(UserIndex).Pos.map, CInt(X), CInt(Y)
     
            'Quito el objeto.
           QuitarUserInvItem UserIndex, Slot, Amount
     
           'Actualizo el inventario
            UpdateUserInv False, UserIndex, Slot
     
            'Preparo el mensaje.
           tString = "Has arrojado "

                   tString = tString & tObj.Amount & " - " & ObjData(tObj.ObjIndex).name

     
            'ENvio.
           WriteConsoleMsg UserIndex, tString, FontTypeNames.FONTTYPE_INFO
     
    End Sub
     
    Private Function CanDragToPos(ByVal map As Integer, _
                                 ByVal X As Byte, _
                                 ByVal Y As Byte, _
                                 ByRef error As String) As Boolean
     
      
           CanDragToPos = False
     
           'Zona segura?
     
            If Not MapInfo(map).Pk Then
                    error = "No está permitido arrojar objetos al suelo en zonas seguras."
                    Exit Function
            End If
     
            'Ya hay objeto?
     
           If Not MapData(map, X, Y).ObjInfo.ObjIndex = 0 Then
                error = "¡Hay un objeto en esa posición!"
                Exit Function
           End If
     
           'Tile bloqueado?
     
            If Not MapData(map, X, Y).Blocked = 0 Then
                    error = "No puedes arrojar objetos en esa posición"
                    Exit Function
     
            End If
           
            If HayAgua(map, X, Y) Then
                    error = "No puedes arrojar objetos al agua"
                    Exit Function
            End If
     
            CanDragToPos = True
     
    End Function
     
Private Function CanDragObj(ByVal ObjIndex As Integer, _
                                ByVal Navegando As Boolean, _
                                ByRef error As String, ByVal UserIndex As Integer) As Boolean
    CanDragObj = False
    
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData()) Then Exit Function
    
    'Objeto newbie?
    If ObjData(ObjIndex).Newbie <> 0 Then
        error = "No puedes arrojar objetos newbies!"
        Exit Function
    End If
    
    'Está navgeando?
    If Navegando And UserList(UserIndex).Invent.BarcoObjIndex = ObjIndex Then
        error = "No puedes arrojar un barco si estás navegando!"
        Exit Function
    End If
    CanDragObj = True
     
End Function

Public Sub HandleDragInventory(ByVal UserIndex As Integer)
     
        On Error GoTo errcito
     
            Dim ObjSlot1   As Byte
            Dim ObjSlot2   As Byte
     
            Dim tmpUserObj As UserOBJ
     
            If UserList(UserIndex).incomingData.length < 3 Then
                    Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
     
                    Exit Sub
     
            End If
     
            With UserList(UserIndex)
           
                    'Leemos el paquete
                   Call .incomingData.ReadByte
         
                   ObjSlot1 = .incomingData.ReadByte
                   ObjSlot2 = .incomingData.ReadByte
                   
                   If ObjSlot2 > .CurrentInventorySlots Or ObjSlot1 > .CurrentInventorySlots Then
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Comerciando Then Exit Sub
                    ''If UserList(UserIndex).flags.Comerciando Then Exit Sub
                   'Cambiamos si alguno es un anillo
     
                    If .Invent.AnilloEqpSlot = ObjSlot1 Then
                            .Invent.AnilloEqpSlot = ObjSlot2
                    ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
                            .Invent.AnilloEqpSlot = ObjSlot1
                    End If
                    
                    If .Invent.AnilloEqpSlot2 = ObjSlot1 Then
                            .Invent.AnilloEqpSlot2 = ObjSlot2
                    ElseIf .Invent.AnilloEqpSlot2 = ObjSlot2 Then
                            .Invent.AnilloEqpSlot2 = ObjSlot1
                    End If
                    
                    'Cambiamos si alguno es un armor
     
                   If .Invent.ArmourEqpSlot = ObjSlot1 Then
                           .Invent.ArmourEqpSlot = ObjSlot2
                   ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
                           .Invent.ArmourEqpSlot = ObjSlot1
                   End If
         
                   'Cambiamos si alguno es un barco
     
                    If .Invent.BarcoSlot = ObjSlot1 Then
                            .Invent.BarcoSlot = ObjSlot2
                    ElseIf .Invent.BarcoSlot = ObjSlot2 Then
                            .Invent.BarcoSlot = ObjSlot1
                    End If
           
                    'Cambiamos si alguno es un casco
     
                   If .Invent.CascoEqpSlot = ObjSlot1 Then
                           .Invent.CascoEqpSlot = ObjSlot2
                   ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
                           .Invent.CascoEqpSlot = ObjSlot1
                   End If
         
                   'Cambiamos si alguno es un escudo
     
                    If .Invent.EscudoEqpSlot = ObjSlot1 Then
                            .Invent.EscudoEqpSlot = ObjSlot2
                    ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
                            .Invent.EscudoEqpSlot = ObjSlot1
                    End If
           
                    'Cambiamos si alguno es munición
     
                   If .Invent.MunicionEqpSlot = ObjSlot1 Then
                           .Invent.MunicionEqpSlot = ObjSlot2
                   ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
                           .Invent.MunicionEqpSlot = ObjSlot1
                   End If
         
                   'Cambiamos si alguno es un arma
     
                    If .Invent.WeaponEqpSlot = ObjSlot1 Then
                            .Invent.WeaponEqpSlot = ObjSlot2
                    ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
                            .Invent.WeaponEqpSlot = ObjSlot1
                    End If
           
                    'Hacemos el intercambio propiamente dicho
                   tmpUserObj = .Invent.Object(ObjSlot1)
                   .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
                   .Invent.Object(ObjSlot2) = tmpUserObj
     
                   'Actualizamos los 2 slots que cambiamos solamente
                    Call UpdateUserInv(False, UserIndex, ObjSlot1)
                    Call UpdateUserInv(False, UserIndex, ObjSlot2)
                    If UserList(UserIndex).flags.Comerciando = True Then
                        Call WriteTradeOK(UserIndex)
                        Call UpdateVentanaBanco(UserIndex)
                    End If
            End With
Exit Sub
errcito:
     Debug.Print Err.Number & " " & Err.description
End Sub
Public Sub HandleDragToPos(ByVal UserIndex As Integer)
     
         
     
            Dim X      As Byte
            Dim Y      As Byte
            Dim Slot   As Byte
            Dim Amount As Integer
     
            Call UserList(UserIndex).incomingData.ReadByte
     
            X = UserList(UserIndex).incomingData.ReadByte()
            Y = UserList(UserIndex).incomingData.ReadByte()
            Slot = UserList(UserIndex).incomingData.ReadByte()
            Amount = UserList(UserIndex).incomingData.ReadInteger()
     
    If Slot <= 0 Then Exit Sub
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then Exit Sub
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then
       Amount = UserList(UserIndex).Invent.Object(Slot).Amount
    End If
    
    If Amount <= 0 Then
        Call WriteConsoleMsg(UserIndex, "Cantidad inválida.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim OtroUserIndex As Integer
    If UserList(UserIndex).flags.Comerciando Then
        OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
        If OtroUserIndex > 0 And OtroUserIndex <= maxusers Then
             Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes usar drag & drop mientras comercias!!", FontTypeNames.FONTTYPE_TALK)
             Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
             Call LimpiarComercioSeguro(UserIndex)
             Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
    With MapData(UserList(UserIndex).Pos.map, X, Y)
        
        If .NpcIndex <> 0 Then
            mod_DragDrop.DragToNPC UserIndex, .NpcIndex, Slot, Amount
        ElseIf .UserIndex <> 0 Then
            mod_DragDrop.DragToUser UserIndex, .UserIndex, Slot, Amount
        Else
            mod_DragDrop.DragToPos UserIndex, X, Y, Slot, Amount
        End If
    End With
End Sub

