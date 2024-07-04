Attribute VB_Name = "mod_Party"
Option Explicit



Private Type tUserParty
    UserIndex As Integer
    Porcentaje As Byte
    ExpAcumulada As Long
End Type

Public Const MaxPartys As Integer = 300
Public Const MaxUsuariosParty As Byte = 5
Private Type tPartys
    User(1 To MaxUsuariosParty) As tUserParty
    Lider As Integer
    Activa As Boolean
    Solicitudes(20) As Integer
End Type
Public Party(1 To 300) As tPartys
Private Const ExpAlDisolver As Boolean = True
Private Function cantUsers(ByVal pIndex As Integer) As Byte
    Dim X As Long
    If pIndex <= 0 Then Exit Function
    With Party(pIndex)
        For X = 1 To MaxUsuariosParty
            If .User(X).UserIndex > 0 Then
                cantUsers = cantUsers + 1
            End If
        Next X
    End With
    Exit Function
errh:
    Call LogError("Error en sub cantUsers de mod_party: " & Err.Number & " - " & Err.description)
    
End Function
Public Sub EcharParty(ByVal pIndex As Integer, ByVal UserIndex As Integer, Optional ByVal Echado As Boolean = True)
On Error GoTo errh
    If pIndex <= 0 Then Exit Sub
    If UserIndex <= 0 Then Exit Sub
    With Party(pIndex)
        If Echado = False Then
            Call MensajeParty(pIndex, "El usuario " & UserList(UserIndex).Name & " ha salido de la party")
        Else
            Call MensajeParty(pIndex, "El usuario " & UserList(UserIndex).Name & " ha sido echado de la party")
        End If
        Dim X As Long
        Dim Y As Long
        
        
        For X = 1 To MaxUsuariosParty
            If .User(X).UserIndex = UserIndex Then
                If ExpAlDisolver Then
                    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + .User(X).ExpAcumulada
                    Call CheckUserLevel(UserIndex)
                End If
                UserList(UserIndex).PartyIndex = 0
                .User(X).UserIndex = 0

                Dim cpers As Byte
                cpers = .User(X).Porcentaje
                .User(1).Porcentaje = .User(1).Porcentaje + cpers
                .User(X).Porcentaje = 0
                If Echado Then
                    Call WriteConsoleMsg(UserIndex, "Party> Has sido echado de la party", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "Party> Has salido de la party", FontTypeNames.FONTTYPE_PARTY)
                End If
                WriteCerrarPartyForm UserIndex
            End If
        Next X
        WritePartyForm .Lider
    End With
        Exit Sub
errh:
    Call LogError("Error en sub EcharParty: " & Err.Number & " - " & Err.description)
    
End Sub

Public Sub DisolverParty(ByVal pIndex As Integer)
On Error GoTo errh
    If pIndex <= 0 Then Exit Sub
    Call MensajeParty(pIndex, "La party ha sido disuelta por el lider de la misma")
    Dim X As Long
    For X = 1 To MaxUsuariosParty
        If ExpAlDisolver Then
            If Party(pIndex).User(X).UserIndex > 0 Then
                UserList(Party(pIndex).User(X).UserIndex).Stats.Exp = UserList(Party(pIndex).User(X).UserIndex).Stats.Exp + Party(pIndex).User(X).ExpAcumulada
                Call CheckUserLevel(Party(pIndex).User(X).UserIndex)
                WriteCerrarPartyForm Party(pIndex).User(X).UserIndex
            End If
        End If
    Next X
    Call LimpiarParty(pIndex, True)
        Exit Sub
errh:
    Call LogError("Error en sub DisolverParty: " & Err.Number & " - " & Err.description)
    
End Sub

Sub DeslogeaParty(ByVal UserIndex As Integer)
On Error GoTo errh
    Dim nParty As Integer
    With UserList(UserIndex)
        nParty = .PartyIndex
        If nParty <= 0 Then Exit Sub
        If Party(nParty).Lider = UserIndex Then
            Call DisolverParty(nParty)
        Else
            Call EcharParty(nParty, UserIndex, False)
        End If
            
    End With
        Exit Sub
errh:
    Call LogError("Error en sub DeslogeaParty: " & Err.Number & " - " & Err.description)
    
End Sub

Public Sub LimpiarParty(ByVal pIndex As Integer, Optional ByVal LimpiarPindex As Boolean = False)
On Error GoTo errh
    If pIndex <= 0 Then Exit Sub
    With Party(pIndex)
        
        .Activa = False
        .Lider = 0
        Dim X As Long
        For X = 1 To MaxUsuariosParty
            .User(X).Porcentaje = 0
            If LimpiarPindex Then
                If .User(X).UserIndex > 0 Then
                    UserList(.User(X).UserIndex).PartyIndex = 0
                End If
            .User(X).UserIndex = 0
            End If
        Next X
        For X = 0 To 20
            .Solicitudes(X) = 0
        Next X
    End With
        Exit Sub
errh:
    Call LogError("Error en sub LimpiarParty: " & Err.Number & " - " & Err.description)
    
End Sub
Public Function PorcPartyValidos(ByVal pIndex As Integer) As Boolean
On Error GoTo errh
    If pIndex <= 0 Then Exit Function
    With Party(pIndex)
        Dim X As Long, Count As Byte, Y As Long, maxPorc As Byte
        For X = 1 To MaxUsuariosParty
            If .User(X).UserIndex > 0 Then
                Count = Count + .User(X).Porcentaje
                If .User(X).Porcentaje > maxPorc Then maxPorc = .User(X).Porcentaje
            End If
        Next X
        If maxPorc > UserList(.Lider).Stats.UserSkills(eSkill.Liderazgo) Then
            PorcPartyValidos = False
            Exit Function
        End If
        If Count > 100 Then
            PorcPartyValidos = False
        Else
            PorcPartyValidos = True
        End If
        
        Exit Function
        For X = 1 To MaxUsuariosParty
            If .User(X).UserIndex > 0 Then
                If .User(X).Porcentaje < 10 Then
                    For Y = 1 To MaxUsuariosParty
                        If .User(Y).Porcentaje >= 20 Then
                            .User(Y).Porcentaje = .User(Y).Porcentaje - 10
                            .User(X).Porcentaje = .User(X).Porcentaje + 10
                            Exit For
                        End If
                    Next Y
                End If
            End If
        Next X
        Call PartyInfo(pIndex)
    End With
        Exit Function
errh:
    Call LogError("Error en sub PorcPartyValidos: " & Err.Number & " - " & Err.description)
    
End Function

Public Sub PartyInfo(ByVal pIndex As Integer)
On Error GoTo errh
    If pIndex <= 0 Or pIndex > 300 Then Exit Sub
    With Party(pIndex)
        Dim X As Long, Y As Long
        For Y = 1 To MaxUsuariosParty
            If .User(Y).UserIndex > 0 Then
                For X = 1 To MaxUsuariosParty
                    If .User(X).UserIndex > 0 Then
                        Call WriteConsoleMsg(.User(Y).UserIndex, "Party> " & UserList(.User(X).UserIndex).Name & "(" & .User(X).Porcentaje & "%)" & IIf(.Lider = .User(X).UserIndex, "(LIDER)", "") & IIf(ExpAlDisolver, "(Exp acumulada: " & .User(X).ExpAcumulada & ")", ""), FontTypeNames.FONTTYPE_PARTY)
                    End If
                Next X
            End If
        Next Y
    End With
        Exit Sub
errh:
    Call LogError("Error en sub PartyInfo: " & Err.Number & " - " & Err.description)
    
End Sub

Private Function PuedeCrearParty(ByVal UserIndex As Integer) As String
On Error GoTo errh
    With UserList(UserIndex)
        If .flags.Muerto > 0 Then
            PuedeCrearParty = "Party> ¡Estás muerto!"
            Exit Function
        End If
        If .PartyIndex <> 0 Then
            PuedeCrearParty = "Party> Ya estás en una party"
            Exit Function
        End If
        If .Stats.ELV < 25 Then
            PuedeCrearParty = "Party> No tienes el nivel suficiente para fundar una party"
            Exit Function
        End If
        If .Stats.UserSkills(eSkill.Liderazgo) < 50 Then
            PuedeCrearParty = "Tu carisma y liderazgo no son suficientes para liderar una party."
            Exit Function
        End If
    End With
    Exit Function
errh:
    Call LogError("Error en sub PuedeCrearParty: " & Err.Number & " - " & Err.description)
    
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
On Error GoTo errh
    Dim puede As String
    
    puede = PuedeCrearParty(UserIndex)
    If Len(puede) = 0 Then
        Dim X As Long, BuscarPIndex As Integer
        BuscarPIndex = -1
        For X = 1 To MaxPartys
            If Party(X).Activa = False Then
                BuscarPIndex = X
                Exit For
            End If
        Next X

        If BuscarPIndex = -1 Then
            Call WriteConsoleMsg(UserIndex, "Party> Se ha alcanzado el maximo de partys del servidor.", FontTypeNames.FONTTYPE_PARTY)
            Exit Sub
        End If
        With Party(BuscarPIndex)
            .Activa = True
            .Lider = UserIndex
            .User(1).UserIndex = UserIndex
            .User(1).Porcentaje = 100
            .User(1).ExpAcumulada = 0
        End With
        UserList(UserIndex).PartyIndex = BuscarPIndex
        Call WriteConsoleMsg(UserIndex, "Party> La party se ha creado correctamente", FontTypeNames.FONTTYPE_PARTY)
        Call WritePartyForm(UserIndex)
    Else
        Call WriteConsoleMsg(UserIndex, puede, FontTypeNames.FONTTYPE_PARTY)
    End If
    Exit Sub
errh:
    Call LogError("Error en sub crearParty: " & Err.Number & " - " & Err.description)
    
End Sub

Sub GanaExperiencia(ByVal Exp As Long, ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim pIndex As Integer
        pIndex = .PartyIndex
    End With
    
    With Party(pIndex)
        Dim X As Long, p As Byte, b As Byte
        For X = 1 To MaxUsuariosParty
            With .User(X)
                If .UserIndex > 0 Then
                    
                    p = .Porcentaje
                    b = b + .Porcentaje
                    If UserList(.UserIndex).flags.Muerto = 0 Then
                        If UserList(.UserIndex).Pos.map = UserList(UserIndex).Pos.map Then
                            If Not ExpAlDisolver Then
                                UserList(.UserIndex).Stats.Exp = UserList(.UserIndex).Stats.Exp + Round((Exp / 100) * p)
                                Call WriteConsoleMsg(.UserIndex, "Has ganado " & Round((Exp / 100) * p) & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
                            Else
                                .ExpAcumulada = .ExpAcumulada + Round((Exp / 100) * p)
                            End If
                        End If
                    End If
                End If
            End With
        Next X
        If b > 100 Or b < 100 Then
            Debug.Print "hace algo, los porcentajes tienen algo raro"
        End If
    End With
End Sub

Sub MensajeParty(ByVal pIndex As Integer, ByVal Msg As String)
    Dim X As Long
    If pIndex <= 0 Then Exit Sub
    With Party(pIndex)
        For X = 1 To MaxUsuariosParty
            If .User(X).UserIndex <> 0 Then
                Call WriteConsoleMsg(.User(X).UserIndex, Msg, FontTypeNames.FONTTYPE_PARTY)
            End If
        Next X
    End With
End Sub











