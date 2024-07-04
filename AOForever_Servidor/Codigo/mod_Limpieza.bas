Attribute VB_Name = "mod_Limpieza"
Option Explicit
Private Const MAX_ITEMS_SUELO As Integer = 1500 '// MÁXIMA CANTIDAD DE ITEMS EN EL SUELO
Public Limpieza_Obj(1 To MAX_ITEMS_SUELO) As WorldPos
Public Limpieza_LastObj As Long

Public Sub AgregarObj(ByVal map As Integer, ByVal x As Byte, ByVal Y As Byte)
On Error GoTo errh
1    Limpieza_LastObj = Limpieza_LastObj + 1 '// NUEVO ITEM EN EL SUELO
2    '//EN EL NUEVO ITEM EN EL SUELO GUARDAMOS LAS COORDENADAS:
3    Limpieza_Obj(Limpieza_LastObj).map = map
4    Limpieza_Obj(Limpieza_LastObj).x = x
5    Limpieza_Obj(Limpieza_LastObj).Y = Y
6    MapData(map, x, Y).limpSlot = Limpieza_LastObj
7    If Limpieza_LastObj = (MAX_ITEMS_SUELO - 100) Then _
        Call LimpiarMundo '// SI LLEGA AL MÁXIMO LIMPIAMOS
        
        Exit Sub
errh:
    Call LogError("Error en 'AgregarObj' en mod_Limpieza.bas en linea " & Err.line)
    
    
End Sub

Public Sub QuitarObj(ByVal map As Integer, ByVal x As Byte, ByVal Y As Byte)
1 On Error GoTo errh
2    With MapData(map, x, Y)
3        If .limpSlot > 0 And .limpSlot <= Limpieza_LastObj Then
4            If Limpieza_Obj(.limpSlot).map = map And Limpieza_Obj(.limpSlot).x = x And Limpieza_Obj(.limpSlot).Y = Y Then
5                Limpieza_Obj(.limpSlot).map = 0
6                MapData(map, x, Y).limpSlot = 0
7            End If
8        End If
9    End With
Exit Sub
errh:
    Call LogError("Error en 'QuitarObj' en mod_Limpieza.bas en linea " & Err.line)
End Sub

Public Sub LimpiarMapa(ByVal map As Integer)
On Error GoTo errh
    If Limpieza_LastObj > 0 Then
        Dim x As Long
        For x = 1 To Limpieza_LastObj
            With Limpieza_Obj(x)
                If .map = map Then
                    Call EraseObj(MapData(.map, .x, .Y).ObjInfo.Amount, .map, .x, .Y)
                End If
            End With
            '// LIMPIAMOS
        Next x
    End If
    Exit Sub
errh:
    Call LogError("Error en 'LimpiarMapa' en mod_Limpieza.bas")
End Sub

Public Sub LimpiarMundo()
On Error GoTo errh
1    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando mundo", FontTypeNames.FONTTYPE_SERVER))
    
2    If Limpieza_LastObj > 0 Then
3        Dim x As Long
4        For x = 1 To Limpieza_LastObj
5            With Limpieza_Obj(x)
6                If .map <> 0 Then
7                    Call EraseObj(MapData(.map, .x, .Y).ObjInfo.Amount, .map, .x, .Y)
8                End If
9            End With
            '// LIMPIAMOS
10           Limpieza_Obj(x).map = 0
11           Limpieza_Obj(x).x = 0
12           Limpieza_Obj(x).Y = 0
13        Next x
14    End If
15    Limpieza_LastObj = 0 '// AHORA HAY 0 ITEMS EN EL SUELO.
16    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza finalizada", FontTypeNames.FONTTYPE_SERVER))
Exit Sub
errh:
    Call LogError("Error en 'LimpiarMundo' en mod_Limpieza.bas en linea " & Err.line)
End Sub

