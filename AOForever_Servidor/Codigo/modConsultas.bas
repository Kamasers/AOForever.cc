Attribute VB_Name = "modConsultas"
Option Explicit

Public Enum eTipoConsulta
    Reporte = 1
    Denuncia = 2
    Consulta = 3
    Sugerencia = 4
End Enum

Private Type tConsulta
    Texto As String
    ocupada As Boolean
    Posicion As WorldPos
    Tipo As eTipoConsulta
End Type

Public Const maxConsultas As Byte = 50
Public Consultas(1 To maxConsultas) As tConsulta

Public Sub Agregar(ByVal UserIndex As Integer, ByVal Tipo As eTipoConsulta, ByVal Texto As String)
    Dim nuevoSlot As Byte
    nuevoSlot = BuscarSlot
    If nuevoSlot = 0 Then Exit Sub
    With Consultas(nuevoSlot)
        .ocupada = True
        .Posicion = UserList(UserIndex).Pos
        .Texto = Texto
        .Tipo = Tipo
    End With
End Sub

Private Function BuscarSlot() As Byte
    Dim x As Long
    For x = 1 To maxConsultas
        If Consultas(x).ocupada = False Then
            BuscarSlot = CByte(x)
            Exit Function
        End If
    Next x
    BuscarSlot = 0
End Function







