Attribute VB_Name = "mod_Proyectiles"
Option Explicit

Type tProyectil
    ActualX As Integer
    ActualY As Integer
    GrhIndex As Long
    Usado As Boolean
End Type
    
Public Const ball_speed As Double = 1

Public Function Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Double
    Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
End Function


Public Sub CrearProyectil(ByVal fromCI As Integer, ByVal toCI As Integer, ByVal GrhIndex As Integer)
    With charlist(toCI)
        If Distance(.Pos.X, .Pos.Y, charlist(fromCI).Pos.X, charlist(fromCI).Pos.Y) < 2 Then Exit Sub
        Dim loopc As Long, find As Byte
        For loopc = 1 To 4
            If .Proyectil(loopc).Usado = False Then
                find = CByte(loopc)
                Exit For
            End If
        Next loopc
        If find = 0 Then Exit Sub
        .Proyectil(find).Usado = True
        .Proyectil(find).ActualX = Engine_TPtoSPX(charlist(fromCI).Pos.X) 'Aca necesito sacar la posicion en el render del char FROMCI
        .Proyectil(find).ActualY = Engine_TPtoSPY(charlist(fromCI).Pos.Y)
        .Proyectil(find).GrhIndex = GrhIndex
    End With
End Sub

Public Function GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: [url=http://www.vbgore.com/GameClient.TileEn]http://www.vbgore.com/GameClient.TileEn[/url] ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    GetAngle = (Atn(-GetAngle / Sqr(-GetAngle * GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then GetAngle = 360 - GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    GetAngle = 0
 
Exit Function
 
End Function

