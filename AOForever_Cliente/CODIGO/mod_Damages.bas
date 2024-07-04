Attribute VB_Name = "mod_Damages"
Option Explicit

Type tDamage
    Label As String
    Alpha As Byte
    r As Byte
    g As Byte
    b As Byte
    Using As Boolean
    Wait As Byte
    OffSetY As Integer
    d3dColor As Long
End Type

'Public Damages(250) As tDamage

Public Sub CreateDamage(ByVal Label As String, r As Byte, g As Byte, b As Byte, tX As Byte, tY As Byte)
    Dim nDmg As Byte
    nDmg = NewDamageIndex(tX, tY)
    If nDmg = 9 Then Exit Sub
    With MapData(tX, tY).Damage(nDmg)
        .Label = Label
        .r = r
        .g = g
        .b = b
        .Alpha = 255
        .Using = True
        .Wait = 5
        .OffSetY = 0
        .d3dColor = D3DColorXRGB(.r, .g, .b)
    End With
End Sub

Private Function NewDamageIndex(ByVal tX As Byte, ByVal tY As Byte) As Byte
    Dim X As Long
    For X = 0 To 8
        If MapData(tX, tY).Damage(X).Using = False Then
            NewDamageIndex = X
            Exit Function
        End If
    Next X
    NewDamageIndex = 9
End Function

