Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
Public EnInventario As Boolean
Public tInv As Long
    Public ctimeInv As Long
    Dim OffsetCounterX As Single
    Dim OffsetCounterY As Single
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
 Public Movement_Speed As Single
Public UsaVSync As Boolean
Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" (lpvDest As Any, _
lpvSource As Any, ByVal cbCopy As Long)
'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type
Private Type POINTAPI
    x As Long
    Y As Long
End Type



 
'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Public cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont
 
 Public alphaTecho As Byte
'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

Public FxGrh() As Grh
'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    LastMov As Long
    Rotac As Double
    Movimient As Boolean
    Proyectil(1 To 4) As tProyectil
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iCasco As Integer
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    aura(0 To 4) As tAuras
End Type

'Info de un objeto
Public Type obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Damage(8) As tDamage
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX8 Objects
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

' Directx8 Fonts
Private Type FontInfo
    MainFont As DxVBLibA.D3DXFont
    MainFontDesc As IFont
    MainFontFormat As New StdFont
    Color As Long
End Type: Private font() As FontInfo

Public Type TLVERTEX
    x As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type
Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public fps As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    If Not FileExist(App.path & "\init\Cabezas.ind", vbNormal) Then End
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    ReDim FxGrh(1 To NumFxs) As Grh
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        .iCasco = Casco
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.Y = Y
        
        'Make active
        .Active = 1
        Dim LoopC As Long
        
        For LoopC = 1 To 4
            .Proyectil(LoopC).Usado = False
        Next LoopC
    End With
    
    'Plot on map
    MapData(x, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .Atacable = False
        .FxIndex = 0
        .invisible = False
        
#If SeguridadAlkon Then
        Call MI(CualMI).ResetInvisible(CharIndex)
#End If
        
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.x = 0
        .Pos.Y = 0
        
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    If GrhIndex > UBound(GrhData) Then Exit Sub
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = x + addX
        nY = Y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.Y = nY
        MapData(x, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.x, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.x, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.x, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.Y)
    End If
End Sub
Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        x = .Pos.x
        Y = .Pos.Y
        
        MapData(x, Y).CharIndex = 0
        
        addX = nX - x
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.x = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
        Call DoPasosFx(CharIndex)
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    
    For J = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.x = J
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim LoopC As Long
    Dim Dale As Boolean
    
    LoopC = 1
    Do While charlist(LoopC).Active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & GraphicsFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
            
        End With
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(x, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(x, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.x, UserPos.Y) Then
                    If Not HayAgua(x, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(x, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(x, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, lightvalue() As Long)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        
        'Draw
        Call Device_Textured_Render(x, Y, SurfaceDB.Surface(.FileNum), SourceRect, lightvalue)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Alpha As Boolean, ByVal AlphaB As Byte, lvalue() As Long, Optional ByVal Angle As Single)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth * 0.5) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight

        'Draw
        Call Device_Textured_Render(x, Y, SurfaceDB.Surface(.FileNum), SourceRect, lvalue, Alpha, AlphaB, , , , Angle)
        
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, light_value() As Long, Optional ByVal Alpha As Boolean, Optional ByVal AlphaByte As Byte = 255 _
                           , Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Byte = 0)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
'On Error GoTo error
    If Grh.GrhIndex <= 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            If Grh.Speed > 0 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                    
                    If Grh.Loops <> INFINITE_LOOPS Then
                        If Grh.Loops > 0 Then
                            Grh.Loops = Grh.Loops - 1
                        Else
                            Grh.Started = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then Exit Sub
        
    If Grh.FrameCounter < 1 Then Grh.FrameCounter = 1
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
         Call Device_Textured_Render(x, Y, SurfaceDB.Surface(.FileNum), SourceRect, light_value, Alpha, AlphaByte, Shadow, .sX, .sY, Angle, Grh.GrhIndex)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub


Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(ByVal desthDC As Long, ByVal grh_index As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
 
' / Author: Emanuel Matias 'Dunkan'
' / Note: Dibujar pictures del 'Crear Personaje'
 
'On Error Resume Next
   
    Dim file_path   As String
    Dim src_x       As Integer
    Dim src_y       As Integer
    Dim src_width   As Integer
    Dim src_height  As Integer
    Dim hdcsrc      As Long
    Dim MaskDC      As Long
    Dim PrevObj     As Long
    Dim PrevObj2    As Long
    Dim screen_x    As Integer
    Dim screen_y    As Integer
   
    screen_x = destRect.Left
    screen_y = destRect.Top
   
    If grh_index <= 0 Then Exit Sub
 
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
        file_path = App.path & "\grafs\" & CStr(GrhData(grh_index).FileNum) & ".bmp"
       
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
           
        hdcsrc = CreateCompatibleDC(desthDC)
         
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
       
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
 
        DeleteDC hdcsrc
 
End Sub

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef destRect As RECT, ByVal TransparentColor)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/22/2009
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    Dim Color As Long
    Dim x As Long
    Dim Y As Long
    
    For x = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            Color = GetPixel(srchdc, x, Y)
            
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (x - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), Color)
            End If
        Next Y
    Next x
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub


Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim x           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim lvalue(3) As Long
    Dim tmpDmg As Long
    If tSetup.noche = True And esDeNoche = True Then
        If LogAlpha > 155 Then LogAlpha = validbyte(LogAlpha - 3)
        If LogAlpha < 153 Then LogAlpha = validbyte(LogAlpha + 3)
        lvalue(0) = D3DColorXRGB(LogAlpha, LogAlpha, LogAlpha + IIf(LogAlpha > 180, 0, LogAlpha / 5))
    Else
        If LogAlpha < MaxAlpha Then LogAlpha = validbyte(LogAlpha + 3)
        If MaxAlpha < LogAlpha Then LogAlpha = validbyte(LogAlpha - 3)
        lvalue(0) = D3DColorXRGB(LogAlpha, LogAlpha, LogAlpha)
    End If
    lvalue(1) = lvalue(0)
    lvalue(2) = lvalue(0)
    lvalue(3) = lvalue(0)
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            
            'Layer 1 **********************************
            If Not mostrarcapa(1) Then If Not MapData(x, Y).Graphic(1).GrhIndex = 1 Then _
            Call DDrawGrhtoSurface(MapData(x, Y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                0, 1, lvalue)
            '******************************************
            
            'Layer 2 **********************************
            If Not mostrarcapa(2) Then
                If MapData(x, Y).Graphic(2).GrhIndex <> 0 Then
                    Call DDrawGrhtoSurface(MapData(x, Y).Graphic(2), _
                    (ScreenX - 1) * TilePixelWidth + PixelOffsetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffsetY, _
                    0, 1, lvalue)
                End If
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Transparent Layers
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            With MapData(x, Y)
            
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, lvalue)
                End If
                '***********************************************
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    If Not mostrarcapa(0) Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
                
                'Layer 3 *****************************************
                If EsArbol(.Graphic(3).GrhIndex) And tSetup.transArboles = True Then
                    If (Y > (UserPos.Y - 2) And Y < (UserPos.Y + 7)) And (x > (UserPos.x - 4) And x < (UserPos.x + 4)) Then
                        If .Graphic(3).GrhIndex <> 0 Then
                            'Draw
                            If Not mostrarcapa(3) Then Call DDrawTransGrhtoSurface(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, lvalue, , 155)
                        End If
                    Else
                        If .Graphic(3).GrhIndex <> 0 Then
                            'Draw
                           If Not mostrarcapa(3) Then Call DDrawTransGrhtoSurface(.Graphic(3), _
                                    PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, lvalue)
                        End If
                    End If
                Else
                    If .Graphic(3).GrhIndex <> 0 Then
                        'Draw
                        If Not mostrarcapa(3) Then Call DDrawTransGrhtoSurface(.Graphic(3), _
                                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, lvalue)
                    End If
                End If
                '************************************************
                
                'Damage layer *****************************************
                For tmpDmg = 0 To 8 '
                    With .Damage(tmpDmg)
                        If .Using = True Then
                            If .Wait <= 0 Then
                                If .Alpha >= 3 Then .Alpha = validbyte(.Alpha - timerElapsedTime / 5)
                            Else
                                .Wait = .Wait - 1
                            End If

                            .OffSetY = .OffSetY + validbyte(timerElapsedTime / 20)
                            
                            If .Alpha <= 50 Then .Using = False
                            
                            If tSetup.EfectosPelea Then Call DrawText(PixelOffsetXTemp + 16, PixelOffsetYTemp - Round(.OffSetY / 2), .Label, D3DColorXRGB(validbyte(.r - (255 - .Alpha)), validbyte(.g - (255 - .Alpha)), validbyte(.g - (255 - .Alpha))), 255, True, 0)
                        End If
                    End With
                Next tmpDmg
                If showbloqs Then
                    If MapData(x, Y).Blocked <> 0 Then
                        Call DrawText(PixelOffsetXTemp + 16, PixelOffsetYTemp, "B", ColorToDX8(vbRed), , True)
                    End If
                End If
                '*****************************************************
                
            End With
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
    
    ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
            For x = minX To maxX
                
                'Layer 4 **************************
                If MapData(x, Y).Graphic(4).GrhIndex Then
                    
                            If Not mostrarcapa(4) Then Call DDrawTransGrhtoSurface(MapData(x, Y).Graphic(4), _
                                ScreenX * TilePixelWidth + PixelOffsetX, _
                                ScreenY * TilePixelHeight + PixelOffsetY, _
                                1, 1, lvalue, , alphaTecho)
                    
                    
                
                End If
                
                '**********************************
                
                ScreenX = ScreenX + 1
            Next x
            ScreenY = ScreenY + 1
        Next Y
        
    If ModoCombate Then
        Call DrawText(5, 5, "Modo Combate", D3DColorXRGB(LogAlpha, 0, 0))
    End If
    
    If frmMain.macrotrabajo.Enabled Then
        Call DrawText(5, 5, "Trabajando...", D3DColorXRGB(LogAlpha, 0, 0))
    End If
    

End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    'Dim SurfaceDesc As DDSURFACEDESC2
    'Dim ddck As DDCOLORKEY
    
    IniPath = App.path & "\Init\"
    Movement_Speed = 1
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    fps = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
On Error GoTo 0
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    Call LoadGraphics
    
    InitTileEngine = True
End Function

Public Sub DirectXInit()
    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        If tSetup.VSync = False Then
            .SwapEffect = D3DSWAPEFFECT_COPY
            UsaVSync = False
        Else
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            UsaVSync = True
        End If
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = 800
        .BackBufferHeight = 600
        .EnableAutoDepthStencil = 1
        
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With
    
    Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.MainViewPic.hwnd, _
                        D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                        D3DWindow)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    With DirectDevice
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR
    End With
    
    MainViewRect.Left = 0
    MainViewRect.Top = 0
    MainViewRect.Right = frmMain.MainViewPic.ScaleWidth
    MainViewRect.Bottom = frmMain.MainViewPic.ScaleHeight

    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    

    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
End Sub

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next

    Set DirectD3D = Nothing
    
    Set DirectX = Nothing
End Sub

Sub ShowNextFrame(ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************


    If EngineRun Then
        DirectDevice.BeginScene
        DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        
        
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If
        If ClientSetup.bActive Then
            If isCapturePending Then
                Call ScreenCapture(True)
                isCapturePending = False
            End If
        End If
        Call Dialogos.Render
        Call DibujarCartel
        
        Call DialogosClanes.Draw
        
        If bTecho Then
            If tSetup.AlphaBlending Then
                If alphaTecho - (timerElapsedTime / 15) >= 70 Then
                    alphaTecho = validbyte(alphaTecho - (timerElapsedTime / 10))
                Else
                    alphaTecho = 70
                End If
            Else
                alphaTecho = 0
            End If
        Else
            If tSetup.AlphaBlending Then
                If alphaTecho + (timerElapsedTime / 15) <= 255 Then
                    alphaTecho = validbyte(alphaTecho + (timerElapsedTime / 10))
                Else
                    alphaTecho = 255
                End If
            Else
                alphaTecho = 255
            End If
        End If
        
        DirectDevice.EndScene
        DirectDevice.Present MainViewRect, ByVal 0, frmMain.MainViewPic.hwnd, ByVal 0
        
        
        
        If EnInventario = True Then
            ctimeInv = GetTickCount
            If ctimeInv - tInv >= 150 Then
                tInv = ctimeInv
                Inventario.DrawInv
            
            End If
        End If
        ''Static fpsLastChk As Long, dif As Long
       '' dif = GetTickCount - fpsLastCheck
       '' If dif > 0 Then
       ''     Do While FramesPerSecCounter / (GetTickCount - fpsLastCheck) > (IIf(tSetup.LimitFps = True, 18, 65) / 1000)  ''0.018
       ''         Sleep 5
       ''     Loop
       '' End If
        Static SpeedLimit As Byte
        If tSetup.LimitFps = True Then
            SpeedLimit = 56
        Else ''para 36 fps es 28
            SpeedLimit = 16
        End If
        If SpeedLimit <> 0 Then
           'While (GetTickCount - fpsLastCheck) / SpeedLimit < FramesPerSecCounter
           '    Sleep 5
            'Wend
        End If
        'Limitado
        
        'FPS update
        If fpsLastCheck + 1000 < GetTickCount Then
            fps = FramesPerSecCounter + IIf(tSetup.LimitFps = True, 0, 2)
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If

End Sub

Public Function validbyte(ByVal param As Long) As Byte
    If param < 0 Then param = 0
    If param > 255 Then param = 255
    validbyte = param
End Function

Public Function SetElapsedTime(ByVal Start As Boolean) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 23/05/2011 By MaTeO
'Gets the time that past since the last call
'[MaTeO] Agrego cambios a la funcion
'**************************************************************
    Dim Start_Time As Currency
    Static End_Time As Currency
    Static Timer_Freq As Currency
    'Get the timer frequency
    If Timer_Freq = 0 Then
        QueryPerformanceFrequency Timer_Freq
    End If
   
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
   
    If Not Start Then
        'Calculate elapsed time
        SetElapsedTime = (Start_Time - End_Time) / Timer_Freq * 1000
   
        'Get next end time
    Else
        Call QueryPerformanceCounter(End_Time)
    End If
End Function

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static End_Time As Currency
    Static Timer_Freq As Currency

    'Get the timer frequency
    If Timer_Freq = 0 Then
        QueryPerformanceFrequency Timer_Freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - End_Time) / Timer_Freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(End_Time)
End Function

Public Function Engine_Get_2_Points_Angle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************

    Engine_Get_2_Points_Angle = Engine_Get_X_Y_Angle((X2 - X1), (Y2 - Y1))
   
End Function
Public Function Engine_Get_X_Y_Angle(ByVal x As Double, ByVal Y As Double) As Double
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************

Dim dblres              As Double
 
    dblres = 0
   
    If (Y <> 0) Then
        dblres = Engine_Convert_Radians_To_Degrees(Atn(x / Y))
        If (x <= 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (x > 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (x < 0 And Y > 0) Then
            dblres = dblres + 360
        End If
    Else
        If (x > 0) Then
            dblres = 90
        ElseIf (x < 0) Then
            dblres = 270
        End If
    End If
   
    Engine_Get_X_Y_Angle = dblres
   
End Function

 
Public Function Engine_Convert_Radians_To_Degrees(ByVal s_radians As Double) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/25/2004
'Converts a radian to degrees
'**************************************************************

      Engine_Convert_Radians_To_Degrees = (s_radians * 180) / 3.14159265358979
 
End Function


Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long

    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                .LastMov = GetTickCount
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                .LastMov = GetTickCount
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        'If done moving stop animation
        'If done moving stop animation
        If .Heading = 0 Then .Heading = SOUTH
        If Not moved Then
            If GetTickCount - .LastMov > 50 Then
                'Stop animations
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
                
                If Not .Movimient Then
                    .Arma.WeaponWalk(.Heading).Started = 0
                    .Arma.WeaponWalk(.Heading).FrameCounter = 1
                    
                    .Escudo.ShieldWalk(.Heading).Started = 0
                    .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                End If
                
                .Moving = False
            End If
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        
        
        Dim lvalue(3) As Long
        lvalue(0) = D3DColorXRGB(LogAlpha, LogAlpha, LogAlpha)
        lvalue(1) = lvalue(0)
        lvalue(2) = lvalue(0)
        lvalue(3) = lvalue(0)
        
        Dim vAlpha As Byte
        vAlpha = 255
        If .invisible Then
            If InviConAlpha(CharIndex) Then
                vAlpha = 155
            Else
                'vAlpha = 0
            End If
        End If
        
        If .Head.Head(.Heading).GrhIndex Then
            If (Not .invisible) Or (vAlpha = 155) Then
                
                If prgRun = False Then
                    Dim lvalues(3) As Long
                    lvalues(0) = D3DColorXRGB(validbyte(ColoresPJ(.priv).r - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).g - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).b - (255 - LogAlpha)))
                    lvalues(1) = lvalues(0)
                    lvalues(2) = lvalues(0)
                    lvalues(3) = lvalues(0)
                    Static Rotac As Double
                    Rotac = Rotac + 0.01
                    If Rotac >= 360 Then Rotac = 0
''                    Call DDrawTransGrhIndextoSurface(15385, PixelOffsetX, PixelOffsetY + 0, 1, True, 255, lvalues, 0)
                    Call DDrawTransGrhIndextoSurface(15372, PixelOffsetX, PixelOffsetY + 40, 1, True, 255, lvalues, Rotac)
                End If
                
               '' .Rotac = .Rotac + 0.004
               '' If .Rotac >= 360 Then .Rotac = 0
                ''AURA
                ''Dim loopxx As Long
                ''For loopxx = 0 To 4
                ''    If .aura(loopxx).AuraGrh Then
                ''''''        Dim Coloir(3) As Long
                  ''      Coloir(0) = .aura(loopxx).Color
                  ''      Coloir(1) = .aura(loopxx).Color
                 ''       Coloir(2) = .aura(loopxx).Color
                ''        Coloir(3) = .aura(loopxx).Color
                     
                ''        ''Call DDrawTransGrhIndextoSurface(.aura(loopxx).AuraGrh, PixelOffsetX + .aura(loopxx).OffSetX, PixelOffsetY + .aura(loopxx).OffSetY, 1, True, 255, Coloir(), IIf((.aura(loopxx).Giratoria = True), .Rotac, 0))
               ''     End If
               '' Next loopxx
                Movement_Speed = 0.7
                'Draw Body
                If (vAlpha = 155 And tSetup.AlphaBlending = True And .invisible) Or (Not .invisible) Then
                    If .Body.Walk(.Heading).GrhIndex Then
                        If Not .muerto Then
                            If Not (.iBody = 8 Or .iBody = 145) Then
                                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = False, 255, vAlpha))
                            Else
                                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = True, 127, 255))
                            End If
                        Else
                            Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = True, 127, 255))
                        End If
                    End If
                End If
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    
                    
                    If (vAlpha = 155 And tSetup.AlphaBlending = True And .invisible) Or (Not .invisible) Then
                        If Not .muerto Then
                            If Not (.iHead = 500 Or .iHead = 514) Then
                                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = False, 255, vAlpha))
                            Else
                                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = False, 127, 255))
                            End If
                        Else
                            Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, lvalue, , IIf(tSetup.AlphaBlending = False, 127, 255))
                        End If
                        Dim toffx As Integer, toffy As Integer
                        If .iCasco = 7 Or .iCasco = 8 Then
                            Select Case .Heading
                                Case E_Heading.NORTH
                                    toffx = 0
                                Case E_Heading.SOUTH
                                    toffx = -2
                                Case E_Heading.EAST
                                    toffx = 0
                                Case E_Heading.WEST
                                    toffx = -2
                            End Select
                            toffy = -1
                        Else
                            toffx = 0
                        End If
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), toffx + PixelOffsetX + .Body.HeadOffset.x, toffy + PixelOffsetY + .Body.HeadOffset.Y, 1, 0, lvalue, , IIf(tSetup.AlphaBlending = False, 255, vAlpha))
                        
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = False, 255, vAlpha))
                        
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , IIf(tSetup.AlphaBlending = False, 255, vAlpha))
                    
                    End If
                    
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then 'And (esGM(UserCharIndex) Or Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2) Then
                            Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            
                            If .priv = 0 Then
                                If .Atacable Then
                                    Color = D3DColorXRGB(validbyte(ColoresPJ(48).r - (255 - LogAlpha)), validbyte(ColoresPJ(48).g - (255 - LogAlpha)), validbyte(ColoresPJ(48).b - (255 - LogAlpha)))
                                Else
                                    If .Criminal Then
                                        Color = D3DColorXRGB(validbyte(ColoresPJ(50).r - (255 - LogAlpha)), validbyte(ColoresPJ(50).g - (255 - LogAlpha)), validbyte(ColoresPJ(50).b - (255 - LogAlpha)))
                                    Else
                                        Color = D3DColorXRGB(validbyte(ColoresPJ(49).r - (255 - LogAlpha)), validbyte(ColoresPJ(49).g - (255 - LogAlpha)), validbyte(ColoresPJ(49).b - (255 - LogAlpha)))
                                    End If
                                End If
                            Else
                                Color = D3DColorXRGB(validbyte(ColoresPJ(.priv).r - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).g - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).b - (255 - LogAlpha)))
                            End If
                            
                            If vAlpha = 155 Then Color = vbWhite
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call DrawText(PixelOffsetX + 17, PixelOffsetY + 30, line, Color, IIf(tSetup.AlphaBlending = False, 255, vAlpha), True)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            If .priv > 0 Then
                                Select Case .priv
                                    Case PlayerType.Admin
                                        line = "<Administrador>"
                                        
                                    Case PlayerType.Dios
                                        line = "<Dios>"
                                        
                                    Case PlayerType.SemiDios
                                        line = "<Semidios>"
                                        
                                    Case PlayerType.Consejero
                                        line = "<Consejero>"
                                End Select
                            End If
                            Call DrawText(PixelOffsetX + 17, PixelOffsetY + 45, line, Color, IIf(tSetup.AlphaBlending = False, 255, vAlpha), True)

                        End If
                    End If
                End If
            End If
        Else
            'Draw Body
            If (Not .invisible) Or vAlpha = 155 Then
                If (vAlpha = 155 And tSetup.AlphaBlending = True And .invisible) Or (Not .invisible) Then
                    If .Body.Walk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, lvalue, , vAlpha)
                End If
                'If vAlpha = 155 Then
                    If LenB(.Nombre) > 0 Then
                        If Nombres Then 'And (esGM(UserCharIndex) Or Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2) Then
                            Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                            
                            If .priv = 0 Then
                                If .Atacable Then
                                    Color = D3DColorXRGB(validbyte(ColoresPJ(48).r - (255 - LogAlpha)), validbyte(ColoresPJ(48).g - (255 - LogAlpha)), validbyte(ColoresPJ(48).b - (255 - LogAlpha)))
                                Else
                                    If .Criminal Then
                                        Color = D3DColorXRGB(validbyte(ColoresPJ(50).r - (255 - LogAlpha)), validbyte(ColoresPJ(50).g - (255 - LogAlpha)), validbyte(ColoresPJ(50).b - (255 - LogAlpha)))
                                    Else
                                        Color = D3DColorXRGB(validbyte(ColoresPJ(49).r - (255 - LogAlpha)), validbyte(ColoresPJ(49).g - (255 - LogAlpha)), validbyte(ColoresPJ(49).b - (255 - LogAlpha)))
                                    End If
                                End If
                            Else
                                Color = D3DColorXRGB(validbyte(ColoresPJ(.priv).r - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).g - (255 - LogAlpha)), validbyte(ColoresPJ(.priv).b - (255 - LogAlpha)))
                            End If
                            
                            If vAlpha = 155 Then Color = vbWhite
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call DrawText(PixelOffsetX + 17, PixelOffsetY + 30, line, Color, IIf(tSetup.AlphaBlending = False, 255, vAlpha), True)
                            
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call DrawText(PixelOffsetX + 17, PixelOffsetY + 45, line, Color, IIf(tSetup.AlphaBlending = False, 255, vAlpha), True)

                        End If 'if nombres
                    End If ' if lenb
                
                End If ' if valpha
            'End If 'ifnot invisible
        End If

        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)   '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        Movement_Speed = 1
        Dim LoopC As Long
        For LoopC = 1 To 4
            If .Proyectil(LoopC).Usado = True Then
                Call ActualizarProyectil(CharIndex, PixelOffsetX, PixelOffsetY, LoopC)
                If tSetup.EfectosPelea Then
                    Call DDrawTransGrhIndextoSurface(.Proyectil(LoopC).GrhIndex, .Proyectil(LoopC).ActualX, .Proyectil(LoopC).ActualY, 1, False, 255, lvalue)
                End If
            End If
        Next LoopC
        'Draw FX
        If .FxIndex <> 0 Then
        
            Dim XDATAFX As Integer, YDATAFX As Integer
        
            '@Nota de Dunkan: Arreglar desde el INDICE.
            If .FxIndex = 1 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 18 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY - 15
            ElseIf .FxIndex = 17 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY - 15
            ElseIf .FxIndex = 19 Then
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 7 Then    'TORMENTA DE FUEGO
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 30
            ElseIf .FxIndex = 8 Then    'PARALIZAR
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 35
            ElseIf .FxIndex = 9 Then    'CURAR GRAVES
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 25
            ElseIf .FxIndex = 12 Then   'INMO
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY + 20
            Else
                XDATAFX = (PixelOffsetX + FxData(.FxIndex).OffSetX)
                YDATAFX = PixelOffsetY
            End If
            
            Dim tmpGralalpha As Byte
            tmpGralalpha = 255
            If tSetup.AlphaBlending = True Then
                tmpGralalpha = 127
            Else
                tmpGralalpha = 255
            End If
            Call DDrawTransGrhtoSurface(FxGrh(.FxIndex), XDATAFX, YDATAFX, 1, 1, lvalue, False, tmpGralalpha)
                    
            'Check if animation is over
            If FxGrh(.FxIndex).Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Private Function InviConAlpha(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex)
        Dim Pos As Integer, CLAN As String
        Pos = getTagPosition(.Nombre)
        CLAN = mid$(.Nombre, Pos)
        
        If CharIndex = UserCharIndex Then
            InviConAlpha = True
            Exit Function
        End If
        
        If LenB(CLAN) <= 0 Then Exit Function
        
        Pos = getTagPosition(charlist(UserCharIndex).Nombre)
        If CLAN = mid$(charlist(UserCharIndex).Nombre, Pos) Then
            InviConAlpha = True
            Exit Function
        End If
        
        
        InviConAlpha = False
    End With
End Function

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        If .FxIndex > 0 Then
            Call InitGrh(FxGrh(fX), FxData(fX).Animacion)
            FxGrh(fX).Loops = Loops
        End If
    End With
End Sub

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    Dim r As RECT
    'Call BackBufferSurface.BltColorFill(r, vbBlack)
End Sub
Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef RGB_List() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), 0, src.Left / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), 0, 1, 1)
    End If

End Sub

Public Function Geometry_Create_TLVertex(ByVal x As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Public Sub Device_Textured_Render(ByVal x As Integer, ByVal Y As Integer, _
                                  ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, _
                                  light_value() As Long, Optional Alpha As Boolean = False, _
                                  Optional AlphaByte As Byte = 255, _
                                  Optional ByVal Shadow As Byte = 0, _
                                  Optional ByVal src_height As Integer = 0, _
                                  Optional ByVal src_width As Integer = 0, Optional ByVal Angle As Single = 0, _
                                  Optional ByVal GrhIndex As Integer)
                                
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    
    Dim srdesc As D3DSURFACE_DESC
    
    'light_value(0) = -1 'rgb_list(0)
    'light_value(1) = -1 'rgb_list(1)
    'light_value(2) = -1 'rgb_list(2)
    'light_value(3) = -1 'rgb_list(3)
    'If (light_value(0) = 0) Then light_value(0) = d3dcolorx
    'If (light_value(1) = 0) Then light_value(1) = BaseColor
    'If (light_value(2) = 0) Then light_value(2) = BaseColor
    'If (light_value(3) = 0) Then light_value(3) = BaseColor
    
    
    With dest_rect
        .Bottom = Y + (src_rect.Bottom - src_rect.Top) ' src_height
        .Left = x
        .Right = x + (src_rect.Right - src_rect.Left)
        .Top = Y
    End With
    
    Dim texwidth As Long, texheight As Long
    If Texture Is Nothing Then Exit Sub
    Texture.GetLevelDesc 0, srdesc
    
    
    texwidth = srdesc.Width
    texheight = srdesc.Height
    If Shadow Then
        Dim Color_Shadow(3) As Long
        Engine_Long_To_RGB_List Color_Shadow(), D3DColorARGB(50, 0, 0, 0)
    
          Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_Shadow, texwidth, texheight, 0 '', False, False

        
        ''Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_Shadow(), texwidth, texheight, Angle
    Else
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, Angle
    End If
    
    ''Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, angle ', Shadow, scr_height, src_width ' angle
    
    DirectDevice.SetTexture 0, Texture
    

    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, 3
        DirectDevice.SetRenderState D3DRS_DESTBLEND, 2
    End If
    
    ''If Alpha Then
     ''   DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
     ''   DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   '' End If
    Call DirectDevice.SetRenderState(D3DRS_TEXTUREFACTOR, D3DColorARGB(AlphaByte, 0, 0, 0)) 'ENGINE OPTIMIZAR, SE PUEDE HACER UN ARRAY(255) CON LOS POSIBLES COLORES.
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub Draw_FillBox(ByVal x As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static RGB_List(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    RGB_List(0) = Color
    RGB_List(1) = Color
    RGB_List(2) = Color
    RGB_List(3) = Color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .Bottom = Y + Height
        .Left = x
        .Right = x + Width
        .Top = Y
    End With
    
    With Outline
        .Bottom = Y + Height + 1
        .Left = x - 1
        .Right = x + Width + 1
        .Top = Y - 1
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, RGB_List(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub


Public Sub DrawText(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, ByVal Color As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal Center As Boolean = False, Optional ByVal Shadow As Byte = 1, Optional ByVal fontt As Byte = 2)
'*********************************************************
'****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
'*********************************************************
    
    'Engine_Render_Text cfonts(1), Text, left - 1, top, 0, Center, Alpha
    'Engine_Render_Text cfonts(1), Text, left + 1, top, 0, Center, Alpha
    'Engine_Render_Text cfonts(1), Text, left, top - 1, 0, Center, Alpha
    'Engine_Render_Text cfonts(1), Text, left, top + 1, 0, Center, Alpha
    If Shadow = 1 Then Engine_Render_Text cfonts(fontt), Text, Left - 2, Top - 1, 0, Center, Alpha
    Engine_Render_Text cfonts(fontt), Text, Left, Top, Color, Center, Alpha
End Sub

 Public Function GetR(ByVal lColor As Long)

GetR = lColor And RGB(255, 0, 0)

End Function

Public Function GetG(ByVal lColor As Long)

GetG = (lColor And RGB(0, 255, 0)) / 256

End Function

Public Function GetB(ByVal lColor As Long)

GetB = (lColor And RGB(0, 0, 255)) / 65536

End Function
Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal x As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Center As Boolean = False, Optional ByVal Alpha As Byte = 255)
Dim TempVA(0 To 3) As TLVERTEX
Dim tempstr() As String
Dim Count As Integer
Dim ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim J As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim YOffset As Single
 
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
   
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
   
    'Set the temp color (or else the first character has no color)
    TempColor = Color
 
    'Set the texture
    DirectDevice.SetTexture 0, UseFont.Texture
   
    If Center Then
        x = x - Engine_GetTextWidth(UseFont, Text) * 0.5
    End If
   
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
       
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
       
            'Loop through the characters
            For J = 1 To Len(tempstr(i))
 
                'Check for a key phrase
                'If ascii(j - 1) = 124 Then 'If Ascii = "|"
                '    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                '    If KeyPhrase Then TempColor = ARGB(255, 0, 0, alpha) Else ResetColor = 1
                'Else
 
                    'Render with triangles
                    'If AlternateRender = 0 Then
 
                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(J - 1)).Vertex(0), 32 * 4
 
                        'Set up the verticies
                        TempVA(0).x = x + Count
                        TempVA(0).Y = Y + YOffset
                       
                        TempVA(1).x = TempVA(1).x + x + Count
                        TempVA(1).Y = TempVA(0).Y
 
                        TempVA(2).x = TempVA(0).x
                        TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
 
                        TempVA(3).x = TempVA(1).x
                        TempVA(3).Y = TempVA(2).Y
                       
                        'Set the colors
                        TempVA(0).Color = TempColor
                        TempVA(1).Color = TempColor
                        TempVA(2).Color = TempColor
                        TempVA(3).Color = TempColor
                       
                        'Draw the verticies
                        Call DirectDevice.SetRenderState(D3DRS_TEXTUREFACTOR, D3DColorARGB(Alpha, 0, 0, 0))
                        DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                       
                     
                    'Shift over the the position to render the next character
                    Count = Count + UseFont.HeaderInfo.CharWidth(ascii(J - 1))
               
                'End If
               
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
               
            Next J
           
        End If
    Next i
   
End Sub

Public Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth]http://www.vbgore.com/GameClient.TileEn ... tTextWidth[/url]
'***************************************************
Dim i As Integer
 
    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
   
    'Loop through the text
    For i = 1 To Len(Text)
       
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
       
    Next i
 
End Function

Sub Init_FontRender()
    Engine_Init_FontTextures
    Engine_Init_FontSettings
End Sub
 
Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures]http://www.vbgore.com/GameClient.TileEn ... ntTextures[/url]
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
 
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    '*** Default font ***
   
    'Set the texture
    Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.path & "\Init\Font.bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
   
    'Store the size of the texture
    cfonts(1).TextureSize.x = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
   Set cfonts(2).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.path & "\Init\Fontaa.bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
   
    'Store the size of the texture
    cfonts(2).TextureSize.x = TexInfo.Width
    cfonts(2).TextureSize.Y = TexInfo.Height
    Exit Sub
eDebug:
    If Err.number = "-2005529767" Then
        MsgBox "Error en la textura utilizada de DirectX 8", vbCritical
        End
    End If
    End
 
End Sub
 Sub Engine_Init_FontSettings2()
'*********************************************************
'****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
'*********************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single
 
    '*** Default font ***
 
    'Load the header information
    FileNum = FreeFile
    Open App.path & "\Init\FontDataaa.dat" For Binary As #FileNum
        Get #FileNum, , cfonts(2).HeaderInfo
    Close #FileNum
   
    'Calculate some common values
    cfonts(2).CharHeight = cfonts(2).HeaderInfo.CellHeight - 4
    cfonts(2).RowPitch = cfonts(2).HeaderInfo.BitmapWidth \ cfonts(2).HeaderInfo.CellWidth
    cfonts(2).ColFactor = cfonts(2).HeaderInfo.CellWidth / cfonts(2).HeaderInfo.BitmapWidth
    cfonts(2).RowFactor = cfonts(2).HeaderInfo.CellHeight / cfonts(2).HeaderInfo.BitmapHeight
   
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
       
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(2).HeaderInfo.BaseCharOffset) \ cfonts(2).RowPitch
        u = ((LoopChar - cfonts(2).HeaderInfo.BaseCharOffset) - (Row * cfonts(2).RowPitch)) * cfonts(2).ColFactor
        v = Row * cfonts(2).RowFactor
 
        'Set the verticies
        With cfonts(2).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).x = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
           
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(2).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = cfonts(2).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
           
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(2).RowFactor
            .Vertex(2).x = 0
            .Vertex(2).Y = cfonts(2).HeaderInfo.CellHeight
            .Vertex(2).z = 0
           
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(2).ColFactor
            .Vertex(3).tv = v + cfonts(2).RowFactor
            .Vertex(3).x = cfonts(2).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(2).HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
       
    Next LoopChar
End Sub
Sub Engine_Init_FontSettings()
'*********************************************************
'****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
'*********************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single
 
    '*** Default font ***
 Engine_Init_FontSettings2
    'Load the header information
    FileNum = FreeFile
    Open App.path & "\Init\FontData.dat" For Binary As #FileNum
        Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
   
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
   
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
       
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        v = Row * cfonts(1).RowFactor
 
        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).x = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
           
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(1).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
           
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(1).RowFactor
            .Vertex(2).x = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).z = 0
           
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(1).ColFactor
            .Vertex(3).tv = v + cfonts(1).RowFactor
            .Vertex(3).x = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
       
    Next LoopChar
End Sub


Function EsArbol(ByVal GhrNumber As Long) As Boolean
    EsArbol = (GhrNumber = 7000 Or _
    GhrNumber = 7001 Or _
    GhrNumber = 7002 Or _
    GhrNumber = 641 Or _
    GhrNumber = 643 Or _
    GhrNumber = 644 Or _
    GhrNumber = 647 Or _
    GhrNumber = 735 Or _
    GhrNumber = 6581 Or _
    GhrNumber = 6582 Or _
    GhrNumber = 6583 Or _
    GhrNumber = 7222 Or _
    GhrNumber = 7223 Or _
    GhrNumber = 7224 Or _
    GhrNumber = 7225 Or _
    GhrNumber = 7226)
End Function

Private Function esDeNoche() As Boolean
    If Hour(time) >= 19 Or Hour(time) <= 5 Then
        esDeNoche = True
    End If
End Function


Sub RenderConnect()

    Dim lighthandle(3) As Long

    AlphaB = 255 ''validbyte((Caida - (CaidaConst - 580)) / 2)
    lighthandle(0) = D3DColorXRGB(AlphaB, AlphaB, AlphaB)
    lighthandle(1) = lighthandle(0)
    lighthandle(2) = lighthandle(0)
    lighthandle(3) = lighthandle(0)
 
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
    Dim SrcRect As RECT
    
    With SrcRect
        .Left = 0
        .Top = 0
        .Right = 800
        .Bottom = 600
    End With
    Device_Textured_Render 0, 0, SurfaceDB.Surface(9997), SrcRect, lighthandle
    
    
     

    Device_Textured_Render 0, -(CaidaConst - Caida), SurfaceDB.Surface(9998, True), SrcRect, lighthandle
            Call DrawGui
    Device_Textured_Render 0, 0, SurfaceDB.Surface(9996), SrcRect, lighthandle
    'Call DDrawTransGrhIndextoSurface(426, 500, 500, 0, Light)  'Carga un GrhIndex para el Render.
        Call EffectConnect 'Llamamos Al EffectConnect asi tenemos los textos y FillBoxs
        
        
        While (GetTickCount - fpsLastCheck) / 28 < FramesPerSecCounter
           Sleep 5
        Wend
        
        'FPS update
        If fpsLastCheck + 1000 < GetTickCount Then
            fps = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        

 
    DirectDevice.EndScene
    DirectDevice.Present ByVal 0, ByVal 0, frmConnect.MainViewPic.hwnd, ByVal 0
End Sub
 
Function EffectConnect()
    Call EfectoCaida

    
    DrawText 20, 580, "V " & App.Major & "." & App.Minor & " Release " & App.Revision, D3DColorXRGB(229, 220, 33), AlphaB, , 0
    
    
End Function


Public Sub ActualizarProyectil(ByVal CharIndex As Integer, ByVal x As Long, ByVal Y As Long, ByVal ind As Byte)
    Dim modi As Double
    modi = 1
    timerElapsedTime = timerElapsedTime / modi
    
    Dim target_Angle As Single
    With charlist(CharIndex)
        If .Proyectil(ind).Usado = True Then
            target_Angle = GetAngle(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y)
            .Proyectil(ind).ActualX = (.Proyectil(ind).ActualX + Sin(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            .Proyectil(ind).ActualY = (.Proyectil(ind).ActualY - Cos(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            If Distance(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y) <= (timerElapsedTime / 4) Then
                .Proyectil(ind).Usado = False
                Exit Sub
            End If
            target_Angle = GetAngle(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y)
            .Proyectil(ind).ActualX = (.Proyectil(ind).ActualX + Sin(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            .Proyectil(ind).ActualY = (.Proyectil(ind).ActualY - Cos(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            If Distance(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y) <= (timerElapsedTime / 4) Then
                .Proyectil(ind).Usado = False
                Exit Sub
            End If
            target_Angle = GetAngle(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y)
            .Proyectil(ind).ActualX = (.Proyectil(ind).ActualX + Sin(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            .Proyectil(ind).ActualY = (.Proyectil(ind).ActualY - Cos(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            If Distance(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y) <= (timerElapsedTime / 4) Then
                .Proyectil(ind).Usado = False
                Exit Sub
            End If
            target_Angle = GetAngle(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y)
            .Proyectil(ind).ActualX = (.Proyectil(ind).ActualX + Sin(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            .Proyectil(ind).ActualY = (.Proyectil(ind).ActualY - Cos(target_Angle * DegreeToRadian) * (timerElapsedTime / 4))
            If Distance(.Proyectil(ind).ActualX, .Proyectil(ind).ActualY, x, Y) <= (timerElapsedTime / 4) Then
                .Proyectil(ind).Usado = False
                Exit Sub
            End If
        End If
    End With
timerElapsedTime = timerElapsedTime * modi
    
End Sub

Function Engine_PixelPosX(ByVal x As Integer) As Integer

'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

    Engine_PixelPosX = (x - 1) * TilePixelWidth

End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer

'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * TilePixelHeight

End Function
Public Function Engine_TPtoSPX(ByVal x As Byte) As Long

'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'************************************************************

    Engine_TPtoSPX = Engine_PixelPosX(x - minX) + OffsetCounterX - 288 + ((10 - TileBufferSize) * 32)

End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long

'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - minY) + OffsetCounterY - 288 + ((10 - TileBufferSize) * 32)

End Function

Public Sub Engine_D3DColor_To_RGB_List(RGB_List() As Long, Color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    RGB_List(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)
End Sub

Public Sub Engine_Long_To_RGB_List(RGB_List() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    RGB_List(0) = long_color
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)
End Sub

