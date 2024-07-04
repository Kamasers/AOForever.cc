VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private ImgContainer As Boolean

'Private InvSurface As DirectD3DSurface7            'DD Surface used to render everything

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed

    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 14/01/2010
'Determines wether inventory can be scrolled up / down
'***************************************************
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
End Function

Public Property Get MaxItems()
    MaxItems = UBound(Inventory)
    
End Property


Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then

    End If
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it

End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).amount = amount
    
    'Update Amount

End Sub



Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    If slot <= 0 Then Exit Sub
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
If slot <= 0 Then Exit Sub
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    If slot <= 0 Then Exit Sub
    MaxDef = Inventory(slot).MaxDef
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    If slot <= 0 Then Exit Sub
    MinDef = Inventory(slot).MinDef
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    valor = Inventory(slot).valor
End Property

Public Property Get amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        amount = Inventory(slot).amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = Inventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    If slot <= 0 Then Exit Sub
    ItemName = Inventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = Inventory(slot).Equipped
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
    For slot = 1 To Me.MaxObjs

    Next slot
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .valor = eValor
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)

End Sub

Public Function ClickItem(ByVal x As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
 
    temp_x = x \ TilePixelWidth
    temp_y = Y \ TilePixelHeight
 
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
   
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then
        'Make sure slot isn't empty
        'If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        'Else
            'ClickItem = 0
        'End If
    End If
 
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
End Function

Private Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
On Local Error Resume Next
    Dim LoopC As Long
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    
    With TempRect
        .Bottom = InventoryWindow.ScaleHeight
        .Right = InventoryWindow.Width
    End With
    
    destRect = TempRect
    Dim lvalue(3) As Long
    lvalue(0) = D3DColorXRGB(255, 255, 255)
    lvalue(1) = lvalue(0)
    lvalue(2) = lvalue(0)
    lvalue(3) = lvalue(0)
    'Clear the inventory window
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
        For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
            startdX = ((LoopC - 1) Mod (InventoryWindow.Width / 32)) * 32
            startdY = ((LoopC - 1) \ (InventoryWindow.Width / 32)) * 32
            
            If Inventory(LoopC).GrhIndex Then
            
                'Get source rect|
                With SrcRect
                    .Left = GrhData(Inventory(LoopC).GrhIndex).sX
                    .Top = GrhData(Inventory(LoopC).GrhIndex).sY
                    .Right = .Left + ownTilePixelWidth
                    .Bottom = .Top + ownTilePixelHeight
                End With
                
                
                Device_Textured_Render startdX, startdY, SurfaceDB.Surface(CLng(GrhData(Inventory(LoopC).GrhIndex).FileNum)), SrcRect, lvalue
                'Render a box around the selected item
                If InvSelectedItem = LoopC Then
                    If ImgContainer Then
                        Device_Textured_Render startdX, startdY, SurfaceDB.Surface(9995), SrcRect, lvalue
                    End If
                End If
                
                
                'If equipped we render "+"
                If Inventory(LoopC).Equipped Then
                    DrawText startdX + 20, startdY + 20, "+", -256, , , 0
                End If
                
                'Render the item grh and the amount
                DrawText startdX + IIf(Inventory(LoopC).amount = 10000, 0, 2), startdY, Inventory(LoopC).amount, lvalue(0), , , 0
                
            End If
        Next LoopC
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, InventoryWindow.hwnd, ByVal 0
    
End Sub
Public Sub DrawInv()
DrawInventory
End Sub
#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectD3D As DirectD3D7, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
    ByRef PopUpMenu As Menu)

#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal StartX As Integer = 0, _
                            Optional ByVal StartY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True)
#End If
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    'initialize DX stuff
    'Dim SurfaceDesc As DDSURFACEDESC2
    
    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = StartX
    startdY = StartY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText

    'TODO : Fonts should be in a separate class / collection
    Dim font As New StdFont
    Dim Ifnt As IFont
    
    font.Name = "Verdana"
    font.bold = True
    font.italic = False
    font.Size = FontSize
    font.Underline = False
    font.Strikethrough = False
    
    Set Ifnt = font
    
'\TODO
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If x < 0 Or Y < 0 Or x > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = x \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(Inventory) Then
        InventoryWindow.ToolTipText = Inventory(TempItem).Name
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If x < 0 Or Y < 0 Or x > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(Y))
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Or InvSelectedItem Then Call DrawInventory
    End If

End Sub
