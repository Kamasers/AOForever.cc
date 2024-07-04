VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComerciar 
   BackColor       =   &H001D4A78&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   3360
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   393216
   End
   Begin VB.PictureBox picComerciar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   315
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   9
      Top             =   360
      Width           =   2910
   End
   Begin VB.PictureBox picUsuario 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   3435
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   8
      Top             =   375
      Width           =   2910
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "1"
      Top             =   3120
      Width           =   720
   End
   Begin VB.Timer tmInventario 
      Interval        =   100
      Left            =   3720
      Top             =   3120
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   2
      Left            =   5670
      Top             =   3030
      Width           =   735
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haz click en un item para mas información."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   75
      Width           =   3675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   2925
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   4245
      MouseIcon       =   "frmComerciar.frx":22C12
      Tag             =   "1"
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   0
      Left            =   1125
      MouseIcon       =   "frmComerciar.frx":22D64
      Tag             =   "1"
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   4
      Top             =   975
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2730
      TabIndex        =   2
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   750
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   0
      Top             =   450
      Width           =   45
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private objdrag As Byte
Private drag_modo As Byte
Private last_i As Long

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = 1
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub



Private Sub Form_Load()
'Cargamos la interfase
'Me.Picture = LoadPicture(App.path & "\Graficos\comerciar.jpg")
'Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
'Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")

End Sub




''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub

Select Case Index
    Case 0
        If Not InvComNpc.SelectedItem <> 0 Then Exit Sub

        If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).valor, Val(cantidad.Text)) Then
            Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text, 0)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   
   Case 1
        If Not InvComUsu.SelectedItem <> 0 Then Exit Sub
        Call WriteCommerceSell(InvComUsu.SelectedItem, cantidad.Text)
        
    Case 2
        Call WriteCommerceEnd
        
End Select

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
            Image1(0).Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bBuyComerciar.jpg")
            Image1(0).Tag = 0
            Image1(1).Picture = LoadPicture("")
            Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
            Image1(1).Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bSellComerciar.jpg")
            Image1(1).Tag = 0
            Image1(0).Picture = LoadPicture("")
            Image1(0).Tag = 1
        End If
        
    Case 2
        Image1(2).Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bOkComercio.jpg")
        Image1(2).Tag = 0
        
End Select
End Sub

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private Sub picComerciar_Click()
    With InvComNpc
        If .SelectedItem <> 0 Then _
        lblData.Caption = .ItemName(.SelectedItem) & " Def: " & .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem) & " Hit: " & .MinHit(.SelectedItem) & "/" & .MaxHit(.SelectedItem) & " Valor: " & CalculateSellPrice(.valor(.SelectedItem), 1)
    End With
End Sub





Private Sub picUsuario_Click()

    With InvComUsu
        If .SelectedItem <> 0 Then _
        lblData.Caption = .ItemName(.SelectedItem) & " Def: " & .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem) & "/" & .MaxHit(.SelectedItem) & " Valor: " & CalculateSellPrice(.valor(.SelectedItem), 1)
    End With
End Sub











Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If drag_modo <> 0 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
    End If
    'End If
    
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture("")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture("")
    Image1(1).Tag = 1
End If
    Image1(2).Picture = LoadPicture("")
    Image1(2).Tag = 1
End Sub

Private Sub picComerciar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If drag_modo = 2 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        Call WriteCommerceSell(InvComUsu.SelectedItem, cantidad.Text)
    End If
    
    If drag_modo = 1 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
    End If
    
    If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture("")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture("")
    Image1(1).Tag = 1
End If
    Image1(2).Picture = LoadPicture("")
    Image1(2).Tag = 1
End Sub


Private Sub picUsuario_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture("")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture("")
    Image1(1).Tag = 1
End If
    Image1(2).Picture = LoadPicture("")
    Image1(2).Tag = 1
    If drag_modo = 1 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text, 0)
        
    End If
    If drag_modo = 2 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        WriteDragInventory objdrag, InvComUsu.ClickItem(x, Y)
    End If
End Sub


Private Sub picUsuario_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim Poss As Integer
    Dim file As String
    Dim i As Integer
    If InvComUsu.SelectedItem <= 0 Then Exit Sub
    If InvComUsu.SelectedItem > InvComUsu.MaxItems Then Exit Sub
    With InvComUsu
        If .SelectedItem <> 0 Then _
        lblData.Caption = .ItemName(.SelectedItem) & " Def: " & .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem) & "/" & .MaxHit(.SelectedItem) & " Valor: " & CalculateSellPrice(.valor(.SelectedItem), 1)
    End With
    If drag_modo <> 0 Then Exit Sub
    If Button = vbRightButton Then
    
        If InvComUsu.GrhIndex(InvComUsu.SelectedItem) > 0 Then
            objdrag = InvComUsu.SelectedItem
            last_i = InvComUsu.SelectedItem
            Poss = BuscarI(InvComUsu.GrhIndex(InvComUsu.SelectedItem))
            drag_modo = 2 '1 = de npc a inventario
            If Poss = 0 Then
                i = GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum
                
                file = App.path & "\grafs\" & GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum & ".bmp"
                
                Me.ImageList1.ListImages.Add , CStr("g" & InvComUsu.GrhIndex(InvComUsu.SelectedItem)), Picture:=LoadPicture(file)
                Poss = Me.ImageList1.ListImages.Count
            End If
            
            
            Set Me.MouseIcon = Me.ImageList1.ListImages(Poss).ExtractIcon
            Me.MousePointer = vbCustom

        End If
    End If
End Sub

Private Sub picComerciar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim Poss As Integer
    Dim file As String
    Dim i As Integer
    If InvComNpc.SelectedItem <= 0 Then Exit Sub
    If InvComNpc.SelectedItem > InvComNpc.MaxItems Then Exit Sub
        
    With InvComNpc
        If .SelectedItem <> 0 Then _
        lblData.Caption = .ItemName(.SelectedItem) & " Def: " & .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem) & " Hit: " & .MinHit(.SelectedItem) & "/" & .MaxHit(.SelectedItem) & " Valor: " & CalculateSellPrice(.valor(.SelectedItem), 1)
    End With
    
    If drag_modo <> 0 Then Exit Sub
    If Button = vbRightButton Then
        
        If (InvComNpc.GrhIndex(InvComNpc.SelectedItem) > 0) Then
            last_i = InvComNpc.SelectedItem
            Poss = BuscarI(InvComNpc.GrhIndex(InvComNpc.SelectedItem))
            drag_modo = 1 '1 = de npc a inventario
            If Poss = 0 Then
                i = GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum
                
                file = App.path & "\grafs\" & GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum & ".bmp"
                
                Me.ImageList1.ListImages.Add , CStr("g" & InvComNpc.GrhIndex(InvComNpc.SelectedItem)), Picture:=LoadPicture(file)
                Poss = Me.ImageList1.ListImages.Count
            End If
            
            
            Set Me.MouseIcon = Me.ImageList1.ListImages(Poss).ExtractIcon
            Me.MousePointer = vbCustom

        End If
    End If
End Sub


Private Function BuscarI(gh As Integer) As Integer
Dim i As Integer

For i = 1 To Me.ImageList1.ListImages.Count
    If Me.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
 
End Function

Private Sub Timer1_Timer()
    If frmComerciar.Visible = False Then Exit Sub
                InvComNpc.DrawInv
                InvComUsu.DrawInv
End Sub
