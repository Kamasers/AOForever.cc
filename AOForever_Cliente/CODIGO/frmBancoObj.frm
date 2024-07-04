VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBancoObj 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   3765
   ClientTop       =   2550
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":0000
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   5520
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   393216
   End
   Begin VB.Timer aInv 
      Interval        =   100
      Left            =   5400
      Top             =   4560
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   5520
      TabIndex        =   10
      Text            =   "1"
      Top             =   3000
      Width           =   525
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   1920
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   9
      Top             =   3240
      Width           =   2910
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   11640
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1080
      Width           =   555
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   1
      Left            =   12120
      TabIndex        =   1
      Top             =   1800
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   1800
      Width           =   2490
   End
   Begin VB.PictureBox picBoveda 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   225
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   8
      Top             =   360
      Width           =   3870
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   18
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   17
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Label Bovedalbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Index           =   0
      Left            =   4080
      TabIndex        =   15
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   14
      Top             =   5040
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   13
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Inventariolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   3600
      Width           =   1770
   End
   Begin VB.Image Command2 
      Height          =   480
      Left            =   4920
      Top             =   5070
      Width           =   1110
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   1
      Left            =   10080
      Top             =   960
      Width           =   570
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   0
      Left            =   10080
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
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
      Left            =   8310
      TabIndex        =   7
      Top             =   150
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   5520
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3300
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   765
      Index           =   0
      Left            =   5535
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   2205
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   8310
      TabIndex        =   6
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
      Left            =   8550
      TabIndex        =   5
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
      Left            =   8490
      TabIndex        =   4
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   7245
      TabIndex        =   3
      Top             =   450
      Width           =   45
   End
End
Attribute VB_Name = "frmBancoObj"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean
Private last_i As Long
Private drag_modo As Byte

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

Private Sub CmdMoverBov_Click(Index As Integer)
If List1(0).ListIndex = -1 Then Exit Sub

If NoPuedeMover Then Exit Sub

Select Case Index
    Case 1 'subir
        If List1(0).ListIndex <= 0 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("No puedes mover el objeto en esa dirección.", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        LastIndex1 = List1(0).ListIndex - 1
    Case 0 'bajar
        If List1(0).ListIndex >= List1(0).ListCount - 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("No puedes mover el objeto en esa dirección.", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        LastIndex1 = List1(0).ListIndex + 1
End Select

NoPuedeMover = True
LasActionBuy = True
LastIndex2 = List1(1).ListIndex
Call WriteMoveBank(Index, List1(0).ListIndex + 1)
End Sub

Private Sub Command2_Click()
    Call WriteBankEnd
    NoPuedeMover = False
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command2.Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bOkBanco.jpg")
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
'Me.Picture = LoadPicture(App.path & "\Graficos\comerciar.jpg")
'Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
'Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")

'CmdMoverBov(1).Picture = LoadPicture(App.path & "\Graficos\FlechaSubirObjeto.jpg")
'CmdMoverBov(0).Picture = LoadPicture(App.path & "\Graficos\FlechaBajarObjeto.jpg")

End Sub




Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Not IsNumeric(cantidad.Text) Then Exit Sub

Select Case Index
    Case 1
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = InvBanco(0).SelectedItem
        LasActionBuy = True
        Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
        
   Case 0
        LastIndex2 = InvBanco(1).SelectedItem
        LasActionBuy = False
        Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
End Select


End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Index
    Case 0
      '  If Image1(0).Tag = 1 Then
      '          Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprarApretado.jpg")
      '          Image1(0).Tag = 0
      '          Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvender.jpg")
      '          Image1(1).Tag = 1
      '  End If
      Image1(0).Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bUpArrow.jpg")
        
    Case 1
      '  If Image1(1).Tag = 1 Then
        '        Image1(1).Picture = LoadPicture(App.path & "\Graficos\Botónvenderapretado.jpg")
        '        Image1(1).Tag = 0
         '       Image1(0).Picture = LoadPicture(App.path & "\Graficos\BotónComprar.jpg")
          '      Image1(0).Tag = 1
       ' End If
       Image1(1).Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bDownArrow.jpg")
        
End Select
End Sub



Private Sub picBoveda_Click()
    
    With InvBanco(0)
        If .SelectedItem <= 0 Then Exit Sub
        Bovedalbl(0) = .ItemName(.SelectedItem)
        Bovedalbl(1) = .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem)
        Bovedalbl(2) = .MaxHit(.SelectedItem)
        Bovedalbl(3) = .MinHit(.SelectedItem)
    End With
End Sub




Private Sub picUser_Click()
    With InvBanco(1)
        Inventariolbl(0) = .ItemName(.SelectedItem)
        Inventariolbl(1) = .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem)
        Inventariolbl(2) = .MaxHit(.SelectedItem)
        Inventariolbl(3) = .MinHit(.SelectedItem)
    End With
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If drag_modo <> 0 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
    End If
    'End If
Command2.Picture = LoadPicture("")
Image1(0).Picture = LoadPicture("")
Image1(1).Picture = LoadPicture("")
End Sub

Private Sub picBoveda_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If drag_modo = 1 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        WriteBankDeposit InvBanco(1).SelectedItem, Val(cantidad.Text)
    End If
    If drag_modo = 2 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
    End If
End Sub


Private Sub picUser_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If drag_modo = 2 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        WriteBankExtractItem InvBanco(0).SelectedItem, Val(cantidad.Text)
    End If
    If drag_modo = 1 And Button <> vbRightButton Then
        Me.MousePointer = vbDefault
        drag_modo = 0
        WriteDragInventory last_i, InvBanco(1).SelectedItem
    End If
End Sub

Private Sub picBoveda_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim Poss As Integer
Dim file As String
Dim i As Integer
        If InvBanco(0).SelectedItem <= 0 Then Exit Sub
    If InvBanco(0).SelectedItem > InvBanco(0).MaxItems Then Exit Sub
With InvBanco(0)
    If .SelectedItem <= 0 Then Exit Sub
    Bovedalbl(0) = .ItemName(.SelectedItem)
    Bovedalbl(1) = .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem)
    Bovedalbl(2) = .MaxHit(.SelectedItem)
    Bovedalbl(3) = .MinHit(.SelectedItem)
End With
    
If drag_modo <> 0 Then Exit Sub
If Button = vbRightButton Then

  If InvBanco(0).GrhIndex(InvBanco(0).SelectedItem) > 0 Then

        'If last_i > 0 And last_i <= MAX_INVENTORY_SLOTS Then
            
            last_i = InvBanco(0).SelectedItem
            Poss = BuscarI(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem))
            drag_modo = 2 '1 = de inventario a boveda
            If Poss = 0 Then
                i = GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum
 
                 file = App.path & "\grafs\" & GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum & ".bmp"
                 
                 frmBancoObj.ImageList1.ListImages.Add , CStr("g" & InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)), Picture:=LoadPicture(file)
                 Poss = frmBancoObj.ImageList1.ListImages.Count
            End If
           
               
            Set frmBancoObj.MouseIcon = frmBancoObj.ImageList1.ListImages(Poss).ExtractIcon
            frmBancoObj.MousePointer = vbCustom
            
        'End If
  End If
 
End If
End Sub
Private Sub picUser_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim Poss As Integer
Dim file As String
Dim i As Integer
    If InvBanco(1).SelectedItem <= 0 Then Exit Sub
    If InvBanco(1).SelectedItem > InvBanco(1).MaxItems Then Exit Sub
With InvBanco(1)
        Inventariolbl(0) = .ItemName(.SelectedItem)
        Inventariolbl(1) = .MinDef(.SelectedItem) & "/" & .MaxDef(.SelectedItem)
        Inventariolbl(2) = .MaxHit(.SelectedItem)
        Inventariolbl(3) = .MinHit(.SelectedItem)
    End With
If drag_modo <> 0 Then Exit Sub
If Button = vbRightButton Then
    

  If InvBanco(1).GrhIndex(InvBanco(1).SelectedItem) > 0 Then

        'If last_i > 0 And last_i <= MAX_INVENTORY_SLOTS Then
            
            last_i = InvBanco(1).SelectedItem
            Poss = BuscarI(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem))
            drag_modo = 1 '1 = de inventario a boveda
            If Poss = 0 Then
                i = GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum
 
                 file = App.path & "\grafs\" & GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum & ".bmp"
                 
                 frmBancoObj.ImageList1.ListImages.Add , CStr("g" & InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)), Picture:=LoadPicture(file)
                 Poss = frmBancoObj.ImageList1.ListImages.Count
            End If
           
               
            Set frmBancoObj.MouseIcon = frmBancoObj.ImageList1.ListImages(Poss).ExtractIcon
            frmBancoObj.MousePointer = vbCustom
            
        'End If
  End If
 
End If
End Sub

Private Function BuscarI(gh As Integer) As Integer
Dim i As Integer

For i = 1 To frmBancoObj.ImageList1.ListImages.Count
    If frmBancoObj.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
 
End Function

Private Sub Timer1_Timer()
                InvBanco(0).DrawInv
                InvBanco(1).DrawInv
End Sub
