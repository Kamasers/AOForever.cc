VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MainViewPic 
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      Picture         =   "frmConnect.frx":15F950
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12000
      Begin VB.Image imgBorrarPj 
         Height          =   255
         Left            =   6720
         Top             =   6360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image imgOlvidePass 
         Height          =   255
         Left            =   3720
         Top             =   6360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgConectarse 
         Height          =   615
         Left            =   3240
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Image imgSalir 
         Height          =   615
         Left            =   4560
         Top             =   7500
         Width           =   2655
      End
      Begin VB.Image imgCrearPj 
         Height          =   615
         Left            =   6120
         Top             =   6720
         Width           =   2535
      End
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matías Fernando Pequeño
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Código Postal 1405

Option Explicit
Private tmpPasswd As String


Public Mx As Integer
Public mY As Integer
Public mb As Integer

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()
'On Error Resume Next



''txtPasswd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call IniciarCaida(1)
    End If
    mod_Gui.Connect_KeyPress KeyAscii
    If KeyAscii = vbKeyTab Then
        Gui(1).HasFocus = Not Gui(1).HasFocus
        Gui(2).HasFocus = Not Gui(1).HasFocus
        KeyAscii = 0
    End If
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible

End Sub

Private Sub Form_Load()
        '[CODE 002]:MatuX
    EngineRun = False
    '[END]

     '[CODE]:MatuX
    '
    '  El código para mostrar la versión se genera acá para
    ' evitar que por X razones luego desaparezca, como suele
    ' pasar a veces :)
       ''version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    '[END]'
    Set MainViewPic.Picture = Nothing
    ''Me.Picture = LoadPicture(App.path & "\graficos\VentanaConectar.jpg")

End Sub





Private Sub imgBorrarPj_Click()
frmBorrarPj.Show , Me
End Sub

Private Sub imgConectarse_Click()
Conectarse
End Sub

Private Sub imgCrearPj_Click()
    

    EstadoLogin = E_MODO.Dados
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.state <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

End Sub


Private Sub imgOlvidePass_Click()
    frmOlvidePass.Show , frmConnect
End Sub

Private Sub imgSalir_Click()
    IniciarCaida 1
End Sub

Private Sub MainViewPic_Click()
    Call mod_Gui.Gui_Click
End Sub

Private Sub MainViewPic_DblClick()
Call mod_Gui.Gui_Click(True)
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Mx = x
    mY = Y
    mb = Button
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        imgConectarse_Click
        KeyAscii = 0
    End If
    
End Sub

