VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   720
      Picture         =   "frmCargando.frx":15F942
      Top             =   3840
      Width           =   9960
   End
End
Attribute VB_Name = "frmCargando"
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
'La Plata - Pcia, Buenos Aires - Repub lica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Private NextWidth As Integer
Private ActualWidth As Integer
Private NextPercentage As Integer
Public Sub NewPercentage(ByVal Porc As Byte)
    ''664 width total
    ActualWidth = Round(Porc / 100 * 664)
    SetBarWidth

    If Porc = 100 Then
        Sleep 200
        Unload Me
    End If
    'Do While ActualWidth < NextWidth
    '    SubirWidth
    '    DoEvents
    '    Sleep 2
    'Loop
    
End Sub
Private Sub SetBarWidth()
    Image1.Width = ActualWidth
End Sub

Private Sub SubirWidth()
    If NextWidth > ActualWidth Then
        ActualWidth = ActualWidth + 10
        SetBarWidth
    End If

    If ActualWidth = 664 Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Image1.Width = 30
    ''    Dim poss As Byte, file As String
    ''file = App.path & "\Extras\hand.bmp"
    
   '' frmCargando.ImageList1.ListImages.Add , CStr("g1"), Picture:=LoadPicture(file)
    ''poss = frmCargando.ImageList1.ListImages.Count
            
    
   '' Set frmCargando.MouseIcon = frmCargando.ImageList1.ListImages(poss).ExtractIcon
    ''frmCargando.MousePointer = vbCustom
    ''Set frmConnect.MouseIcon = frmCargando.ImageList1.ListImages(poss).ExtractIcon
    ''frmConnect.MousePointer = vbCustom
End Sub

