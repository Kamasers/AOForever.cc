VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewPassword.frx":0000
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Image BotonVolver 
      Height          =   480
      Left            =   130
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Image BotonCambiar 
      Height          =   480
      Left            =   1590
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password nuevo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Re ingrese su pasword nuevo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password viejo"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BotonCambiar_Click()
If Text2.Text <> Text3.Text Then
        Call MsgBox("Las contraseñas no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub
    End If
    
    Call WriteChangePassword(Text1.Text, Text2.Text)
    Unload Me
End Sub

Private Sub BotonCambiar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BotonCambiar.Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bChangePasswd.jpg")
End Sub

Private Sub BotonVolver_Click()
Unload Me
End Sub

Private Sub BotonVolver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BotonVolver.Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bCancelPasswd.jpg")
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BotonVolver.Picture = LoadPicture("")
BotonCambiar.Picture = LoadPicture("")
End Sub
