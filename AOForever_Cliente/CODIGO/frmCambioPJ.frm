VERSION 5.00
Begin VB.Form frmCambioPJ 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar solicitud de cambio de pj"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmCambioPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnviar_Click()
    If LenB(txtNick.Text) > 0 Then _
    Call WriteSolicitudCambioPj(txtNick.Text) Else MsgBox ("Introduce un nombre")
End Sub

