VERSION 5.00
Begin VB.Form frmCambiarPJ 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar personaje"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana: 1700"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Vida: 390"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel: 47"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblRaza 
      BackStyle       =   0  'Transparent
      Caption         =   "Raza: Humano"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblClase 
      BackStyle       =   0  'Transparent
      Caption         =   "Clase: Clerigo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nhelk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Width           =   4335
   End
End
Attribute VB_Name = "frmCambiarPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAceptar_Click()
    Dim pin As String
    pin = InputBox("Introduce el código de pin de tu personaje", "Cambiar personaje")
    If Len(pin) < 4 Or Len(pin) > 8 Then
        MsgBox "El pin tiene un minimo de 4 caracteres y un maximo de 8"
        Exit Sub
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

