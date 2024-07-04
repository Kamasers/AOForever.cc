VERSION 5.00
Begin VB.Form frmMQuitarPj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quitar personaje de la venta"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2520
      Width           =   3975
   End
   Begin AOFClient.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
   Begin AOFClient.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Pin de seguridad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nick:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"frmMQuitarPj.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMQuitarPj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
    If Len(Text3.Text) < 4 Then
        MsgBox "Pin invalido. Tiene un minimo de 4 caracteres y un máximo de 8."
        Exit Sub
    End If
    If Len(Text1.Text) < 0 Or Len(Text2.Text) < 0 Then
        MsgBox "Escribe el nick y la contraseña"
        Exit Sub
    End If
    Call writequitarpj(Text1.Text, Text2.Text, Text3.Text)
    
End Sub

Private Sub lvButtons_H2_Click()
frmMIntercambio.Show , frmMain
    Unload Me
End Sub
