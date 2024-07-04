VERSION 5.00
Begin VB.Form frmMPjVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mundos del Sur - Mercado"
   ClientHeight    =   4365
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4110
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Venta privada (Con contraseña)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox txtPj 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtPin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin AOFClient.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.TextBox txtPrecio 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin AOFClient.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Personaje donde recibiras el oro:"
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
      Width           =   3855
   End
   Begin VB.Label Label3 
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
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Recuerda que una vez que hayas vendido/cambiado el personaje, no podrás recuperarlo."
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Precio:"
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
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "frmMPjVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "Si activas el modo candado, no podrás usar tu pj mientras éste se encuentre a la venta. Para sacarlo de la venta deberás hacerlo desde otro personaje."
End Sub
Private Sub Check1_Click()
    Text1.Enabled = Check1.Value
End Sub

Private Sub Check2_Click()
        txtPrecio.Enabled = Check2.Value
        txtPj.Enabled = Check2.Value
        Text1.Enabled = Check2.Value
        Check1.Enabled = Check2.Value
End Sub

Private Sub lvButtons_H1_Click()
    If Val(txtPrecio.Text) < 100000 Or Val(txtPrecio.Text) > 200000000 Then
        MsgBox "Precio inválido. El minimo es de 100.000 y el maximo de 200.000.000", , "Mercado de Usuarios"
        Exit Sub
    End If
    If Len(txtPin.Text) < 4 Or Len(txtPin.Text) > 8 Then
        MsgBox "El pin tiene un minimo de 4 caracteres y un maximo de 8"
        Exit Sub
    End If
    If Len(txtPj.Text) = 0 Then
        MsgBox "Escribe el nombre del personaje donde deseas recibir el oro."
        Exit Sub
    End If
    WritePublicarUser Val(txtPrecio.Text), txtPin.Text, txtPj.Text, Text1.Text
    Unload Me
End Sub

Private Sub lvButtons_H2_Click()
frmMIntercambio.Show , frmMain
    Unload Me
End Sub

Private Sub txtPrecio_Change()
    txtPrecio.Text = Val(txtPrecio.Text)
End Sub
