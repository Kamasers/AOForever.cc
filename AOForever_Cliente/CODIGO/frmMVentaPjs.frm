VERSION 5.00
Begin VB.Form frmMVentaPjs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprar un personaje"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin AOFClient.lvButtons_H lvButtons_H1 
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
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
      Begin AOFClient.lvButtons_H lvButtons_H2 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
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
      Begin AOFClient.lvButtons_H lvButtons_H3 
         Height          =   615
         Left            =   1800
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
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
      Begin VB.Label lblRaza 
         Caption         =   "Raza: Humano"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblClase 
         Caption         =   "Clase: Clérigo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblVida 
         Caption         =   "Vida: 340"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblNivel 
         Caption         =   "Nivel: 40 (57%)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblPrecio 
         Caption         =   "Precio: 5000 monedas"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.ListBox lstPjs 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMVentaPjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub lstPjs_Click()
    With UserMercado(lstPjs.ListIndex + 1)
        If .Ocupado = False Then
            lvButtons_H1.Enabled = False
            lblPrecio.Visible = False
            lblNivel.Visible = False
            lblVida.Visible = False
            lblClase.Visible = False
            lblRaza.Visible = False
            Exit Sub
        Else
            lvButtons_H1.Enabled = True
            lblPrecio.Visible = True
            lblNivel.Visible = True
            lblVida.Visible = True
            lblClase.Visible = True
            lblRaza.Visible = True
        End If
        lblPrecio.Caption = "Precio: " & .precio
        lblNivel.Caption = "Nivel: " & .Nivel & " (" & .Porcentaje & "%)"
        lblVida.Caption = "Vida: " & .Vida
        lblClase.Caption = "Clase: " & ListaClases(.Clase)
        lblRaza.Caption = "Raza: " & ListaRazas(.Raza)
        lvButtons_H3.Enabled = (.precio = 0)
        lvButtons_H1.Enabled = Not (.precio = 0)
    End With
End Sub

Private Sub lvButtons_H1_Click()
    Dim pin As String
    Dim Privado As String
    Privado = InputBox("Introduce la contraseña de la venta")
    pin = InputBox("Introduce el nuevo pin del personaje que quieres comprar")
    If Len(pin) > 8 Or Len(pin) < 4 Then
        MsgBox "Introduce un pin de seguridad de entre 4 y 8 caracteres"
        Exit Sub
    End If
    Call writeComprarUsr(lstPjs.ListIndex + 1, pin, Privado)
    
End Sub

Private Sub lvButtons_H2_Click()
    frmMIntercambio.Show , frmMain
    Unload Me
End Sub

