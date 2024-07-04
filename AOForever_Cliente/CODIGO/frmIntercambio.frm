VERSION 5.00
Begin VB.Form frmMIntercambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intercambio de Pjs"
   ClientHeight    =   2670
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Lista de personajes."
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar personaje."
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Publicar personaje."
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMIntercambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmMPjVenta.Show , frmMain
    Unload Me
End Sub

Private Sub Command2_Click()
    frmMQuitarPj.Show , frmMain
    Unload Me
    
End Sub

Private Sub Command3_Click()
    writeRequestListaMercado
    ''Unload Me
    
End Sub
