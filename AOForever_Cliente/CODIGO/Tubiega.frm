VERSION 5.00
Begin VB.Form Tubiega 
   Caption         =   "Eventos"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Items"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Min_Win"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Max_Win"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Evento"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Tubiega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    Call WriteDoAim(Val(Text1.Text), Val(Text2.Text), (Check1.Value = 1))
End Sub
