VERSION 5.00
Begin VB.Form frmOptions2 
   Caption         =   "Más opciones"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin tdsClonClient.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración del juego"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "No FullScreen"
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
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Árboles con transparencia"
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
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Efectos de combate"
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
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Limitar fps"
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
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Efecto noche"
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
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Caption         =   "AlphaBlending(Hechizos, meditaciones,etc)"
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
         Top             =   2040
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmOptions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim settingFile As String
Private Sub Check1_Click()
    Call WriteVar(settingFile, "Init", "NoFullScreen", Check1.value)
End Sub

Private Sub Check2_Click()
    Call WriteVar(settingFile, "Init", "TreeTransparence", Check2.value)
    
End Sub

Private Sub Check3_Click()
    Call WriteVar(settingFile, "Init", "FightingEfects", Check3.value)
End Sub

Private Sub Check4_Click()
    Call WriteVar(settingFile, "Init", "FpsLimit", Check4.value)
End Sub

Private Sub Check5_Click()
    Call WriteVar(settingFile, "Init", "Night", Check5.value)
End Sub

Private Sub Check6_Click()
    Call WriteVar(settingFile, "Init", "AlphaBlending", Check6.value)
End Sub

Private Sub Form_Load()
    settingFile = App.path & "/init/Settings.mds"
    Check1.value = Val(GetVar(settingFile, "Init", "NoFullScreen"))
    Check2.value = Val(GetVar(settingFile, "Init", "TreeTransparence"))
    Check3.value = Val(GetVar(settingFile, "Init", "FightingEfects"))
    Check4.value = Val(GetVar(settingFile, "Init", "FpsLimit"))
    Check5.value = Val(GetVar(settingFile, "Init", "Night"))
    Check6.value = Val(GetVar(settingFile, "Init", "AlphaBlending"))
End Sub

Private Sub lvButtons_H1_Click()
    Unload Me
End Sub
