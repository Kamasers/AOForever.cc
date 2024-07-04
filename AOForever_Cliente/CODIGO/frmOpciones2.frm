VERSION 5.00
Begin VB.Form frmOpciones2 
   BackColor       =   &H00004080&
   Caption         =   "Opciones"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00004080&
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox Check10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   23
      Top             =   2640
      Width           =   195
   End
   Begin VB.CheckBox Check9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   21
      Top             =   2400
      Width           =   195
   End
   Begin VB.CheckBox Check8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   19
      Top             =   2160
      Width           =   195
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00004080&
      Caption         =   "Option1"
      Height          =   195
      Left            =   6600
      TabIndex        =   17
      Top             =   1920
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00004080&
      Caption         =   "Option1"
      Height          =   195
      Left            =   6600
      TabIndex        =   15
      Top             =   1680
      Width           =   195
   End
   Begin VB.CheckBox Check7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   195
   End
   Begin VB.CheckBox Check6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   195
   End
   Begin VB.CheckBox Check5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   2040
      Width           =   195
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   195
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   195
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cursores MDS"
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
      Left            =   6840
      TabIndex        =   24
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar Tips"
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
      Left            =   6840
      TabIndex        =   22
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar clave"
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
      Height          =   195
      Left            =   6840
      TabIndex        =   20
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   ".BMP"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ".JPG"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Capturar"
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
      Left            =   6960
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizar memoria de video"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Efectos de pelea"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "No fullscreen"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Noche"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alphablending"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Limitar FPS"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Arboles c/t"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmOpciones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Combo1.AddItem "Español"
    Combo1.AddItem "English"
End Sub

