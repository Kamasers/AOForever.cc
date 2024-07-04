VERSION 5.00
Begin VB.Form frmEstadisticas 
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   65430
   ClientWidth     =   6855
   ForeColor       =   &H00000000&
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "FrmEstadisticas.frx":000C
   ScaleHeight     =   6270
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image command1 
      Height          =   210
      Index           =   38
      Left            =   5880
      Top             =   5130
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   36
      Left            =   5880
      Top             =   4935
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   34
      Left            =   5880
      Top             =   4680
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   32
      Left            =   5880
      Top             =   4470
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   30
      Left            =   5880
      Top             =   4215
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   28
      Left            =   5880
      Top             =   3990
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   5880
      Top             =   3750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   24
      Left            =   5880
      Top             =   3525
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   22
      Left            =   5880
      Top             =   3270
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   20
      Left            =   5880
      Top             =   3045
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   18
      Left            =   5880
      Top             =   2820
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   16
      Left            =   5880
      Top             =   2565
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   14
      Left            =   5880
      Top             =   2340
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   12
      Left            =   5880
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   10
      Left            =   5880
      Top             =   1860
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   8
      Left            =   5880
      Top             =   1635
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   6
      Left            =   5880
      Top             =   1395
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   4
      Left            =   5880
      Top             =   1170
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   2
      Left            =   5880
      Top             =   945
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   0
      Left            =   5880
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   39
      Left            =   3240
      Top             =   5130
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   37
      Left            =   3240
      Top             =   4935
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   1
      Left            =   3240
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   35
      Left            =   3240
      Top             =   4680
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   33
      Left            =   3240
      Top             =   4470
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   31
      Left            =   3240
      Top             =   4215
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   29
      Left            =   3240
      Top             =   3990
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   27
      Left            =   3240
      Top             =   3750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   25
      Left            =   3240
      Top             =   3525
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   23
      Left            =   3240
      Top             =   3270
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   21
      Left            =   3240
      Top             =   3045
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   19
      Left            =   3240
      Top             =   2820
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   17
      Left            =   3240
      Top             =   2565
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   15
      Left            =   3240
      Top             =   2340
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   13
      Left            =   3240
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   11
      Left            =   3240
      Top             =   1860
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   9
      Left            =   3240
      Top             =   1635
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   7
      Left            =   3240
      Top             =   1395
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   5
      Left            =   3240
      Top             =   1170
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   3
      Left            =   3240
      Top             =   945
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   20
      Left            =   4560
      TabIndex        =   62
      Top             =   5175
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   19
      Left            =   5040
      TabIndex        =   61
      Top             =   4935
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   18
      Left            =   5280
      TabIndex        =   60
      Top             =   4695
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   17
      Left            =   4800
      TabIndex        =   59
      Top             =   4470
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   16
      Left            =   4440
      TabIndex        =   58
      Top             =   4230
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   15
      Left            =   4320
      TabIndex        =   57
      Top             =   3990
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   14
      Left            =   4440
      TabIndex        =   56
      Top             =   3765
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   13
      Left            =   4200
      TabIndex        =   55
      Top             =   3525
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   12
      Left            =   4080
      TabIndex        =   54
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   11
      Left            =   5160
      TabIndex        =   53
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   10
      Left            =   4320
      TabIndex        =   52
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   51
      Top             =   2580
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   8
      Left            =   4680
      TabIndex        =   50
      Top             =   2350
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   7
      Left            =   4330
      TabIndex        =   49
      Top             =   2115
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   48
      Top             =   1875
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   4200
      TabIndex        =   47
      Top             =   1650
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   5160
      TabIndex        =   46
      Top             =   1410
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   45
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   4110
      TabIndex        =   44
      Top             =   945
      Width           =   495
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   43
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblDATOS 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   5920
      Width           =   5535
   End
   Begin VB.Image lblCerrar 
      Height          =   315
      Left            =   5905
      Top             =   5889
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   41
      Top             =   5580
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   40
      Top             =   5340
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   39
      Top             =   5100
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   38
      Top             =   4860
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   37
      Top             =   4620
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   4380
      Width           =   2475
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   945
      TabIndex        =   35
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   7
      Left            =   285
      TabIndex        =   25
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   6
      Left            =   285
      TabIndex        =   24
      Top             =   3420
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   285
      TabIndex        =   23
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   285
      TabIndex        =   22
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   285
      TabIndex        =   21
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   285
      TabIndex        =   20
      Top             =   2475
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   19
      Top             =   2235
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   945
      TabIndex        =   18
      Top             =   1965
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4650
      TabIndex        =   6
      Top             =   255
      Width           =   465
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   5
      Top             =   1365
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   4
      Top             =   1155
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   3
      Top             =   945
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   735
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   510
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1005
      TabIndex        =   0
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Navegacion: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   20
      Left            =   3585
      TabIndex        =   34
      Top             =   5160
      Width           =   945
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate sin armas: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   19
      Left            =   3585
      TabIndex        =   33
      Top             =   4935
      Width           =   1470
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Armas con proyectiles: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   18
      Left            =   3585
      TabIndex        =   32
      Top             =   4695
      Width           =   1680
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domar animales: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   17
      Left            =   3585
      TabIndex        =   31
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liderazgo: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   16
      Left            =   3585
      TabIndex        =   30
      Top             =   4230
      Width           =   795
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Herreria: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   15
      Left            =   3585
      TabIndex        =   29
      Top             =   3990
      Width           =   690
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carpinteria: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   14
      Left            =   3585
      TabIndex        =   28
      Top             =   3750
      Width           =   900
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mineria: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   13
      Left            =   3585
      TabIndex        =   27
      Top             =   3525
      Width           =   615
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesca: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   12
      Left            =   3585
      TabIndex        =   26
      Top             =   3285
      Width           =   525
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa con escudos: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   11
      Left            =   3585
      TabIndex        =   17
      Top             =   3045
      Width           =   1635
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comercio: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   10
      Left            =   3585
      TabIndex        =   16
      Top             =   2820
      Width           =   765
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Talar árboles: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   9
      Left            =   3585
      TabIndex        =   15
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervivencia: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   8
      Left            =   3600
      TabIndex        =   14
      Top             =   2340
      Width           =   1110
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   7
      Left            =   3585
      TabIndex        =   13
      Top             =   2115
      Width           =   795
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apuñalar: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   6
      Left            =   3585
      TabIndex        =   12
      Top             =   1875
      Width           =   750
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meditar: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   5
      Left            =   3585
      TabIndex        =   11
      Top             =   1635
      Width           =   645
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate con armas: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   3585
      TabIndex        =   10
      Top             =   1410
      Width           =   1530
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tacticas de combate: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   3585
      TabIndex        =   9
      Top             =   1170
      Width           =   1575
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   8
      Top             =   930
      Width           =   540
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magia: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   3585
      TabIndex        =   7
      Top             =   705
      Width           =   525
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Private clsFormulario As clsFormMovementManager
Private Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
End Enum
Dim selectedIndex As Byte
Private botonMas As Picture
Private botonMenos As Picture
Private botonMas1 As Picture
Private botonMenos1 As Picture
Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer

For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = AtributosNames(i) & ": " & UserAtributos(i)
Next
'For i = 1 To NUMSKILLS
'    Skills(i).Caption = SkillsNames(i) & ": "
'Next
For i = 1 To NUMSKILLS
Text1(i).Caption = UserSkills(i)
Next




Label4(1).Caption = "Asesino: " & UserReputacion.AsesinoRep
Label4(2).Caption = "Bandido: " & UserReputacion.BandidoRep
Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
Label4(4).Caption = "Ladrón: " & UserReputacion.LadronesRep
Label4(5).Caption = "Noble: " & UserReputacion.NobleRep
Label4(6).Caption = "Plebe: " & UserReputacion.PlebeRep

If UserReputacion.Promedio < 0 Then
    Label4(7).ForeColor = &H8080FF
    Label4(7).Caption = "Status: CRIMINAL"
Else
    Label4(7).ForeColor = &HC0C000
    Label4(7).Caption = "Status: Ciudadano"
End If

With UserEstadisticas
    Label6(0).Caption = "Criminales matados: " & .CriminalesMatados
    Label6(1).Caption = "Ciudadanos matados: " & .CiudadanosMatados
    Label6(2).Caption = "Usuarios matados: " & .UsuariosMatados
    Label6(3).Caption = "NPCs matados: " & .NpcsMatados
    Label6(4).Caption = "Clase: " & .Clase
    Label6(5).Caption = "Tiempo restante en carcel: " & .PenaCarcel
End With
If SkillPoints <> 0 Then
For i = 0 To NUMSKILLS * 2 - 1
    If (i And &H1) = 0 Then
        'Command1(i).Picture = LoadPicture(App.path & "\Graficos\Button\NonSelected\BotónMás.jpg")
        'Command1(i).Visible = True
    Else
        'Command1(i).Picture = LoadPicture(App.path & "\Graficos\Button\NonSelected\BotónMenos.jpg")
        'Command1(i).Visible = True
    End If
Next
Else
'For i = 0 To NUMSKILLS
'Command1(i).Visible = False
'Next
End If

Alocados = SkillPoints
lblDATOS.Caption = "Nivel: " & UserLvl & " Experiencia: " & UserExp & "/" & UserPasarNivel & " Skills Libres: " & SkillPoints
End Sub

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If Alocados > 0 Then
        indice = Index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If Val(Text1(indice).Caption) < MAXSKILLPOINTS Then
            Text1(indice).Caption = Val(Text1(indice).Caption) + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = Index \ 2 + 1
        If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) - 1
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If


lblDATOS.Caption = "Nivel: " & UserLvl & " Experiencia: " & UserExp & "/" & UserPasarNivel & " Skills Libres: " & Alocados
End Sub



Private Sub lblCerrar_Click()
 Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i
    
    Call WriteModifySkills(skillChanges())
    
    'If Alocados = 0 Then frmMain.Label1.Visible = False
    SkillPoints = Alocados
    Unload Me
End Sub



Private Sub lblCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCerrar.Picture = LoadPicture(App.path & "\Graficos\Button\Selected\bExitEstadisticas.jpg")
End Sub


Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
If Not Command1(Index).Tag = "1" Then
Select Case Index
    Case 0
        Command1(Index).Picture = botonMas1
    Case 1
        Command1(Index).Picture = botonMenos1
    Case 2
        Command1(Index).Picture = botonMas1
    Case 3
        Command1(Index).Picture = botonMenos1
    Case 4
        Command1(Index).Picture = botonMas1
    Case 5
        Command1(Index).Picture = botonMenos1
    Case 6
        Command1(Index).Picture = botonMas1
    Case 7
        Command1(Index).Picture = botonMenos1
    Case 8
        Command1(Index).Picture = botonMas1
    Case 9
        Command1(Index).Picture = botonMenos1
    Case 10
        Command1(Index).Picture = botonMas1
    Case 11
        Command1(Index).Picture = botonMenos1
    Case 12
        Command1(Index).Picture = botonMas1
    Case 13
        Command1(Index).Picture = botonMenos1
    Case 14
        Command1(Index).Picture = botonMas1
    Case 15
        Command1(Index).Picture = botonMenos1
    Case 16
        Command1(Index).Picture = botonMas1
    Case 17
        Command1(Index).Picture = botonMenos1
    Case 18
        Command1(Index).Picture = botonMas1
    Case 19
        Command1(Index).Picture = botonMenos1
    Case 20
        Command1(Index).Picture = botonMas1
    Case 21
        Command1(Index).Picture = botonMenos1
    Case 22
        Command1(Index).Picture = botonMas1
    Case 23
        Command1(Index).Picture = botonMenos1
    Case 24
        Command1(Index).Picture = botonMas1
    Case 25
        Command1(Index).Picture = botonMenos1
    Case 26
        Command1(Index).Picture = botonMas1
    Case 27
        Command1(Index).Picture = botonMenos1
    Case 28
        Command1(Index).Picture = botonMas1
    Case 29
        Command1(Index).Picture = botonMenos1
    Case 30
        Command1(Index).Picture = botonMas1
    Case 31
        Command1(Index).Picture = botonMenos1
    Case 32
        Command1(Index).Picture = botonMas1
    Case 33
        Command1(Index).Picture = botonMenos1
    Case 34
        Command1(Index).Picture = botonMas1
    Case 35
        Command1(Index).Picture = botonMenos1
    Case 36
        Command1(Index).Picture = botonMas1
    Case 37
        Command1(Index).Picture = botonMenos1
    Case 38
        Command1(Index).Picture = botonMas1
    Case 39
        Command1(Index).Picture = botonMenos1
    Case 40
        Command1(Index).Picture = botonMas1
    Case 41
        Command1(Index).Picture = botonMenos1
    End Select
    Command1(Index).Tag = "1"
End If
    
selectedIndex = Index \ 2 + 1
If Not Text1(selectedIndex).Tag = "1" Then
    Skills(selectedIndex).FontBold = True
    Text1(selectedIndex).FontBold = True
    Text1(selectedIndex).Left = Skills(selectedIndex).Left + Skills(selectedIndex).Width
    Text1(selectedIndex).Tag = "1"
End If
End Sub
Private Sub Form_Load()
    Set botonMas = LoadPicture(App.path & "\Graficos\Button\NonSelected\BotónMás.jpg")
    Set botonMenos = LoadPicture(App.path & "\Graficos\Button\NonSelected\BotónMenos.jpg")
    Set botonMas1 = LoadPicture(App.path & "\Graficos\Button\Selected\BotónMás.jpg")
    Set botonMenos1 = LoadPicture(App.path & "\Graficos\Button\Selected\BotónMenos.jpg")
    MirandoAsignarSkills = True
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Flags para saber que skills se modificaron
    ReDim flags(1 To NUMSKILLS)
     Dim i As Long
    For i = 0 To NUMSKILLS * 2 - 1
        If (i And &H1) = 0 Then
                Command1(i).Picture = botonMas
                Command1(i).Visible = True
        Else
                Command1(i).Picture = botonMenos
                Command1(i).Visible = True
        End If
    Next
        
    For i = 1 To NUMSKILLS
        Skills(i).FontBold = False
        Text1(i).FontBold = False
        If Text1(i).FontBold = False Then
            Text1(i).Left = Skills(i).Left + Skills(i).Width
        End If
    Next i
    lblCerrar.Picture = LoadPicture("")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MirandoAsignarSkills = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim i As Long
For i = 0 To NUMSKILLS * 2 - 1
    If (i And &H1) = 0 Then
        If Command1(i).Tag = "1" Then
            Command1(i).Picture = botonMas
            Command1(i).Visible = True
            Command1(i).Tag = "0"
        End If
    Else
        If Command1(i).Tag = "1" Then
            Command1(i).Picture = botonMenos
            Command1(i).Visible = True
            Command1(i).Tag = "0"
        End If
    End If
Next
    
For i = 1 To 20
    If Text1(i).Tag = "1" Then
        Text1(selectedIndex).Tag = "0"
        Skills(i).FontBold = False
        Text1(i).FontBold = False
        Text1(i).Left = Skills(i).Left + Skills(i).Width
    End If
Next i
lblCerrar.Picture = LoadPicture("")
End Sub
