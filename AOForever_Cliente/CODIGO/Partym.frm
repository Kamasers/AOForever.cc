VERSION 5.00
Begin VB.Form frmParty 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Partym.frx":0000
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4575
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Exp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Personaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   20
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   19
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   18
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   1560
         TabIndex        =   16
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   15
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   14
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Personaje1"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   795
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   0
      Left            =   3720
      Top             =   7500
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image bRechazar 
      Height          =   420
      Left            =   3705
      Picture         =   "Partym.frx":1D2E9
      Top             =   3210
      Width           =   1155
   End
   Begin VB.Image bAceptar 
      Height          =   420
      Left            =   2520
      Picture         =   "Partym.frx":200B8
      Top             =   3210
      Width           =   1125
   End
   Begin VB.Image bExpulsar 
      Height          =   480
      Left            =   120
      Picture         =   "Partym.frx":22E36
      Top             =   3180
      Width           =   2250
   End
   Begin VB.Image bCambiarPorcentajes 
      Height          =   480
      Left            =   2520
      Picture         =   "Partym.frx":263C4
      Top             =   3600
      Width           =   2325
   End
   Begin VB.Image bSalirParty 
      Height          =   480
      Left            =   120
      Picture         =   "Partym.frx":29C88
      Top             =   3645
      Width           =   2265
   End
   Begin VB.Image boton 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "Partym.frx":2D2FD
      Top             =   3645
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Integrantes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SalirParty As clsGraphicalButton
Private Aceptar As clsGraphicalButton
Private Rechazar As clsGraphicalButton
Private CambiarPorcentajes As clsGraphicalButton
Private Expulsar As clsGraphicalButton
Public LastPressed As clsGraphicalButton
Private Sub LoadButtons()
    Set SalirParty = New clsGraphicalButton
    Set Aceptar = New clsGraphicalButton
    Set Rechazar = New clsGraphicalButton
    Set CambiarPorcentajes = New clsGraphicalButton
    Set Expulsar = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    
    Dim BPath As String
    BPath = App.path & "\Graficos\Button\Party\"
    
    Call SalirParty.Initialize(bSalirParty, BPath & "bQuitParty.jpg", _
                               BPath & "bQuitPartyS.jpg", _
                               BPath & "bQuitPartyS.jpg", Me)
                               
    Call Rechazar.Initialize(bRechazar, BPath & "bRejectParty.jpg", _
                               BPath & "bRejectPartyS.jpg", _
                               BPath & "bRejectPartyS.jpg", Me, _
                               BPath & "bRejectPartyN.jpg")
    
    Call Aceptar.Initialize(bAceptar, BPath & "bAcceptParty.jpg", _
                               BPath & "bAcceptPartyS.jpg", _
                               BPath & "bAcceptPartyS.jpg", Me, _
                               BPath & "bAcceptPartyN.jpg")
                        
    Call CambiarPorcentajes.Initialize(bCambiarPorcentajes, BPath & "bChangePorc.jpg", _
                               BPath & "bChangePorcS.jpg", _
                               BPath & "bChangePorcS.jpg", Me, _
                               BPath & "bChangePorcN.jpg")
                               
    Call Expulsar.Initialize(bExpulsar, BPath & "bRemoveParty.jpg", _
                               BPath & "bRemovePartyS.jpg", _
                               BPath & "bRemovePartyS.jpg", Me, _
                               BPath & "bRemovePartyN.jpg")

End Sub

Private Sub bAceptar_Click()
If Not SoyLider Then Exit Sub
If LenB(List1.Text) > 0 Then Call writeAceptarParty(List1.Text)
End Sub

Private Sub bCambiarPorcentajes_Click()
If Not SoyLider Then Exit Sub
frmPartyPorc.Show , frmParty
End Sub

Private Sub bExpulsar_Click()
If Not SoyLider Then Exit Sub
If LenB(List2.Text) > 0 Then Call WriteEcharParty(List2.Text)
End Sub

Private Sub bRechazar_Click()
If Not SoyLider Then Exit Sub
End Sub

Private Sub bSalirParty_Click()
            Call WriteEcharParty(vbNullString)
            Unload Me
End Sub

Private Sub Form_Load()
LoadButtons

DoEvents

End Sub
Public Sub ToggleButtons()
    If SoyLider Then
        Call Rechazar.EnableButton(True)
        Call Aceptar.EnableButton(True)
        Call CambiarPorcentajes.EnableButton(True)
        Call Expulsar.EnableButton(True)
    Else
        Call Rechazar.EnableButton(False)
        Call Aceptar.EnableButton(False)
        Call CambiarPorcentajes.EnableButton(False)
        Call Expulsar.EnableButton(False)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastPressed.ToggleToNormal
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LastPressed.ToggleToNormal
End Sub

Private Sub Label1_Click()

    Call Unload(Me)
    
    Call frmMain.SetFocus

End Sub

Private Sub Label6_Click()

    If (Label6.Caption = ">>") Then
        Label6.Caption = "<<"
        List1.Visible = True
        List2.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Frame1.Visible = False
    Else
        Label6.Caption = ">>"
        List1.Visible = False
        List2.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Frame1.Visible = True
    End If
    
End Sub
















