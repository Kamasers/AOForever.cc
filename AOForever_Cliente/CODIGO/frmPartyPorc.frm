VERSION 5.00
Begin VB.Form frmPartyPorc 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Acomodar Porcentajes"
   ClientHeight    =   2985
   ClientLeft      =   4305
   ClientTop       =   3105
   ClientWidth     =   3270
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPartyPorc.frx":0000
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   9
      Text            =   "0"
      Top             =   2010
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   8
      Text            =   "0"
      Top             =   1650
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   7
      Text            =   "0"
      Top             =   1290
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Text            =   "0"
      Top             =   930
      Width           =   375
   End
   Begin VB.TextBox Porc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Text            =   "0"
      Top             =   570
      Width           =   375
   End
   Begin VB.Image bAceptar 
      Height          =   375
      Left            =   150
      Top             =   2490
      Width           =   975
   End
   Begin VB.Image bCancelar 
      Height          =   375
      Left            =   2160
      Top             =   2505
      Width           =   975
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Lin 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      Visible         =   0   'False
      X1              =   120
      X2              =   3120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Pj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pj1"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   225
   End
End
Attribute VB_Name = "frmPartyPorc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Aceptar As clsGraphicalButton
Private Cancelar As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private Sub LoadButtons()
    Set Aceptar = New clsGraphicalButton
    Set Cancelar = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    
    Dim BPath As String
    BPath = App.path & "\Graficos\Button\Party\"
                               
    Call Aceptar.Initialize(bAceptar, "", _
                               BPath & "bAcceptPartyPorcS.jpg", _
                               BPath & "bAcceptPartyPorcS.jpg", Me)
                               
    Call Cancelar.Initialize(bCancelar, "", _
                               BPath & "bCancelPartyPorcS.jpg", _
                               BPath & "bCancelPartyPorcS.jpg", Me)
End Sub

Private Function PorcentajesValidos() As Boolean
    Dim x As Long, c As Byte
    For x = 1 To MaxUsuariosParty
        If Porc(x).Enabled = True Then
            c = c + Val(Porc(x).Text)
        End If
    Next x
    If c > 100 Or c < 100 Then
        PorcentajesValidos = False
    Else
        PorcentajesValidos = True
    End If
        
End Function

Private Sub bAceptar_Click()
    If Not PorcentajesValidos Then MsgBox "Porcentajes inválidos": Unload Me: Exit Sub
    Dim x As Long
    For x = 1 To MaxUsuariosParty
        With Party(x)
            .Porc = Porc(x).Text
        End With
    Next x
    Call WriteSavePartyPorc
    Unload Me
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim x As Long
    For x = 1 To MaxUsuariosParty
        With Party(x)
            Pj(x).Caption = .Nick
            Porc(x).Text = .Porc
            If LenB(.Nick) <= 0 Then Porc(x).Text = 0: Porc(x).Enabled = False
        End With
    Next x
    If MaxPorc < 50 Then MaxPorc = 50
    
    LoadButtons
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
LastPressed.ToggleToNormal
End Sub

Private Function cantUsers() As Byte
    Dim x As Long
    For x = 1 To MaxUsuariosParty
        If LenB(Party(x).Nick) > 0 Then
            cantUsers = cantUsers + 1
        End If
    Next x
End Function

Private Sub Porc_Change(Index As Integer)
    Porc(Index).Text = Val(Porc(Index).Text)
    If Val(Porc(Index).Text) > MaxPorc And cantUsers > 1 Then
        Porc(Index).Text = MaxPorc
    End If
End Sub











