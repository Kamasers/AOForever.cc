Attribute VB_Name = "mod_Barras"
Option Explicit

Public ModoCombate As Boolean
Public Cursor As New clsAniCursor
Public Sub InitBarras()
    
    Dim X As Long, w As Integer, h As Integer
    For X = 1 To 8
        Load frmMain.lblEnergia(X)
        Load frmMain.lblVida(X)
        Load frmMain.lblMana(X)
        Load frmMain.lblSed(X)
        Load frmMain.lblHambre(X)
        Load frmMain.lblLvl(X)
        Load frmMain.lblName(X)
        Load frmMain.lblPorcLvl(X)
        
        
        frmMain.lblEnergia(X).Font.Size = frmMain.lblEnergia(0).Font.Size
        frmMain.lblVida(X).Font.Size = frmMain.lblVida(0).Font.Size
        frmMain.lblMana(X).Font.Size = frmMain.lblMana(0).Font.Size
        frmMain.lblSed(X).Font.Size = frmMain.lblSed(0).Font.Size
        frmMain.lblHambre(X).Font.Size = frmMain.lblHambre(0).Font.Size
        frmMain.lblLvl(X).Font.Size = frmMain.lblLvl(0).Font.Size
        frmMain.lblName(X).Font.Size = frmMain.lblName(0).Font.Size
        frmMain.lblPorcLvl(X).Font.Size = frmMain.lblPorcLvl(0).Font.Size
        
        frmMain.lblEnergia(X).Font.Name = frmMain.lblEnergia(0).Font.Name
        frmMain.lblVida(X).Font.Name = frmMain.lblVida(0).Font.Name
        frmMain.lblMana(X).Font.Name = frmMain.lblMana(0).Font.Name
        frmMain.lblSed(X).Font.Name = frmMain.lblSed(0).Font.Name
        frmMain.lblHambre(X).Font.Name = frmMain.lblHambre(0).Font.Name
        frmMain.lblLvl(X).Font.Name = frmMain.lblLvl(0).Font.Name
        frmMain.lblName(X).Font.Name = frmMain.lblName(0).Font.Name
        frmMain.lblPorcLvl(X).Font.Name = frmMain.lblPorcLvl(0).Font.Name
        
        frmMain.lblEnergia(X).Font.bold = frmMain.lblEnergia(0).Font.bold
        frmMain.lblVida(X).Font.bold = frmMain.lblVida(0).Font.bold
        frmMain.lblMana(X).Font.bold = frmMain.lblMana(0).Font.bold
        frmMain.lblSed(X).Font.bold = frmMain.lblSed(0).Font.bold
        frmMain.lblHambre(X).Font.bold = frmMain.lblHambre(0).Font.bold
        frmMain.lblLvl(X).Font.bold = frmMain.lblLvl(0).Font.bold
        frmMain.lblName(X).Font.bold = frmMain.lblName(0).Font.bold
        frmMain.lblPorcLvl(X).Font.bold = frmMain.lblPorcLvl(0).Font.bold
        
        frmMain.lblEnergia(X).Visible = True '1
        frmMain.lblVida(X).Visible = True '2
        frmMain.lblMana(X).Visible = True '3
        frmMain.lblSed(X).Visible = True '4
        frmMain.lblHambre(X).Visible = True '5
        frmMain.lblLvl(X).Visible = True '6
        frmMain.lblName(X).Visible = True '7
        frmMain.lblPorcLvl(X).Visible = False '8
        
        frmMain.lblEnergia(X) = frmMain.lblEnergia(0)
        frmMain.lblVida(X) = frmMain.lblVida(0)
        frmMain.lblMana(X) = frmMain.lblMana(0)
        frmMain.lblSed(X) = frmMain.lblSed(0)
        frmMain.lblHambre(X) = frmMain.lblHambre(0)
        frmMain.lblLvl(X) = frmMain.lblLvl(0)
        frmMain.lblName(X) = frmMain.lblName(0)
        frmMain.lblPorcLvl(X) = frmMain.lblPorcLvl(0)
        
        frmMain.lblEnergia(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblVida(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblMana(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblSed(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblHambre(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblLvl(X).ForeColor = RGB(0, 0, 0)
        frmMain.lblName(X).ForeColor = RGB(50, 50, 50)
        frmMain.lblPorcLvl(X).ForeColor = RGB(0, 0, 0)
        
        
        SetBarPosition X, frmMain.lblEnergia(X), frmMain.lblEnergia(0).Left, frmMain.lblEnergia(0).Top
        SetBarPosition X, frmMain.lblVida(X), frmMain.lblVida(0).Left, frmMain.lblVida(0).Top
        SetBarPosition X, frmMain.lblMana(X), frmMain.lblMana(0).Left, frmMain.lblMana(0).Top
        SetBarPosition X, frmMain.lblSed(X), frmMain.lblSed(0).Left, frmMain.lblSed(0).Top
        SetBarPosition X, frmMain.lblHambre(X), frmMain.lblHambre(0).Left, frmMain.lblHambre(0).Top
        SetBarPosition X, frmMain.lblLvl(X), frmMain.lblLvl(0).Left, frmMain.lblLvl(0).Top
        SetBarPosition X, frmMain.lblName(X), frmMain.lblName(0).Left, frmMain.lblName(0).Top
        SetBarPosition X, frmMain.lblPorcLvl(X), frmMain.lblPorcLvl(0).Left, frmMain.lblPorcLvl(0).Top
    Next X
    frmMain.STAShp.ZOrder 1
    frmMain.MANShp.ZOrder 1
    frmMain.Hpshp.ZOrder 1
    frmMain.AGUAsp.ZOrder 1
    frmMain.COMIDAsp.ZOrder 1
    frmMain.COMIDAsp.Height = frmMain.COMIDAsp.Height - 2
End Sub

Private Sub SetBarPosition(ByVal i As Long, ByRef obj As Object, ByVal l As Integer, ByVal T As Integer)
    Select Case i
        Case 1
            obj.Left = l - 1
            obj.Top = T
        Case 2
            obj.Left = l + 1
            obj.Top = T
        Case 3
            obj.Left = l
            obj.Top = T - 1
        Case 4
            obj.Left = l
            obj.Top = T + 1
        Case 5
            obj.Left = l - 1
            obj.Top = T - 1
        Case 6
            obj.Left = l - 1
            obj.Top = T + 1
        Case 7
            obj.Left = l + 1
            obj.Top = T + 1
        Case 8
            obj.Left = l + 1
            obj.Top = T - 1
    End Select
End Sub

















