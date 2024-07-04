Attribute VB_Name = "m_DragAndDrop"
Option Explicit
Public CANTDRAG As Integer

Public Sub General_Drop_X_Y(ByVal X As Byte, ByVal Y As Byte)
On Error GoTo Err
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem <= Inventario.MaxObjs) Then
        If MapData(X, Y).Blocked = 1 And MapData(X, Y).CharIndex <= 0 Then
            Call ShowConsoleMsg("Elige una posición válida para tirar tus objetos.")
            Exit Sub
        End If
        
        If HayAgua(X, Y) = True Then
            Call ShowConsoleMsg("No está permitido tirar objetos en el agua.")
            Exit Sub
        End If
        
        If MapData(X, Y).CharIndex <> 0 And frmMain.DragToUser = False Then
            Call ShowConsoleMsg("Debes desactivar el seguro de transferencia de items(Click en la manito debajo del inventario)")
            Exit Sub
        End If

        If GetKeyState(vbKeyShift) < 0 Then
            frmCantDD.Show vbModal, frmMain
            If CANTDRAG <= 0 Then Exit Sub
            Call WriteDragToPos(X, Y, Inventario.SelectedItem, CANTDRAG)
        Else
            Call WriteDragToPos(X, Y, Inventario.SelectedItem, 1)
        End If
        
    End If
Err:

End Sub



