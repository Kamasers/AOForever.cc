Attribute VB_Name = "mod_Connect"
Option Explicit
 
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
"GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long '//Disco.

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
  
  
  
Dim OReg As Registro

Private Function GetSerialNumber(strDrive As String) As Long '//Disco.
    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation(strDrive, Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function

Public Function GetHD() As String
    GetHD = GetSerialNumber(ReadField(1, App.path, Asc("\")) & "\")
End Function

Public Sub BloqConnect()
    Call WriteVar(settingFile, "Init", "A" & "u" & "r" & "a" & "s", "1")
    Set OReg = New Registro
    Call OReg.CrearNuevaClave(HKEY_CURRENT_USER, "m" & "d" & "s" & "c" & "o" & "n" & "f" & "i" & "g")
    Call OReg.EstablecerValor(HKEY_CURRENT_USER, _
                          "m" & "d" & "s" & "c" & "o" & "n" & "f" & "i" & "g", _
                          LCase$("A" & "u" & "r" & "a" & "s"), _
                          "1", REG_SZ)
                          
    Set OReg = Nothing
    End
End Sub

'Consultar valor
Public Function tieneReg() As Boolean
    Set OReg = New Registro
    If OReg.ConsultarValor(HKEY_CURRENT_USER, "m" & "d" & "s" & "c" & "o" & "n" & "f" & "i" & "g", LCase$("A" & "u" & "r" & "a" & "s")) = "1" Then
        tieneReg = True
        Exit Function
    End If
    Set OReg = New Registro
    If GetVar(settingFile, "Init", "A" & "u" & "r" & "a" & "s") = "1" Then
        tieneReg = True
        Exit Function
    End If
    tieneReg = False
End Function


