Attribute VB_Name = "mod_Crypt"
Option Explicit

Const asci As Byte = 47

Public Function Encrypt(ByVal Cadena As String) As String
    Dim x As Long, tLen As Integer, newString As String
    tLen = Len(Cadena)
    For x = 1 To tLen
        newString = newString & Chr$(asc(mid(Cadena, x, 1)) + asci)
    Next x
    Encrypt = newString
End Function

Public Function Decrypt(ByVal Cadena As String) As String
    Dim x As Long, tLen As Integer, newString As String
    tLen = Len(Cadena)
    For x = 1 To tLen
        newString = newString & Chr$(asc(mid(Cadena, x, 1)) - asci)
    Next x
    Decrypt = newString
End Function
