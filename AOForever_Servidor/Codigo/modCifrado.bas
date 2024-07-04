Attribute VB_Name = "modCifrado"
Option Explicit
  
'//For Action parameter in EncryptString
Const ENCRYPT = 1
Const DECRYPT = 2
  
'---------------------------------------------------------------------
' EncryptString
' Modificado por Harvey T.
'---------------------------------------------------------------------
  
Public Function EncryptString( _
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
      
    '//Get UserKey characters
    n = Len(UserKey)
    ReDim UserKeyASCIIS(1 To n)
    For i = 1 To n
        UserKeyASCIIS(i) = Asc(mid$(UserKey, i, 1))
    Next
          
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(mid$(Text, i, 1))
    Next
      
    '//Encryption/Decryption
    If Action = ENCRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TextASCIIS(i) + UserKeyASCIIS(j)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           Temp = TextASCIIS(i) - UserKeyASCIIS(j)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
      
    '//Return
    EncryptString = rtn
End Function

