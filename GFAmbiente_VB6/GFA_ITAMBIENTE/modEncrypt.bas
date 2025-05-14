Attribute VB_Name = "modEncrypt"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Set to True to make the password case-sensitive
#Const CASE_SENSITIVE_PASSWORD = False
Const password = "gfambiente"


Public Function getTextDecrypt(strText As String) As String
    strText = DecryptText(strText, password)
    getTextDecrypt = DecryptText(strText, password)

End Function

Public Function getTextEncrypt(strText As String) As String
    strText = EncryptText(strText, password)
    getTextEncrypt = EncryptText(strText, password)
    
End Function



'Encrypt text
Private Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function


Private Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function


Public Function getPassword(numPass As Integer) As String
    Dim retString As String
    Dim l As Long
    
    retString = Space(255)
    l = GetPrivateProfileString("PASSWORD", "password" & CStr(numPass), "", retString, Len(retString), App.Path & "\config.ini")
    If (l > 0) Then
        retString = left$(retString, l)
    Else
        retString = "": Exit Function
    End If
    
    
    numPass = numPass + 1
    
    
    getPassword = getTextDecrypt(Trim$(retString))
End Function
