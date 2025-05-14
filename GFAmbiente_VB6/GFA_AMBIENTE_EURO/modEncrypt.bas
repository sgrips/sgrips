Attribute VB_Name = "modEncrypt"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'This is a test of a rudimentary encryption/decryption
'algorithm.

'Note: It is possible for the encrypted string to include Chr$(0).
'Visual Basic handles this just fine but Windows (and
'therefore the TextBox control) uses Chr$(0) to signify the end
'of the string. Therefore, programs such as this one that store
'the encrypted data in a text box may end up truncating the data.


'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623


'Set to True to make the password case-sensitive
#Const CASE_SENSITIVE_PASSWORD = False
Const password = "gfambiente"


Public Function getTextEncrypt(strText As String) As String
    strText = DecryptText(strText, password)
    getTextEncrypt = DecryptText(strText, password)
End Function




'Encrypt text
Private Function EncryptText(strText As String, ByVal strPwd As String)
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
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Private Function DecryptText(strText As String, ByVal strPwd As String)
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
    DecryptText = strBuff
End Function


Public Function getPassword(numPass As Integer) As String
    Dim retString As String
    Dim l As Long
    
    retString = Space(255)
    l = GetPrivateProfileString("PASSWORD", "password" & CStr(numPass), "", retString, Len(retString), App.Path & "\config.ini")
    If (l > 0) Then
        retString = Left$(retString, l)
    Else
        retString = "": Exit Function
    End If
    numPass = numPass + 1
    
    
    getPassword = getTextEncrypt(Trim$(retString))
End Function
