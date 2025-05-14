Attribute VB_Name = "modEasyReadWriteRegistry"
'|---------------------------------------|
'| Written by Yash Kumar from LockX.com  |
'| Last Updated: June 16th, 2002         |
'| ActiveLock Open Source Code Version   |
'| Code Version: 1.00                    |
'|---------------------------------------|
'
'|------------------------------------------------|
'| Copyright (C) 2002 Apyl Software & Systems Ltd |
'| www.LockX.com                                  |
'| www.Apyl.com                                   |
'|------------------------------------------------|
'
'|------------------------------------------|
'| This module is available under the same  |
'| Open Source license as ActiveLock itself |
'|------------------------------------------|

Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_DWORD = 4 ' 32-bit number


Public Sub saveKey(hKey As Long, strPath As String)
    Dim keyhand&
    Dim r As Long

    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Public Function getString(hKey As Long, strPath As String, strValue As String) As String
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                getString = Left$(strBuf, intZeroPos - 1)
            Else
                getString = strBuf
            End If
        End If
    End If
    If getString = "" Then
    getString = "No! Value or Key Doesn't exist"
    End If
    
End Function


Public Function SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'text1.text= savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long

    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

    If r = 0 Then
        SaveString = "Success"
    Else
        SaveString = "No! Key to Delete Or Key Not Found"
    End If
    
End Function


Function getDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long

    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            getDWord = lBuf
        End If
        '
    End If
    
    r = RegCloseKey(keyhand)
    If getDWord = "" Then
        getDWord = " No! Value or key doesn't exist"
    End If
    
End Function


Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Text1.text= SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long

    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call e
    '     rrlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
    
    If r = 0 Then
        SaveDword = "Success"
    Else
        SaveDword = "No! Failed to save Value"
    End If
    
End Function


Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long

    r = RegDeleteKey(hKey, strKey)
    
    If r = 0 Then
        DeleteKey = "Success"
    Else
        DeleteKey = "No! Key to Delete Or Key Not Found"
    End If

End Function


Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    Dim r As Long

    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
    
    If r = 0 Then
        DeleteValue = "Success"
    Else
        DeleteValue = "No! Value to Delete"
    End If
    
End Function

        

