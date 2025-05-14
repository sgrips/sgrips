Attribute VB_Name = "modActiveLockSettings"
Option Explicit

' Author: Nick Hall
' Date  : Jul/2002

Public Sub alSaveSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vData As Variant)
 If gb_UseDataLock = True Then
 Else
  alSaveSettingReg strRegHive, strRegPath, strAppname, strSection, strKey, vData
 End If
End Sub

Public Function alGetSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vDefault As Variant) As Variant
 If gb_UseDataLock = True Then
 Else
  alGetSetting = alGetSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vDefault)
 End If
End Function

Public Sub alSaveSettingReg(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vData As Variant)
 Dim bFlag As Boolean
 Dim strRegKey As String
 Dim strRegBase As String
 Dim strValueName As String
 Dim strReturn As String
 Dim strdata As String
 Dim iTemp As Integer
 Dim iStart As Integer
 Dim lRegistryBase As Long
 
 Select Case Left(UCase(strRegHive), 4)
  Case "HKLM"
   lRegistryBase = HKEY_LOCAL_MACHINE
  Case "HKCR"
   lRegistryBase = HKEY_CLASSES_ROOT
  Case Else
   lRegistryBase = HKEY_CURRENT_USER
 End Select
 strRegKey = strRegPath
 If Right(strRegKey, 1) <> "\" Then
  strRegKey = strRegKey & "\"
 End If
 For iTemp = 1 To Len(strAppname)
  If Mid(strAppname, iTemp, 1) <> "\" Then
   strRegKey = strRegKey & Mid(strAppname, iTemp, 1)
  End If
 Next iTemp
 If Right(strRegKey, 1) <> "\" Then
  strRegKey = strRegKey & "\"
 End If
 For iTemp = 1 To Len(strSection)
  If Mid(strSection, iTemp, 1) <> "\" Then
   strRegKey = strRegKey & Mid(strSection, iTemp, 1)
  End If
 Next iTemp
 If Left(strRegKey, 1) = "\" Then
  strRegKey = Mid(strRegKey, 2)
 End If
 'strRegKey = BASEKEY & strRegKey
 'CreateNewKey strRegKey, HKEY_LOCAL_MACHINE
 saveKey lRegistryBase, strRegKey
 strValueName = strKey
 strReturn = ""
 strdata = CStr(vData)
 'SetKeyValue HKEY_LOCAL_MACHINE, strRegKey, strValueName, vSetting, REG_SZ
 SaveString lRegistryBase, strRegKey, strValueName, strdata
End Sub

Public Function alGetSettingReg(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vDefault As Variant) As Variant
 Dim bFlag As Boolean
 Dim strRegKey As String
 Dim strValueName As String
 Dim strReturn As String
 Dim iTemp As Integer
 Dim iStart As Integer
 Dim lRegistryBase As Long
 
 Select Case Left(UCase(strRegHive), 4)
  Case "HKLM"
   lRegistryBase = HKEY_LOCAL_MACHINE
  Case "HKCR"
   lRegistryBase = HKEY_CLASSES_ROOT
  Case Else
   lRegistryBase = HKEY_CURRENT_USER
 End Select
 strRegKey = strRegPath
 If Right(strRegKey, 1) <> "\" Then
  strRegKey = strRegKey & "\"
 End If
 For iTemp = 1 To Len(strAppname)
  If Mid(strAppname, iTemp, 1) <> "\" Then
   strRegKey = strRegKey & Mid(strAppname, iTemp, 1)
  End If
 Next iTemp
 If Right(strRegKey, 1) <> "\" Then
  strRegKey = strRegKey & "\"
 End If
 For iTemp = 1 To Len(strSection)
  If Mid(strSection, iTemp, 1) <> "\" Then
   strRegKey = strRegKey & Mid(strSection, iTemp, 1)
  End If
 Next iTemp
 If Left(strRegKey, 1) = "\" Then
  strRegKey = Mid(strRegKey, 2)
 End If
 strValueName = strKey
 strReturn = ""
 strReturn = getString(lRegistryBase, strRegKey, strValueName)
 alGetSettingReg = strReturn
End Function


