Attribute VB_Name = "modOtherStuff"
Option Private Module
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Global gstrRegistryPath As String

Global gstrRegistryKey As String    '
Global gstrRegistryHive As String   ' Added Jul/2002 --- Nick Hall
Global gb_UseDataLock As Boolean    '
       
' added by bjones 8/29/2002. replces old GetVolumeInformation that was bad
Public Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
    
' Computer name

Private Declare Function GetComputerName& Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)

Private Const MAX_PATH = 260
Private Const MAX_COMPUTERNAME_LENGTH = 15

' Registry stuff

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const ERROR_NONE = 0
Global Const KEY_ALL_ACCESS = &H3F

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Function ComputerName() As String
  Dim s$, sz&, dl&
  sz& = MAX_COMPUTERNAME_LENGTH + 1
  s$ = String$(sz&, 0)
  dl& = GetComputerName(s$, sz&)
  ComputerName = s$
  
End Function

Public Function HDSerialNumber(strDriveLetter As String) As String
' added by bjones 8/29/2002. replaces old driveSerial that was bad
'   Function to return the serial number for a hard drive
'   Accepts:
'       strDriveLetter - a valid drive letter for the PC, in the format "C:\"
'   Returns:
'       The serial number for the drive, formatted as "xxxx-xxxx"
    
    Dim lngReturn As Long, lngDummy1 As Long, lngDummy2 As Long, lngSerial As Long
    Dim strDummy1 As String, strDummy2 As String, strSerial As String
    
    strDummy1 = Space(MAX_PATH)
    strDummy2 = Space(MAX_PATH)
    lngReturn = apiGetVolumeInformation(strDriveLetter, strDummy1, Len(strDummy1), lngSerial, lngDummy1, lngDummy2, strDummy2, Len(strDummy2))
    strSerial = Trim(Hex(lngSerial))
    strSerial = String(8 - Len(strSerial), "0") & strSerial
    strSerial = Left(strSerial, 4) & "-" & Right(strSerial, 4)
    HDSerialNumber = strSerial
End Function

Public Function WindowsProductKey() As String
' bjones, 8/29/2002. changed reg key from ProductKey to ProductId

    Dim strKey As String
  
    strKey = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductId")
    If strKey = "" Then
        strKey = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId")
    End If
    WindowsProductKey = strKey
End Function

Public Sub SoftwareNameError()
    Dim Msg As String
    Msg = "You haven't assigned the software name yet." & vbCrLf _
        & "Solution: Assign the software name property first." _
        & "Example: ActiveLock1.SoftwareName=""MyApp"""
    Err.Raise vbObjectError + 2, "ActiveLock", "ActiveLock Error"
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant

    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)

    QueryValue = vValue
    RegCloseKey (hKey)
End Function
