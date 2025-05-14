VERSION 5.00
Begin VB.UserControl ActiveLock 
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ActiveLock1.ctx":0000
   PropertyPages   =   "ActiveLock1.ctx":030A
   ScaleHeight     =   465
   ScaleWidth      =   480
   ToolboxBitmap   =   "ActiveLock1.ctx":037E
End
Attribute VB_Name = "ActiveLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Public Enum THastType 'Added Apr/2002 --- Francois de Wet
  htSHA1AA1 = 0
  htSHA1AA2 = 1
  htMD5AA1 = 2
  htMD5AA2 = 3
  htMD5AB1 = 4
  htMD5AB2 = 5
End Enum

Private ht_HashAlgorithm As THastType

Private mstr_LockToCustomString As String

Private mstr_SoftwareName As String
Private mi_LiberationKeyLength As Integer
Private mi_SoftwareCodeLength As Integer

Private mstr_SoftwarePassword As String ' added Aug 27,2002---Walter White
Private mstr_RegistryName As String ' added Aug 28,2002---Walter White


Private mb_LockToHardDrive As Boolean
Private mb_LockToWindowsSerial As Boolean
Private mb_LockToRandomNumber As Boolean
Private mb_LockToComputerName As Boolean
Private mb_LockToMACAddress As Boolean

Private mb_UseDataLock As Boolean

Private mstr_RegistryPath As String '
Private mstr_RegistryKey As String  ' Added Jul/2002 --- Nick Hall
Private mstr_RegistryHive As String '


' Added Aug 23, 2002 --- Wade Grimm
'Begin Add
Private mstr_RegKeyCounter As String ' Added Aug 23, 2002 --- Wade Grimm
Private mstr_RegKeyLibKey As String ' Added Aug 23, 2002 --- Wade Grimm
Private mstr_RegKeyLastRunDate As String ' Added Aug 23, 2002 --- Wade Grimm
Private mstr_RegKeyInitialRunDate As String ' Added Aug 23, 2002 --- Wade Grimm
Private mstr_RegKeyRandomKey As String ' Added Aug 23, 2002 --- Wade Grimm
Private mstr_BlowFishEncryptKey As String ' Added Aug 29, 2002 --- Wade Grimm
Private mb_EncRegKeys As Boolean ' Added Aug 29, 2002 --- Wade Grimm
'End Add

Private Const SOFTWARECODE_MAXLength = 16
Private Const LIBERATIONKEY_MAXLength = 16
Private Const SOFTWARENAME_MAXLength = 40
Private Const Softwarepassword_MaxLength = 40
Private Const LockToCustomString_MAXLength = 255

Public Event Registration(WasSuccessful As Boolean)
Attribute Registration.VB_Description = "This event is raised when the LiberationKey is changed."
Attribute Registration.VB_MemberFlags = "40"
Public Event Transference(WasSuccessful As Boolean, LiberationKey As String)
Attribute Transference.VB_Description = "This event is raised after the Transfer method is called. If succesful, it returns the LiberationKey that will unlock the other computer."
Attribute Transference.VB_MemberFlags = "40"
Public Event InvalidDate()

Private Property Let RegEncKey(ByVal vData As Boolean)
Dim tmpCounter As String, tmpLastRunDate As String, tmpInitRunDate As String, tmpRnd As String
Dim bF As BlowFish
Set bF = New BlowFish

  mb_EncRegKeys = vData
  'alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "EncKey", mb_EncRegKeys
  
  If mb_EncRegKeys = True Then
      tmpCounter = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyCounter, "")
      If tmpCounter <> "" Then
        tmpCounter = bF.EncryptString(tmpCounter, EncKey)
        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyCounter, tmpCounter
      End If
      
      tmpLastRunDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLastRunDate, "")
      If tmpLastRunDate <> "" Then
        tmpLastRunDate = bF.EncryptString(tmpLastRunDate, EncKey)
        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLastRunDate, tmpLastRunDate
      End If
    
      tmpInitRunDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, "")
      If tmpInitRunDate <> "" Then
        tmpInitRunDate = bF.EncryptString(tmpInitRunDate, EncKey)
        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, tmpInitRunDate
      End If
      
      tmpRnd = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, "")
      If tmpRnd <> "" Then
        tmpRnd = bF.EncryptString(tmpRnd, EncKey)
        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, tmpRnd
      End If
'    Else
'      tmpCounter = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyCounter, "")
'      If tmpCounter <> "" Then
'        tmpCounter = bF.DecryptString(tmpCounter, EncKey)
'        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyCounter, tmpCounter
'      End If
'
'      tmpLastRunDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLastRunDate, "")
'      If tmpLastRunDate <> "" Then
'        tmpLastRunDate = bF.DecryptString(tmpLastRunDate, EncKey)
'        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLastRunDate, tmpLastRunDate
'      End If
'
'      tmpInitRunDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, "")
'      If tmpInitRunDate <> "" Then
'        tmpInitRunDate = bF.DecryptString(tmpInitRunDate, EncKey)
'        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, tmpInitRunDate
'      End If
'
'      tmpRnd = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, "")
'      If tmpRnd <> "" Then
'        tmpRnd = bF.DecryptString(tmpRnd, EncKey)
'        alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, tmpRnd
'      End If
    End If
  PropertyChanged "RegEncKey"
End Property

Private Property Get RegEncKey() As Boolean
  mb_EncRegKeys = True 'alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "EncKey", False)
 
  RegEncKey = mb_EncRegKeys
End Property

Public Property Let EncKey(ByVal vData As String)
  mstr_BlowFishEncryptKey = vData
  alSaveSetting mstr_RegistryHive, mstr_RegistryKey, _
  mstr_RegistryName, mstr_RegistryPath, "RegEncKey", mstr_BlowFishEncryptKey

  PropertyChanged "EncKey"
End Property

Public Property Get EncKey() As String
  mstr_BlowFishEncryptKey = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, _
  mstr_RegistryPath, "RegEncKey", "Default")
  EncKey = mstr_BlowFishEncryptKey
End Property

Public Property Let RegistryName(ByVal vData As String)
  mstr_RegistryName = vData
  PropertyChanged "RegistryName"
End Property

Public Property Get RegistryName() As String
  RegistryName = mstr_RegistryName
End Property

Public Property Let RegCounterKey(ByVal vData As String)
  mstr_RegKeyCounter = vData
  PropertyChanged "RegCounterKey"
End Property

Public Property Get RegCounterKey() As String
  RegCounterKey = mstr_RegKeyCounter
End Property

Public Property Let RegLiberationKey(ByVal vData As String)
  mstr_RegKeyLibKey = vData
  PropertyChanged "RegLiberationKey"
End Property

Public Property Get RegLiberationKey() As String
  RegLiberationKey = mstr_RegKeyLibKey
End Property

Public Property Let RegLastRunDateKey(ByVal vData As String)
  mstr_RegKeyLastRunDate = vData
  PropertyChanged "RegLastRunDateKey"
End Property

Public Property Get RegLastRunDateKey() As String
  RegLastRunDateKey = mstr_RegKeyLastRunDate
End Property

Public Property Let RegInitialRunDateKey(ByVal vData As String)
  mstr_RegKeyInitialRunDate = vData
  PropertyChanged "RegInitialRunDateKey"
End Property

Public Property Get RegInitialRunDateKey() As String
  RegInitialRunDateKey = mstr_RegKeyInitialRunDate
End Property

Public Property Let RegRandomKey(ByVal vData As String)
  mstr_RegKeyRandomKey = vData
  PropertyChanged "RegRandomKey"
End Property

Public Property Get RegRandomKey() As String
  RegRandomKey = mstr_RegKeyRandomKey
End Property

Public Property Let HashAlgorithm(ByVal vData As THastType)
    ht_HashAlgorithm = vData
    PropertyChanged "HashAlgorithm"
End Property

Public Property Get HashAlgorithm() As THastType
    HashAlgorithm = ht_HashAlgorithm
End Property

Private Function Hash(strHashThis As String) As String
  ' Allow different hash types

  Select Case ht_HashAlgorithm
    Case htSHA1AA1: Hash = SHA1AA1Hash(strHashThis)
    Case htSHA1AA2: Hash = SHA1AA2Hash(strHashThis)
    Case htMD5AA1: Hash = MD5AA1Hash(strHashThis)
    Case htMD5AA2: Hash = MD5AA2Hash(strHashThis)
    Case htMD5AB1: Hash = MD5AB1Hash(strHashThis)
    Case htMD5AB2: Hash = MD5AB2Hash(strHashThis)
    Case Else: Hash = SHA1AA1Hash(strHashThis) ' Default type
  End Select

End Function

Public Sub Reset()
Attribute Reset.VB_Description = "Resets ActiveLock's properties."
  ' Reset ActiveLock properties
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "RandomKey", ""
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "LiberationKey", ""
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "Counter", ""
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "InitialDate", ""
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, "LastRunDate", ""
    PropertyChanged "SoftwareCode"
    PropertyChanged "LiberationKey"
    PropertyChanged "Counter"
    PropertyChanged "InitialDate"
    PropertyChanged "LastRunDate"
  End If
  ' TO-DO: Remove registry key
End Sub

Public Sub Transfer(OtherSoftwareCode As String)
Attribute Transfer.VB_Description = "Transfers the current license to another computer.\r\nThis method requires that you've set the LockToRandomNumber property, since the SoftwareCode is supposed to change."
  ' Transfer license to another computer

  ' Pre-requisites: RegisteredUser = True And mb_LockToRandomNumber = True
  If Not (mb_LockToRandomNumber) Or Not (RegisteredUser) Then
    RaiseEvent Transference(False, "")
    Exit Sub
  End If
  
  ' 1. This computer's license is cancelled
  Dim strRnd As Long
  Randomize
  strRnd = CStr(CLng(Rnd(1) * 2147483647))

  alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, strRnd
  
  ' 2. A new license is generated to the other computer
  Dim strHash As String
  strHash = Hash(OtherSoftwareCode & mstr_SoftwareName)
  strHash = Left(strHash, LIBERATIONKEY_MAXLength)
  
  ' 3. Raise ActiveLock_Transference event
  RaiseEvent Transference(True, strHash)
End Sub

Public Sub Register(LiberationKey As String)
Attribute Register.VB_Description = "Changes the  LiberationKey."
  ' Stores the liberation code in the registry
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    LiberationKey = Left(LiberationKey, LIBERATIONKEY_MAXLength)
    
    alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLibKey, LiberationKey

    PropertyChanged "LiberationKey"
  End If
  RaiseEvent Registration(RegisteredUser) ' True or False
End Sub

Public Property Let LiberationKey(ByVal vData As String)
Attribute LiberationKey.VB_Description = "This is the key that will unlock your software. The user must set the correct value in order to have the software registered.\r\nIt's important to note that each SoftwareCode needs a different LiberationKey."
Attribute LiberationKey.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
Attribute LiberationKey.VB_MemberFlags = "440"
  ' For compatibility purposes only. Use method Register instead.
  Register (vData)
End Property

Public Property Let LockToComputerName(ByVal vData As Boolean)
  mb_LockToComputerName = vData
  PropertyChanged "LockToComputerName"
End Property

Public Property Get LockToComputerName() As Boolean
Attribute LockToComputerName.VB_Description = "Set this property to make the SoftwareCode depend on the computer name."
Attribute LockToComputerName.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToComputerName = mb_LockToComputerName
End Property

Public Property Let LockToHardDrive(ByVal vData As Boolean)
  mb_LockToHardDrive = vData
  PropertyChanged "LockToHardDrive"
End Property

Public Property Get LockToHardDrive() As Boolean
Attribute LockToHardDrive.VB_Description = "Set this property to make the SoftwareCode depend on the hard drive serial number."
Attribute LockToHardDrive.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToHardDrive = mb_LockToHardDrive
End Property

Public Property Let LockToWindowsSerial(ByVal vData As Boolean)
  mb_LockToWindowsSerial = vData
  PropertyChanged "LockToWindowsSerial"
End Property

Public Property Get LockToWindowsSerial() As Boolean
Attribute LockToWindowsSerial.VB_Description = "Set this property to make the SoftwareCode depend on the Windows serial number."
Attribute LockToWindowsSerial.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToWindowsSerial = mb_LockToWindowsSerial
End Property

Public Property Let LockToRandomNumber(ByVal vData As Boolean)
  mb_LockToRandomNumber = vData
  PropertyChanged "LockToRandomNumber"
End Property

Public Property Get LockToRandomNumber() As Boolean
Attribute LockToRandomNumber.VB_Description = "Set this property to make the SoftwareCode depend on an unique random number. \r\nThis is necessary if you want that the SoftwareCode changes when you reset ActiveLock."
Attribute LockToRandomNumber.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToRandomNumber = mb_LockToRandomNumber
End Property

Public Property Let LockToMACAddress(ByVal vData As Boolean)
  mb_LockToMACAddress = vData
  PropertyChanged "LockToMACAddress"
End Property

Public Property Get LockToMACAddress() As Boolean
  LockToMACAddress = mb_LockToMACAddress
End Property

Public Property Let UseDataLock(ByVal vData As Boolean)
  mb_UseDataLock = vData
  gb_UseDataLock = mb_UseDataLock
  PropertyChanged "UseDataLock"
End Property

Public Property Get UseDataLock() As Boolean
Attribute UseDataLock.VB_ProcData.VB_Invoke_Property = "PropertyPage6"
  UseDataLock = mb_UseDataLock
End Property

Public Property Let RegistryHive(ByVal vData As String)
  'If Right(vData, 1) <> "\" Then vData = vData & "\"
  mstr_RegistryHive = vData
  gstrRegistryHive = mstr_RegistryHive
  PropertyChanged "RegistryHive"
End Property

Public Property Get RegistryHive() As String
  RegistryHive = mstr_RegistryHive
End Property

Public Property Let RegistryKey(ByVal vData As String)
  'If Right(vData, 1) <> "\" Then vData = vData & "\"
  mstr_RegistryKey = vData
  gstrRegistryKey = mstr_RegistryKey
  PropertyChanged "RegistryKey"
End Property

Public Property Get RegistryKey() As String
  RegistryKey = mstr_RegistryKey
End Property

Public Property Let RegistryPath(ByVal vData As String)
  'If Right(vData, 1) <> "\" Then vData = vData & "\"
  mstr_RegistryPath = vData
  gstrRegistryPath = mstr_RegistryPath
  PropertyChanged "RegistryPath"
End Property

Public Property Get RegistryPath() As String
Attribute RegistryPath.VB_Description = "Returns/sets the path in the registry where the ActiveLock information is stored."
Attribute RegistryPath.VB_ProcData.VB_Invoke_Property = "PropertyPage2;Text"
  RegistryPath = mstr_RegistryPath
End Property

Public Property Get SoftwareCode() As String
Attribute SoftwareCode.VB_Description = "Returns the code that uniquely identifies each installation of your software. Read-only.\r\n"
Attribute SoftwareCode.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute SoftwareCode.VB_MemberFlags = "400"
  ' The SoftwareCode is a hash based on several
  ' properties (depending on the LockTo... values)
  
  Dim strParam As String
  Dim strCode As String
  
  If mb_LockToRandomNumber Then strParam = strParam & RandomNumber
  If mb_LockToWindowsSerial Then strParam = strParam & WindowsProductKey
  If mb_LockToHardDrive Then strParam = strParam & HDSerialNumber("C:\") ' bjones 8/29/2002. replaces old DriveSerial. getVolumeInformation api also replaced
  If mb_LockToComputerName Then strParam = strParam & ComputerName
  If mb_LockToMACAddress Then strParam = strParam & GetMACAddress()

  ' If no parameters selected, use default ones
  If strParam = "" Then strParam = RandomNumber & WindowsProductKey
  
  strCode = Hash(strParam & mstr_LockToCustomString)  ' use LockToCustomString
  
  ' Return the first characters (depends on SoftwareCodeLength)
  SoftwareCode = Left(strCode, mi_SoftwareCodeLength)
  PropertyChanged "SoftwareCode"
End Property

Private Sub UserControl_InitProperties()
  ' Set module variables (default values)
  
  mstr_SoftwareName = "AppName"
  mstr_SoftwarePassword = "password"
  mi_LiberationKeyLength = 16
  mi_SoftwareCodeLength = 6
  mb_LockToWindowsSerial = True
  mb_LockToRandomNumber = True ' Required by Transference method
  mstr_RegistryName = "MyRegName"                            '
  mstr_RegistryPath = "ActiveLock"                             '
  mstr_RegistryKey = "VB and VBA Program Settings"   ' Added Jul/2002 --- Nick Hall
  mstr_RegistryHive = "HKCU"                                    '
  'Begin Encryption Stuff ; Added by Wade Grimm
  mstr_RegKeyCounter = "Counter"
  mstr_RegKeyLibKey = "LiberationKey"
  mstr_RegKeyLastRunDate = "LastRunDate"
  mstr_RegKeyInitialRunDate = "InitialRunDate"
  mstr_RegKeyRandomKey = "RandomKey"
  mstr_BlowFishEncryptKey = "Default"
  RegEncKey = True
  ' End Encryption Stuff
  ' Set global variables (default values)
  gstrRegistryPath = mstr_RegistryPath
  gstrRegistryKey = mstr_RegistryKey
  gstrRegistryHive = mstr_RegistryHive
  gb_UseDataLock = mb_UseDataLock
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Get module variables from PropBag
  
  mstr_SoftwareName = PropBag.ReadProperty("SoftwareName")
  mstr_SoftwarePassword = PropBag.ReadProperty("SoftwarePassword")
  mi_LiberationKeyLength = PropBag.ReadProperty("LiberationKeyLength")
  mi_SoftwareCodeLength = PropBag.ReadProperty("SoftwareCodeLength")
  mb_LockToHardDrive = PropBag.ReadProperty("LockToHardDrive")
  mb_LockToWindowsSerial = PropBag.ReadProperty("LockToWindowsSerial")
  mb_LockToRandomNumber = PropBag.ReadProperty("LockToRandomNumber")
  mb_LockToComputerName = PropBag.ReadProperty("LockToComputerName")
  mb_LockToMACAddress = PropBag.ReadProperty("LockToMACAddress")
  mb_UseDataLock = PropBag.ReadProperty("UseDataLock")
  mstr_RegistryPath = PropBag.ReadProperty("RegistryPath")
  mstr_RegistryKey = PropBag.ReadProperty("RegistryKey")
  mstr_RegistryName = PropBag.ReadProperty("RegistryName")
  mstr_RegistryHive = PropBag.ReadProperty("RegistryHive")
  mstr_LockToCustomString = PropBag.ReadProperty("LockToCustomString")
  ht_HashAlgorithm = PropBag.ReadProperty("HashAlgorithm")

  mstr_RegKeyCounter = PropBag.ReadProperty("RegCounterKey")
  mstr_RegKeyLibKey = PropBag.ReadProperty("RegLiberationKey")
  mstr_RegKeyLastRunDate = PropBag.ReadProperty("RegLastRunDateKey")
  mstr_RegKeyInitialRunDate = PropBag.ReadProperty("RegInitialRunDateKey")
  mstr_RegKeyRandomKey = PropBag.ReadProperty("RegRandomKey")
  mstr_BlowFishEncryptKey = PropBag.ReadProperty("EncKey")
  mb_EncRegKeys = PropBag.ReadProperty("RegEncKey")
  
  ' Set global variables
  gstrRegistryPath = mstr_RegistryPath
  gstrRegistryKey = mstr_RegistryKey
  gstrRegistryHive = mstr_RegistryHive
  gb_UseDataLock = mb_UseDataLock
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' Write properties into PropBag
  PropBag.WriteProperty "SoftwareName", mstr_SoftwareName
  PropBag.WriteProperty "SoftwarePassword", mstr_SoftwarePassword
  PropBag.WriteProperty "LiberationKeyLength", mi_LiberationKeyLength
  PropBag.WriteProperty "SoftwareCodeLength", mi_SoftwareCodeLength
  PropBag.WriteProperty "LockToHardDrive", mb_LockToHardDrive
  PropBag.WriteProperty "LockToWindowsSerial", mb_LockToWindowsSerial
  PropBag.WriteProperty "LockToRandomNumber", mb_LockToRandomNumber
  PropBag.WriteProperty "LockToComputerName", mb_LockToComputerName
  PropBag.WriteProperty "LockToMACAddress", mb_LockToMACAddress
  PropBag.WriteProperty "UseDataLock", mb_UseDataLock
  PropBag.WriteProperty "RegistryPath", mstr_RegistryPath
  PropBag.WriteProperty "RegistryKey", mstr_RegistryKey
  PropBag.WriteProperty "RegistryName", mstr_RegistryName
  PropBag.WriteProperty "RegistryHive", mstr_RegistryHive
  PropBag.WriteProperty "LockToCustomString", mstr_LockToCustomString
  PropBag.WriteProperty "HashAlgorithm", ht_HashAlgorithm
  
  PropBag.WriteProperty "RegCounterKey", mstr_RegKeyCounter
  PropBag.WriteProperty "RegLiberationKey", mstr_RegKeyLibKey
  PropBag.WriteProperty "RegLastRunDateKey", mstr_RegKeyLastRunDate
  PropBag.WriteProperty "RegInitialRunDateKey", mstr_RegKeyInitialRunDate
  PropBag.WriteProperty "RegRandomKey", mstr_RegKeyRandomKey
  PropBag.WriteProperty "EncKey", mstr_BlowFishEncryptKey
  PropBag.WriteProperty "RegEncKey", mb_EncRegKeys

  
End Sub

Public Property Get Counter() As Long
Attribute Counter.VB_Description = "Returns the number of times this software has been run in the user's machine. Read-only."
Attribute Counter.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Counter.VB_MemberFlags = "400"
  ' Set and returns Counter + 1 everytime it is called
  'Code changed by Mads Moen 03.19.2003
  On Error Resume Next
  Dim varCounter, bF As BlowFish
  Set bF = New BlowFish
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varCounter = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyCounter, "")
    
    If RegEncKey Then
        varCounter = bF.DecryptString(CStr(varCounter), EncKey)
    End If
    
    If varCounter = "" Then varCounter = 0
    
    Counter = CLng(varCounter) + 1
    
        If RegEncKey Then
            varCounter = bF.EncryptString(CStr(Counter), EncKey)
            alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, _
            mstr_RegKeyCounter, varCounter
        Else
            alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, _
            mstr_RegKeyCounter, varCounter
        End If

  End If

  PropertyChanged "Counter"

End Property

Public Property Get UsedDays() As Long
Attribute UsedDays.VB_Description = "Returns the number of days since the first run date. Read-only."
Attribute UsedDays.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute UsedDays.VB_MemberFlags = "400"
  ' Count the number of days since first run time
  On Error Resume Next
  Dim varInitialDate As Variant, bF As BlowFish
  Set bF = New BlowFish
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varInitialDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, "")
    If RegEncKey Then
        varInitialDate = bF.DecryptString(CStr(varInitialDate), EncKey)
    End If

    If varInitialDate = "" Then
      ' If it's not in the registry
      varInitialDate = Now
    
    If RegEncKey Then
        varInitialDate = bF.EncryptString(CStr(varInitialDate), EncKey)
    End If

      alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyInitialRunDate, varInitialDate
    End If
    ' Return the number of days since initial date
    UsedDays = CLng(DateDiff("d", varInitialDate, Now))
    PropertyChanged "UsedDays"
    ' See if the user have changed the time settings
    If LastRunDate > Now Then RaiseEvent InvalidDate
  End If
End Property

Public Property Get RegisteredUser() As Boolean
Attribute RegisteredUser.VB_Description = "Returns a boolean value indicating if the user is correctly registered. Read-only."
Attribute RegisteredUser.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute RegisteredUser.VB_MemberFlags = "400"
  ' Compares LiberationKey with Hash(SoftwareCode & SoftwareName)
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    Dim strHash As String, strKey As String
    strHash = Hash(SoftwareCode & mstr_SoftwarePassword & mstr_SoftwareName)
    
    strKey = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLibKey, "")

    If UCase(Left(strHash, mi_LiberationKeyLength)) = UCase(strKey) Then
        RegisteredUser = True
    Else
        RegisteredUser = False
    End If
    PropertyChanged "RegisteredUser"
  End If
End Property

Public Property Let SoftwareName(ByVal vData As String)
  mstr_SoftwareName = Left(vData, SOFTWARENAME_MAXLength)
  PropertyChanged "SoftwareName"
End Property

Public Property Get SoftwareName() As String
Attribute SoftwareName.VB_Description = "IMPORTANT ! Set this property to identify your software."
Attribute SoftwareName.VB_ProcData.VB_Invoke_Property = "PropertyPage1;Text"
  SoftwareName = mstr_SoftwareName
End Property

Public Property Let SoftwarePassword(ByVal vData As String)
  mstr_SoftwarePassword = Left(vData, Softwarepassword_MaxLength)
  PropertyChanged "SoftwarePassword"
End Property

Public Property Get SoftwarePassword() As String
  SoftwarePassword = mstr_SoftwarePassword
End Property
Public Property Let LockToCustomString(ByVal vData As String)
  mstr_LockToCustomString = Left(vData, LockToCustomString_MAXLength)
  PropertyChanged "LockToCustomString"
End Property

Public Property Get LockToCustomString() As String
  LockToCustomString = mstr_LockToCustomString
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
  frmAbout.Show 1
End Sub

Private Sub UserControl_Paint()
  UserControl.Line (15, 15)-(450, 15), &HFFFFFF
  UserControl.Line (450, 15)-(450, 450), &H808080
  UserControl.Line (450, 450)-(15, 450), &H808080
  UserControl.Line (15, 450)-(15, 15), &HFFFFFF
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = 465
  UserControl.Width = 480
End Sub

Public Property Get LiberationKeyLength() As Integer
Attribute LiberationKeyLength.VB_Description = "Returns/sets the length of the LiberationKey."
Attribute LiberationKeyLength.VB_ProcData.VB_Invoke_Property = "PropertyPage3;Behavior"
  LiberationKeyLength = mi_LiberationKeyLength
End Property

Public Property Let LiberationKeyLength(ByVal vNewValue As Integer)
  If vNewValue > LIBERATIONKEY_MAXLength Then vNewValue = LIBERATIONKEY_MAXLength
  If vNewValue < 1 Then vNewValue = 1
  mi_LiberationKeyLength = vNewValue
  PropertyChanged "LiberationKeyLength"
End Property

Public Property Get SoftwareCodeLength() As Integer
Attribute SoftwareCodeLength.VB_Description = "Returns/sets the length of the SoftwareCode."
Attribute SoftwareCodeLength.VB_ProcData.VB_Invoke_Property = "PropertyPage3;Behavior"
  SoftwareCodeLength = mi_SoftwareCodeLength
End Property

Public Property Let SoftwareCodeLength(ByVal vNewValue As Integer)
  If vNewValue > SOFTWARECODE_MAXLength Then vNewValue = SOFTWARECODE_MAXLength
  If vNewValue < 1 Then vNewValue = 1
  mi_SoftwareCodeLength = vNewValue
  PropertyChanged "SoftwareCodeLength"
End Property

Private Function RandomNumber() As String
  ' Generates a random number, which is stored in the registry
  ' (This random number will be always the same, for a given SoftwareName)
  On Error Resume Next
  Dim strRnd As String, bF As BlowFish
  Set bF = New BlowFish
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    strRnd = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, "")
    
    If RegEncKey Then
        strRnd = bF.DecryptString(strRnd, EncKey)
    End If

    If strRnd = "" Then
      ' If there isn't a random number, generate one
      Randomize
      strRnd = CStr(CLng(Rnd(1) * 2147483647))
      
    If RegEncKey Then
        strRnd = bF.EncryptString(strRnd, EncKey)
    End If

      alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyRandomKey, strRnd

    End If
    ' If there is a random number, return it
    RandomNumber = strRnd
  End If
End Function

Public Property Get LastRunDate() As Date
Attribute LastRunDate.VB_Description = "Returns the last date when the program was executed. Read-only."
Attribute LastRunDate.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute LastRunDate.VB_MemberFlags = "440"
  ' Returns last run date. If LastRunDate < Now(), user may be cheating
  On Error Resume Next
  Dim varLastRunDate As Variant, bF As BlowFish
  Set bF = New BlowFish
  Dim Agora As Date
  Agora = Now
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varLastRunDate = alGetSetting(mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, _
    mstr_RegistryPath, mstr_RegKeyLastRunDate, Agora)
    If RegEncKey Then
        varLastRunDate = bF.DecryptString(CStr(varLastRunDate), EncKey)
    End If

    If Not IsDate(varLastRunDate) Then varLastRunDate = Agora
    ' Store current date only if LastRunDate < Now()
    
    If varLastRunDate <= Agora Then
        If RegEncKey Then
          varLastRunDate = bF.EncryptString(CStr(varLastRunDate), EncKey)
          alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, mstr_RegistryPath, mstr_RegKeyLastRunDate, varLastRunDate
        Else
          alSaveSetting mstr_RegistryHive, mstr_RegistryKey, mstr_RegistryName, _
          mstr_RegistryPath, mstr_RegKeyLastRunDate, Agora
        End If
    End If
    LastRunDate = varLastRunDate
    PropertyChanged "LastRunDate"
  End If
End Property

'' THE END ''
