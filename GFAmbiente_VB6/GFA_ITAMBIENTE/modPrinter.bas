Attribute VB_Name = "modPrint"
Option Explicit



Private Declare Function FormatMessage Lib "kernel32" _
   Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
   ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long



'call SetWinDefPrinter (all the way at the bottom) with name of the printer.
'Setting Windows default printer API stuff
Declare Function GetProfileString Lib "kernel32" _
Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpKeyName As String, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long) As Long


Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long


Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As String) As Long


Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A


Public Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type


Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function SetDefaultPrinter Lib "winspool.drv" _
Alias "SetDefaultPrinterA" _
(ByVal pszPrinter As String) As Long

Public Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long


Public Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long


Public Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long


Public Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long


Public Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long


' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32


' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)


' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4


'Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1


Public Type DEVMODE
dmDeviceName As String * CCHDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCHFORMNAME
dmLogPixels As Integer
dmBitsPerPel As Long
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
dmICMMethod As Long ' // Windows 95 only
dmICMIntent As Long ' // Windows 95 only
dmMediaType As Long ' // Windows 95 only
dmDitherType As Long ' // Windows 95 only
dmReserved1 As Long ' // Windows 95 only
dmReserved2 As Long ' // Windows 95 only
End Type


Public Type PRINTER_INFO_5
pPrinterName As String
pPortName As String
Attributes As Long
DeviceNotSelectedTimeout As Long
TransmissionRetryTimeout As Long
End Type


'------CONSTANTS


'Constants for NT security
'Constants used to make changes to the values contained in the DevMode

Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_DUPLEX = &H1000&
Private Const DM_COPIES = &H100&
Public Const DMDUP_SIMPLEX = 1

Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&


Private Type PRINTER_DEFAULTS
'  The definition of Printer_Defaults in the VB5 API viewer is incorrect.
'  Below, pDevMode has been corrected to LONG.
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type


'------DECLARATIONS

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)



'The following is an unusual declaration of DocumentProperties:
'  pDevModeOutput and pDevModeInput are usually declared ByRef.  They are declared
'  ByVal in this program because we're using a Printer_Info_2 structure.
'  The pi2 structure contains a variable of type LONG which contains the address
'  of the DevMode structure (this is called a pointer).  This LONG variable must
'  be passed ByVal.
'  Normally this function is called with a BYTE ARRAY which contains the DEVMODE
'  structure and the Byte Array is passed ByRef.

Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Any, ByVal pDevModeInput As Any, ByVal fMode As Long) As Long

'------CODE


Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
        PtrCtoVbString = ""
    Else
        PtrCtoVbString = left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If

End Function


Private Sub SetDefaultPrint(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
Dim DeviceLine As String
Dim r As Long
Dim l As Long

DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
' Store the new printer information in the [WINDOWS] section of
' the WIN.INI file for the DEVICE= item
r = WriteProfileString("windows", "Device", DeviceLine)
' Cause all applications to reload the INI file:
l = SendMessageA(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub


Private Sub Win95SetDefaultPrinter(ChangePrinter As String)
Dim Handle As Long 'handle to printer
Dim PrinterName As String
Dim pd As PRINTER_DEFAULTS
Dim x As Long
Dim need As Long ' bytes needed
Dim pi5 As PRINTER_INFO_5 ' your PRINTER_INFO structure
Dim LastError As Long


' determine which printer was selected
PrinterName = ChangePrinter
' none - exit
If PrinterName = "" Then
Exit Sub
End If


' set the PRINTER_DEFAULTS members
pd.pDatatype = 0&
pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess


' Get a handle to the printer
x = OpenPrinter(PrinterName, Handle, pd)
' failed the open
If x = False Then
'error handler code goes here
Exit Sub
End If


' Make an initial call to GetPrinter, requesting Level 5
' (PRINTER_INFO_5) information, to determine how many bytes
' you need
x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
' don't want to check Err.LastDllError here - it's supposed
' to fail
' with a 122 - ERROR_INSUFFICIENT_BUFFER
' redim t as large as you need
ReDim t((need \ 4)) As Long


' and call GetPrinter for keepers this time
x = GetPrinter(Handle, 5, t(0), need, need)
' failed the GetPrinter
If x = False Then
'error handler code goes here
Exit Sub
End If


' set the members of the pi5 structure for use with SetPrinter.
' PtrCtoVbString copies the memory pointed at by the two string
' pointers contained in the t() array into a Visual Basic string.
' The other three elements are just DWORDS (long integers) and
' don't require any conversion
pi5.pPrinterName = PtrCtoVbString(t(0))
pi5.pPortName = PtrCtoVbString(t(1))
pi5.Attributes = t(2)
pi5.DeviceNotSelectedTimeout = t(3)
pi5.TransmissionRetryTimeout = t(4)


' this is the critical flag that makes it the default printer
pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT


' call SetPrinter to set it
x = SetPrinter(Handle, 5, pi5, 0)
' failed the SetPrinter
If x = False Then
MsgBox "SetPrinterFailed. Error code: " & Err.LastDllError
Exit Sub
End If


' and close the handle
ClosePrinter (Handle)


End Sub


Private Sub GetDriverAndPort(ByVal buffer As String, DriverName As _
String, PrinterPort As String)


Dim iDriver As Integer
Dim iPort As Integer
DriverName = ""
PrinterPort = ""


'The driver name is first in the string terminated by a comma
iDriver = InStr(buffer, ",")
If iDriver > 0 Then


'Strip out the driver name
DriverName = left(buffer, iDriver - 1)


'The port name is the second entry after the driver name
'separated by commas.
iPort = InStr(iDriver + 1, buffer, ",")


If iPort > 0 Then
'Strip out the port name
PrinterPort = Mid(buffer, iDriver + 1, _
iPort - iDriver - 1)
End If
End If
End Sub


Private Sub WinNTSetDefaultPrinter(ChangePrinter As String)
Dim buffer As String
Dim DeviceName As String
Dim DriverName As String
Dim PrinterPort As String
Dim PrinterName As String
Dim r As Long
'Get the printer information for the currently selected
'printer in the list. The information is taken from the
'WIN.INI file.
buffer = Space(1024)
PrinterName = ChangePrinter

r = GetProfileString("PrinterPorts", PrinterName, "", _
buffer, Len(buffer))


'Parse the driver name and port name out of the buffer
GetDriverAndPort buffer, DriverName, PrinterPort


If DriverName <> "" And PrinterPort <> "" Then
SetDefaultPrint PrinterName, DriverName, PrinterPort
End If
End Sub


Public Sub SetWinDefPrinter(ChangePrinter As String)
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer


osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)


If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
Call Win95SetDefaultPrinter(ChangePrinter)
Else
'This assumes that future versions of Windows use the NT method
SetDefaultPrinter ChangePrinter
End If
End Sub




Public Sub RotatedText(myObject As Object, _
  ptext As String, rotation As Long, x As Single, y As Single, fontSize As Single)

Dim font As LOGFONT
Dim prevFont As Long, hFont As Long, ret As Long
Dim hdc As Long

myObject.CurrentX = x
myObject.CurrentY = y
With font
  .lfFaceName = Chr(0)
  .lfEscapement = rotation
  .lfFaceName = myObject.FontName & Chr$(0)
  myObject.fontSize = fontSize
  .lfHeight = (myObject.fontSize * -20) / Screen.TwipsPerPixelY
  If myObject.FontBold = True Then
    .lfWeight = 700
  Else
    .lfWeight = 400
  End If
  .lfItalic = myObject.FontItalic
  .lfStrikeOut = myObject.FontStrikethru
  .lfUnderline = myObject.FontUnderline
End With
hFont = CreateFontIndirect(font)
prevFont = SelectObject(myObject.hdc, hFont)

myObject.Print ptext

ret = SelectObject(myObject.hdc, prevFont)
ret = DeleteObject(hFont)
End Sub


Public Function setDefaultSettingPrint(hWnd As Long, PrinterName As String, ByVal orientation As Integer, ByVal size As Integer, ByVal paperlength As Long, MyDevModeOLD As DEVMODE) As Boolean
   Dim PrinterHandle As Long
   Dim pd As PRINTER_DEFAULTS
   Dim MyDevMode As DEVMODE
   Dim Result As Long
   Dim Needed As Long
   Dim pFullDevMode As Long
   Dim pi2_buffer() As Long
   
   On Error GoTo setDefaultSettingPrint
   'This is a block of memory for the     Printer_Info_2 structure
   'If you need to use the Printer_Info_2 User Defined Type, the
   '  definition of Printer_Info_2 in the API viewer is incorrect.
   '  pDevMode and pSecurityDescriptor should be defined As Long.
   pd.pDatatype = vbNullString
   pd.pDevMode = 0&
   'Printer_Access_All is required for NT security
   pd.DesiredAccess = PRINTER_ALL_ACCESS
   Result = OpenPrinter(PrinterName, PrinterHandle, pd)
   If Err.LastDllError <> 0 And Err.LastDllError <> 13 Then Err.Raise Err.LastDllError, "OpenPrinter", GetLastErrorDescription(Err.LastDllError)

   'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
   'This value is divided by 4 since each element of pi2_buffer is a long.
   Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
   
   ReDim pi2_buffer((Needed \ 4))
   Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
   'The seventh element of pi2_buffer is a Pointer to a block of memory
   '  which contains the full DevMode (including the PRIVATE portion).
   pFullDevMode = pi2_buffer(7)
   'Copy the Public portion of FullDevMode into our DevMode structure
   Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))
   Call CopyMemory(MyDevModeOLD, ByVal pFullDevMode, Len(MyDevModeOLD))
   'Make desired changes
   MyDevMode.dmDuplex = DMDUP_SIMPLEX
   MyDevMode.dmFields = DM_DUPLEX Or DM_ORIENTATION
   If orientation > 0 Then MyDevMode.dmOrientation = orientation
   MyDevMode.dmPaperSize = size
   'MyDevMode.dmPaperLength = 20160
   MyDevMode.dmPaperLength = paperlength
   MyDevMode.dmFields = MyDevMode.dmFields Or DM_PAPERSIZE Or DM_PAPERLENGTH
   'Copy our DevMode structure back into FullDevMode
   Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))
   'Copy our changes to "the PUBLIC portion of the DevMode" into "the PRIVATE portion of the DevMode"
   Result = DocumentProperties(hWnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
   'Update the printer's default properties (to verify, go to the Printer folder
   '  and check the properties for the printer)
   Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
   Call ClosePrinter(PrinterHandle)
   
   Exit Function
setDefaultSettingPrint:
   Err.Raise 9999, "setDefaultSettingPrint", Err.Description
End Function







Public Function GetLastErrorDescription(ByVal lLastError As Long) As String
Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF
Const LANG_USER_DEFAULT As Long = &H400


Dim sBuffer As String, lResult As Long


sBuffer$ = String$(1024, vbNullChar)


lResult& = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                         FORMAT_MESSAGE_IGNORE_INSERTS Or _
                         FORMAT_MESSAGE_MAX_WIDTH_MASK, _
                         ByVal 0&, lLastError&, LANG_USER_DEFAULT, _
                         sBuffer$, Len(sBuffer$), ByVal 0&)


If lResult& Then GetLastErrorDescription$ = Trim$(Mid$(sBuffer$, 1, lResult&))


End Function




