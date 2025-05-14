Attribute VB_Name = "CreaProcesso"
Option Explicit

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0

Public Sub ExecuteAndWait(cmdline$)
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim X As Long

    NameStart.cb = Len(NameStart)
    X = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
    X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    X = CloseHandle(NameOfProc.hProcess)
End Sub


