Attribute VB_Name = "modSpyScan"
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

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" _
 (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" _
 Alias "CreateFileA" _
 (ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, _
 ByVal dwShareMode As Long, _
 lpSecurityAttributes As Any, _
 ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, _
 ByVal hTemplateFile As Long) As Long
 
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

' File attributes
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
 
Public Function DoScan() As Boolean
'/*
'Will scan the memory for common debuggers
'or File Monitoring software
'*\

Dim hFile As Long, retVal As Long
Dim sScan As String
Dim sBuffer As String

    Dim sRegMonClass As String, sFileMonClass As String
    '\\We break up the class names to avoid
    '     detection in a hex editor
    sRegMonClass = "R" & "e" & "g" & "m" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    sFileMonClass = "F" & "i" & "l" & "e" & "M" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    '\\See if RegMon or FileMon are running


    Select Case True
        Case FindWindow(sRegMonClass, vbNullString) <> 0
        'Regmon is running...throw an access vio
        '     lation
        DoScan = True
        Exit Function
        Case FindWindow(sFileMonClass, vbNullString) <> 0
        'FileMon is running...throw an access vi
        '     olation
        DoScan = True
        Exit Function
    End Select
'\\So far so good...check for SoftICE in
'     memory
hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


If hFile <> -1 Then
    ' SoftICE is detected.
    retVal = CloseHandle(hFile) ' Close the file handle
    DoScan = True
    Exit Function
Else
    ' SoftICE is not found for windows 9x, c
    '     heck for NT.
    hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


    If hFile <> -1 Then
        ' SoftICE is detected.
        retVal = CloseHandle(hFile) ' Close the file handle
        DoScan = True
        Exit Function
    End If
End If

sScan = "f" & "i" & "l" & "e" & "W" & "A" & "T" & "C" & "H"


If FindWindow(vbNullString, sScan) <> 0 Then
    DoScan = True
    Exit Function
End If

sScan = "F" & "i" & "l" & "e" & "S" & "p" & "y"
If FindWindow(vbNullString, sScan) <> 0 Then
    DoScan = True
    Exit Function
End If
End Function
