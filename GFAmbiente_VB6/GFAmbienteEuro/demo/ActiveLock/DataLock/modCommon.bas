Attribute VB_Name = "modCommon"
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

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'Variables that keep commonly used
'data in the memory
Public colFolders As Collection
Public colFileNames As Collection

'Some common variables
Public sWinPath As String
Public CryptAPI As New clsCryptAPI
Public sLocalKey As String
Public sMainFile As String

'Error variables
Public bTampered As Boolean
Public bAccessDenied As Boolean
Public bFatal As Boolean
Public intProgress As Integer

Public Sub GetFolders()

'/*
'Retrieves folders
'*\

Dim sFolder As String
Dim colTemp As New Collection
Dim intFolder As Integer

Set colFolders = New Collection

'gets window path
If Len(sWinPath) = 0 Then
    sWinPath = WindowsPath
End If

'Adds Window path
colFolders.Add sWinPath

'Gets folders inside the Windows folder
dhFindAllFiles "\", sWinPath & "", colTemp, , False, True
For intFolder = 1 To colTemp.Count
    'If the folder name contains "t", add the
    'folder
    If InStr(1, colTemp(intFolder), "t") <> 0 Then
        colFolders.Add colTemp(intFolder)
    End If
Next

'gets the program directory path
sFolder = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
If Len(sFolder) = 0 Then
    'For Windows NT
    sFolder = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
End If

'adds the program directory path
colFolders.Add sFolder

'Gets all folders within Program files
Set colTemp = New Collection
dhFindAllFiles "\", sFolder & "", colTemp, , False, True

For intFolder = 1 To colTemp.Count
    'If the folder name contains "a", add the
    'folder
    If InStr(1, colTemp(intFolder), "a") <> 0 Then
        colFolders.Add colTemp(intFolder)
    End If
Next

'gets the common files folder
sFolder = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\", "CommonFilesDir")
If Len(sFolder) = 0 Then
    'for windows NT
    sFolder = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\", "CommonFilesDir")
End If

'Adds common files folder
colFolders.Add sFolder

End Sub

Public Function TrimSpaces(ByVal strString As String) As String
'/*
'Removes all spaces from a string
'*\

    Dim lngpos As Long
    Do While InStr(1&, strString$, " ")
        DoEvents
        Let lngpos& = InStr(1&, strString$, " ")
        Let strString$ = Left$(strString$, (lngpos& - 1&)) & Right$(strString$, Len(strString$) - (lngpos& + Len(" ") - 1&))
    Loop
    Let TrimSpaces$ = strString$
End Function

Public Function scramb(ByVal strString As String) As String
'/*
'Scrambles a string
'*\

Dim i As Integer, even As String, odd As String
For i% = 1 To Len(strString$)
If i% Mod 2 = 0 Then
even$ = even$ & Mid(strString$, i%, 1)
Else
odd$ = odd$ & Mid(strString$, i%, 1)
End If
Next i
scramb$ = even$ & odd$
End Function

Public Function dhTrimNull(ByVal strValue As String) As String
'/*
'Removes the leading null in a string
'Useful for API calls
'*\

    Dim intPos As Integer
    
    intPos = InStr(strValue, vbNullChar)
    Select Case intPos
        Case 0
            dhTrimNull = strValue
        Case 1
            dhTrimNull = vbNullString
        Case Is > 1
            dhTrimNull = Left$(strValue, intPos - 1)
    End Select
End Function

Public Function unscramb(ByVal strString As String) As String
'/*
'Unscrambles a string
'*\

Dim X As Integer, evenint As Integer, oddint As Integer
Dim even As String, odd As String, fin As String
X% = Len(strString$)
X% = Int(Len(strString$) / 2) 'adding this returns the actual number like 1.5 instead of returning 2
'Form1.Caption = x
even$ = Mid(strString$, 1, X%)
odd$ = Mid(strString$, X% + 1)
For X = 1 To Len(strString$)
If X% Mod 2 = 0 Then
evenint% = evenint% + 1
fin$ = fin$ & Mid(even$, evenint%, 1)
Else
oddint% = oddint% + 1
fin$ = fin$ & Mid(odd$, oddint%, 1)
End If
Next X%
unscramb$ = fin$
End Function


Public Function WindowsPath() As String
'/*
'Gets windows path
'*\

Dim sPath As String

sPath = Space(255)
GetWindowsDirectory sPath, 255
WindowsPath = dhTrimNull(sPath)

If Right(WindowsPath, 1) <> "\" Then
    WindowsPath = WindowsPath & "\"
End If
End Function

Public Function ReadAllowed(ByVal sFileName As String) As Boolean
'/*
'Checks whether Windows has allowed reading
'a file by attempting a read on that file
'*\

Dim hFile As Integer

On Error GoTo errFile

hFile = FreeFile
If FileExists(sFileName) Then
    Open sFileName For Binary Access Read As hFile
    
    Close #hFile
    ReadAllowed = True
End If

errFile:
End Function

Public Function WriteAllowed(ByVal sFileName As String) As Boolean
'/*
'Checks whether Windows has allowed writing
'to a file by attempting a write on that file
'*\


Dim hFile As Integer

On Error GoTo errFile

hFile = FreeFile
If FileExists(sFileName) Then
    Open sFileName For Append As hFile
    
    Close #hFile
    WriteAllowed = True
Else
    Open sFileName For Output As hFile
    Close #hFile
    
    WriteAllowed = True
    Kill sFileName
End If

errFile:
End Function

Public Function FileExists(ByVal strFile As String) As Boolean
'*\
'checks if a file exists
'/*

    On Local Error Resume Next
    FileExists = (Len(Dir(strFile)) > 0)
End Function

Public Function HDSerial() As String
'*\
'gets the harddrive serial number
'/*

Dim volbuf$, sysname$, sysflags&, componentlength&
Dim serialnum As Long

GetVolumeInformation "C:\", volbuf$, 255, serialnum, componentlength, sysflags, sysname$, 255
HDSerial = CStr(serialnum)
End Function


Public Sub SetKey(ByVal sString As String)
'*\
'sets sLocalKey
'/*

sLocalKey = MD5AA1Hash(sString & HDSerial)

CryptAPI.Key = sLocalKey
End Sub

Public Function FileName(ByVal sFolder As String, ByVal sEXT As String) As String
'*\
'generates a file name for the passed
'folder and extension
'/*

Dim sLocal As String
Dim intLen As Byte
Dim intASC As Byte

sLocal = MD5AA1Hash(sFolder & sEXT & sLocalKey)

intASC = Asc(Left$(sLocal, 1))
intLen = Val(Right$(sLocal, 1))
If intLen = 0 Then intLen = 4

If Right$(sFolder, 1) <> "\" Then
    sFolder = sFolder & "\"
End If

FileName = LCase(sFolder & Right$(sLocal, intLen) & sEXT)
End Function

Public Function IsTampered(ByVal sFileData As String) As Boolean
'*\
'Checks whether FileData has been tampered with.
'The first 8 characters have to be "LockXSec"
'/*

If Len(sFileData) = 0 Then
    IsTampered = True
    Exit Function
End If
If Mid(CryptAPI.DecryptString(sFileData), 1, 8) <> "LockXSec" Then IsTampered = True
End Function

Public Function Folder(ByVal sFile As String) As String
'*\
'retrieves the folder from a full file path
'/*

Dim intCon As Integer
Dim intLen As Integer

intCon = 1
Do Until intCon = 0
    intCon = InStr(intCon, sFile, "\")
    If intCon <> 0 Then
        intLen = intCon
        intCon = intCon + 1
    End If
Loop

Folder = Left$(sFile, intLen)
End Function

Public Function FileEXT(ByVal sFile As String) As String
'*\
'retrieves the file extension from a file name/path
'/*

FileEXT = Mid(sFile, Len(sFile) - 2)
End Function

Public Sub FullClean()
'*\
'Loops through all the files and attempts deleting
'them
'/*


Dim intFolder As Integer
Dim intFile As Integer
Dim sFile As String
Dim sFolder As String
Dim bCached As Boolean

On Error GoTo ErrHandler

If colFileNames.Count = 0 Then
    'Files are not present
    bCached = False
    Set colFileNames = New Collection
Else
    'File are present and therefore will
    'read from colFileNames instead
    bCached = True
End If

If bCached Then
    For intFile = 1 To colFileNames.Count
        sFile = colFileNames(intFile)
        
        intProgress = intFile / colFileNames.Count * 100
        
        If FileExists(sFile) Then
            If WriteAllowed(sFile) Then
                Kill sFile
            End If
        End If
        DoEvents
    Next
Else
    If colFolders.Count = 0 Then
        GetFolders
    End If
    
    For intFolder = 1 To colFolders.Count
        sFolder = colFolders(intFolder)
        sFile = FileName(sFolder, BestEXT(sFolder))
        
        
        intProgress = intFolder / colFolders.Count * 100
        colFileNames.Add sFile
        
        If FileExists(sFile) Then
            If WriteAllowed(sFile) Then
                Kill sFile
            End If
        End If
        DoEvents
    Next
End If

ExitHere:
Exit Sub

ErrHandler:
    If Err.Number = 91 Then 'Object not set
        Resume Next
    Else
        'For debug purposes
        MsgBox Err.Description, vbCritical, "Err Number: " & Err.Number
    End If
End Sub
