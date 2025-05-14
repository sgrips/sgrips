Attribute VB_Name = "modFileAPI"
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

' Functions for searching for files in a given directory
Private Declare Function FindFirstFile Lib "kernel32" _
 Alias "FindFirstFileA" (ByVal lpFileName As String, _
 lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindNextFile Lib "kernel32" _
 Alias "FindNextFileA" (ByVal hFindFile As Long, _
 lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindClose Lib "kernel32" _
 (ByVal hFindFile As Long) As Long
 
Private Declare Function SearchPath Lib "kernel32" _
 Alias "SearchPathA" (ByVal lpPath As String, _
 ByVal lpFileName As String, ByVal lpExtension As String, _
 ByVal nBufferLength As Long, ByVal lpBuffer As String, _
 ByVal lpFilePart As String) As Long
 
 Private Const MAX_PATH = 260

Private Type FILETIME
    lngLowDateTime As Long
    lngHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    lngFileAttributes As Long           ' File attributes
    ftCreationTime As FILETIME          ' Creation time
    ftLastAccessTime As FILETIME        ' Last access time
    ftLastWriteTime As FILETIME         ' Last modified time
    lngFileSizeHigh As Long             ' Size (high word)
    lngFileSizeLow As Long              ' Size (low word)
    lngReserved0 As Long                ' reserved
    lngReserved1 As Long                ' reserved
    strFilename As String * MAX_PATH    ' File name
    strAlternate As String * 14         ' 8.3 name
End Type


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Variant
    bInheritHandle As Boolean
End Type

 
Private Function dhDir(Optional ByVal strPath As String = "", _
 Optional lngAttributes As VbFileAttribute = vbNormal, _
 Optional fExclusive As Boolean = True) As String

    ' Replacement for the Dir function which lets you
    ' specify file attributes for a restrictive search.

    ' In:
    '   strPath (Optional, default = "")
    '       Path and/or file specification to search.
    '   lngAttributes (Optional, default = vbNormal)
    '       File attributes.
    '   fExclusive (Optional, default = True)
    '       If True, only those files with the matching
    '       file attributes are returned.
    ' Out:
    '   Return Value:
    '       If called with a file specification, the first
    '       matching filename is returned. If called without
    '       a file specification, the next matching filename
    '       is returned. When no additional matching filenames
    '       are found, an empty string is returned.
    ' Example:
    '   Dim strDir As String
    '
    '   strDir = dhDir("C:\", vbDirectory)
    '   Do Until strDir = ""
    '       Debug.Print strDir
    '       strDir = dhDir()
    '   Loop
      
    Dim fd As WIN32_FIND_DATA
    Static hFind As Long
    Static lngAttr As Long
    Static fEx As Boolean
    Dim strOut As String
    
    ' If no path was passed, try to find the next file
    If strPath = "" Then
        If hFind > 0 Then
            If CBool(FindNextFile(hFind, fd)) Then
                strOut = dhFindByAttr(hFind, fd, lngAttr, fEx)
            End If
        Else
            Err.Raise 5 ' Invalid procedure call or argument
        End If
        
    ' Otherwise, start a new search
    Else
        ' Store the attributes and exclusive settings
        lngAttr = lngAttributes
        fEx = fExclusive
        
        ' If the path ends in a backslash, assume
        ' all files and append "*.*"
        If Right(strPath, 1) = "\" Then
            strPath = strPath & "*.*"
        End If
        
        ' Find the first file
        hFind = FindFirstFile(strPath, fd)
        If hFind > 0 Then
            strOut = dhFindByAttr(hFind, fd, lngAttr, fEx)
        End If
    End If
    
    ' If the search failed, close the Find handle.
    If Len(strOut) = 0 Then
        If hFind > 0 Then
            Call FindClose(hFind)
        End If
    End If

    dhDir = strOut
End Function

Private Function dhFindByAttr(hFind As Long, _
 fd As WIN32_FIND_DATA, lngAttr As VbFileAttribute, _
 fExclusive As Boolean) As String

    ' Determines if a file matches the specified attrbites.

    ' In:
    '   hFind
    '       Windows API Find handle.
    '   fd
    '       Pointer to populated WIN32_FIND_DATA structure.
    '   lngAttr
    '       Attributes to test for.
    '   fExclusive
    '       If True, only those files with the matching
    '       file attributes are returned.
    ' Out:
    '   Return Value:
    '       Next matching filename.
    ' Example:
    '   See dhDir for usage
 
    Dim fOk As Boolean
 
    ' Continue looking for files until one
    ' matches the given attributes exactly
    ' (if fExclusive is True) or just contains
    ' them (if fExclusive is False)
    Do
        If fExclusive Then
            fOk = fd.lngFileAttributes = lngAttr
        Else
            fOk = (fd.lngFileAttributes And lngAttr) = lngAttr
        End If
            
        If fOk Then
            dhFindByAttr = dhTrimNull(fd.strFilename)
            Exit Do
        End If
    Loop While FindNextFile(hFind, fd)
End Function


Public Function dhFindAllFiles(strSpec As String, _
 ByVal strPath As String, colFound As Collection, _
 Optional lngAttr As Long = -1, _
 Optional fRecursive As Boolean = True, Optional OnlyFolders As Boolean = False) As Long


    ' In:
    '   strSpec
    '       File specification to search for.
    '   strPath
    '       Starting path.
    '   colFound
    '       Pointer to VBA Collection object.
    '   lngAttr (Optional, default = -1)
    '       File attributes to search for (-1 for all files).
    '   fRecursive (Optional, default = True)
    '       If True, a recursive search of all subdirectories is made.
    '   objCallback (Optional)
    '       Optional pointer to a callback object (see chapter
    '       text for details).
    ' Out:
    '   colFound
    '       Contains one element for each matching filename,
    '       including complete path).
    '   Return Value:
    '       Number of files found.
    ' Example:
    '   See dhPrintFoundFiles for example

    Dim strFile As String
    Dim colSubDir As New Collection
    Dim varDir As Variant
    
    On Error GoTo HandleError
    
    ' Make sure strPath ends in a backslash
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    

        
    ' Find all files in the directory--if no
    ' attributes were specified use a non-exclusive
    ' search for all files, otherwise use a
    ' restrictive search for the attributes
    If lngAttr = -1 Then
        strFile = dhDir(strPath & strSpec, , False)
    Else
        strFile = dhDir(strPath & strSpec, lngAttr)
    End If
     
    Do Until strFile = ""
    
        ' Add file to collection if attributes match
        ' (special case directories "." and ".."
        If ((GetAttr(strPath & strFile) And lngAttr) > 0) _
         And (strFile <> ".") And (strFile <> "..") Then
            If OnlyFolders Then
                If InStr(1, strFile, ".") = 0 Then
                    colFound.Add LCase(strPath & strFile)
                End If
            Else
                colFound.Add LCase(strPath & strFile)
            End If
        End If
        

        
        ' Get the next file
        strFile = dhDir
    Loop
    
    ' If the recursive flag is set build a list
    ' of all the subdirectories
    If fRecursive Then
        
        strFile = dhDir(strPath, vbDirectory, False)
        Do Until strFile = ""
            ' Ignore "." and ".."
            If strFile <> "." And strFile <> ".." Then
            
                ' Add each to the directory collection
                colSubDir.Add strPath & strFile
            End If
            strFile = dhDir
        Loop
        
        ' Now recurse through each sub directory
        For Each varDir In colSubDir
            dhFindAllFiles strSpec, varDir, colFound, _
             lngAttr, fRecursive
        Next
    End If
    
    ' Return the number of found files
    dhFindAllFiles = colFound.Count
ExitHere:
    Exit Function
HandleError:
    Select Case Err.Number
        Case 75 ' Path/file access error
            Resume Next
        Case Else
            MsgBox Err.Description, vbExclamation, "Error " & _
             Err.Number & " in dhFindAllFiles"
            Resume ExitHere
    End Select
End Function

'Public Function FileExists(sSource As String) As Boolean
'
'   Dim WFD As WIN32_FIND_DATA
'   Dim hFile As Long
'
'   hFile = FindFirstFile(sSource, WFD)
'   FileExists = hFile <> INVALID_HANDLE_VALUE
'
'   Call FindClose(hFile)
'
'End Function
