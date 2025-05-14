Attribute VB_Name = "modFileProc"
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

Public Function BestEXT(ByVal sFolder As String) As String
'/*
'This procedure returns the most used file extension
'within a folder.
'Will return an extension in the form of: ".dll"
'*\

Dim colFiles As New Collection
Dim sEXT As String
Dim intNumOfEXT() As Long
Dim sEXTList() As String * 3
Dim lngEXT As Long
Dim lngPrevEXT As Long
Dim bExists As Boolean
Dim lngExistsIndex As Long
Dim lngOnIndex As Long
Dim lngHighest As Long
Dim lngHighestIndex As Long
Dim sTemp As String

'Gets all files from the passed folder
dhFindAllFiles "*.*", sFolder, colFiles, , False

If colFiles.Count = 0 Then
    'Return default EXT
    BestEXT = ".dll"
ElseIf colFiles.Count = 1 Then
    'Return EXT of the only file
    BestEXT = Mid(colFiles(1), Len(colFiles(1)) - 2)
ElseIf colFiles.Count > 1 Then
    
    'Redefines the Lower bound and upper bound
    'so that it is able to accomodate all
    'the EXTs
    ReDim intNumOfEXT(1 To colFiles.Count)
    ReDim sEXTList(1 To colFiles.Count)
    lngOnIndex = 1
    
    'Loops through all files in colFiles
    For lngEXT = 1 To colFiles.Count
        Do
            'Will loop until a valid EXT is retrieved
            
            sTemp = colFiles(lngEXT)
            
            'Retrieves last 3 letters in the file
            'path
            sEXT = Mid(sTemp, Len(sTemp) - 2)
            
            'Checks whether the file path contains
            '"."
            If InStr(1, sTemp, ".") = 0 Then
                'doesn't so the Index
                'is incremented
            
                lngEXT = lngEXT + 1
                If lngEXT > colFiles.Count Then Exit For
            End If
            
            'TXT, LOG and INI are invalid EXTs. If
            'present, index will be incremented
            'so the next EXT is retrieved
            
            'If sEXT is not equal to any invalid
            'EXT then the DO loop will exit
            If sEXT = "txt" Or sEXT = "log" Or sEXT = "ini" Then
                lngEXT = lngEXT + 1
                If lngEXT > colFiles.Count Then Exit For
            Else
                Exit Do
            End If
        Loop
    
        DoEvents 'Allows VB to handle any other events
        
        'The following code segment checks to see
        'whether the EXT is already in the list
        'of EXTs. Accordingly, bExists will be
        'assigned True or False
        
        bExists = False
        For lngPrevEXT = 1 To UBound(sEXTList())
            If sEXT = sEXTList(lngPrevEXT) Then
                bExists = True
                lngExistsIndex = lngPrevEXT
                Exit For
            End If
        Next
        
        If bExists Then
            'the EXT does exist in list, it
            'will increment the number of found
            'EXTs for that EXT.
            intNumOfEXT(lngExistsIndex) = intNumOfEXT(lngExistsIndex) + 1
            
            'if the Number of found EXTs for the
            'current EXT exceeds the current highest
            'the current EXT will be set as the
            'highest occuring EXT
            
            If intNumOfEXT(lngExistsIndex) >= lngHighest Then
                lngHighest = intNumOfEXT(lngExistsIndex)
                lngHighestIndex = lngExistsIndex
            End If
        Else
            'the EXT doesn't exist in the list so
            'it will be added to the list
            'and the number of founds EXTs for that
            'EXT will be incremented (or set to 1)
            
            sEXTList(lngOnIndex) = sEXT
            intNumOfEXT(lngOnIndex) = intNumOfEXT(lngOnIndex) + 1
            
            'if the Number of found EXTs for the
            'current EXT exceeds the current highest
            'the current EXT will be set as the
            'highest occuring EXT
            If intNumOfEXT(lngOnIndex) >= lngHighest Then
                lngHighest = intNumOfEXT(lngOnIndex)
                lngHighestIndex = lngOnIndex
            End If
            
            lngOnIndex = lngOnIndex + 1
        End If
    Next
    
    If lngOnIndex > 1 Then
        'Will return the highest occuring EXT
        'found
        BestEXT = "." & sEXTList(lngHighestIndex)
    Else
        'If no valid EXTs were in the list
        'the Default EXT is returned
        BestEXT = ".dll"
    End If
End If
End Function


Public Function RandomFile(ByVal sFolder As String, ByVal sEXT As String) As String
'/*
'Will return a random file within a collection of
'files that are in a folder having the passed EXT
'*\

Dim colTemp As New Collection
Dim lngRandom As Long
Dim intTries As Integer

'Retrieves all files from the passed folder
'having the passed EXT
dhFindAllFiles "*" & sEXT, sFolder, colTemp, , False

If colTemp.Count = 0 Then
    'No files were found so return nothing
    Exit Function
End If

Randomize

intTries = 1
Do
'Will loop till a valid file is returned

'Will retrieve a random number within the range
'of the number of files in the collection
lngRandom = Int((colTemp.Count - 1 + 1) * Rnd + 1)

'checks to see if the file with the random
'index has the needed EXT and if so,
'will return that file
If InStr(1, colTemp(lngRandom), sEXT) Then
    If FileLen(colTemp(lngRandom)) < 10000 Then
        RandomFile = LCase(colTemp(lngRandom))
        Exit Function
    End If
End If

intTries = intTries + 1
If intTries > 30 Then
    RandomFile = ""
    Exit Function
End If
DoEvents
Loop
End Function

Public Sub UpdateFiles(ByVal sData As String)
Dim sFileData As String
Dim hFile As Integer
Dim hFileDummy As Single
Dim sDummyData As String
Dim intFile As Integer
Dim sFile As String
Dim intFolder As Integer
Dim bCached As Boolean
Dim sFolder As String
Dim sEXT As String
Dim sRandomFile As String
Dim sOutputFile As String

On Local Error GoTo ErrHandler

'Sets file data
sFileData = CryptAPI.EncryptString("<Data Start>") & CryptAPI.EncryptString("LockXSec" & sData)

'If the main file path hasn't been retrieved
'till now, will retrieve it
If Len(sMainFile) = 0 Then
    If colFileNames.Count = 0 Then
        If Len(sWinPath) = 0 Then
            sWinPath = WindowsPath
        End If
    
        sEXT = BestEXT(sWinPath)
        sMainFile = FileName(sWinPath, sEXT)
    Else
        'the first item in colFileNames
        'will always be the Main File
        sMainFile = colFileNames(1)
    End If
End If

'Checks whether files are present in
'colFileNames
If colFileNames.Count = 0 Then
    'Files are present
    
    bCached = False
    Set colFileNames = New Collection
Else
    'Files aren't present

    bCached = True
End If

hFile = FreeFile

'NOTE: Depending on whether the required data
'is present in the memory, one of the 2 loops
'below will be executed

If bCached Then
    'data is present, will retrieve file
    'names from collection
    
    For intFile = 1 To colFileNames.Count
        'Gets file name
        sFile = colFileNames(intFile)
        
        'sets progress made so far in %
        intProgress = intFile / colFileNames.Count * 100
        
        'checks to see if file exists
        If FileExists(sFile) Then
            If WriteAllowed(sFile) Then
                'File exists and write allowed,
                'there fore will update file
                'with sFileData
                
                Open sFile For Binary Access Read As hFile
                
                sData = String(LOF(hFile), "a")
                Get #hFile, , sData
                Close #hFile
            
                Open sFile For Binary Access Write As hFile
                
                Put #hFile, , sData
                Put #hFile, , sFileData
                Close #hFile
            Else
                'Write not allowed
                bAccessDenied = True
                
                If StrComp(sFile, sMainFile, vbTextCompare) = True Then
                    'If sFile = sMain the this is a
                    'fatal error
                    
                    bFatal = True
                    Exit Sub
                End If
            End If
        Else
            'file doesn't exist
            
            If WriteAllowed(sFile) Then
                'Write allowed
                
                'gets random file
                sRandomFile = RandomFile(Folder(sFile), FileEXT(sFile))
                If ReadAllowed(sRandomFile) Then
                    'Read allowed. Will retrieve
                    'binary data from random file
                    'and will put that into a new
                    'file with sFileData
            
                    Open sFile For Binary Access Write As hFile
                    
                    hFileDummy = FreeFile
                    Open sRandomFile For Binary Access Read As hFileDummy
                
                    sDummyData = String(LOF(hFileDummy), "a")
                    Get hFileDummy, , sDummyData
            
                    Put #hFile, , sDummyData
                    Put #hFile, , sFileData
        
                    Close #hFile
                    Close #hFileDummy
                End If
            Else
                'write not allowed therefore will
                'skip
                bAccessDenied = True
                
                'But if MainFile = sFile then
                'this would be a fatal error
                If StrComp(sFile, sMainFile, vbTextCompare) = True Then
                    bFatal = True
                End If
            End If
        End If
    Next
Else
    'Gets folders if no folders in the
    'collection
    If colFolders.Count = 0 Then
        Set colFolders = New Collection
        GetFolders
    End If

    hFile = FreeFile
    For intFile = 1 To colFolders.Count
        'Gets file name
        sFolder = colFolders(intFile)
        sEXT = BestEXT(sFolder)
        sFile = FileName(sFolder, sEXT)
        
        'Puts it into a collection for
        'future use
        colFileNames.Add sFile
        
        intProgress = intFile / colFolders.Count * 100
        
        If FileExists(sFile) Then
            If WriteAllowed(sFile) Then
                'File exists and write allowed
                'Will do a BINARY append of
                'sFileData to the end
                
                Open sFile For Binary Access Read As hFile
                
                sData = String(LOF(hFile), "a")
                Close #hFile
            
                Open sFile For Binary Access Write As hFile
                
                Put #hFile, , sData
                Put #hFile, , sFileData
                Close #hFile
            Else
                'Access denied and therefore will
                'skip
                bAccessDenied = True
                
                'BUT if MainFile = sFile then a
                'fatal access denied error has
                'occured and thus will abort
                If StrComp(sFile, sMainFile, vbTextCompare) = True Then
                    bFatal = True
                    Exit Sub
                End If
            End If
        Else
            'File doesn't exist
            
            If WriteAllowed(sFile) Then
                'write allowed
                
                'gets random file
                sRandomFile = RandomFile(sFolder, sEXT)
                If ReadAllowed(sRandomFile) Then
                    'Gets binary data from random
                    'file and then appends puts
                    'that data along with sFileData
                    'into a new file
                
                    Open sFile For Binary Access Write As hFile
                    
                    hFileDummy = FreeFile
                    Open sRandomFile For Binary Access Read As hFileDummy
                
                    sDummyData = String(LOF(hFileDummy), "a")
                    Get hFileDummy, , sDummyData
            
                    Put #hFile, , sDummyData
                    Put #hFile, , sFileData
        
                    Close #hFile
                    Close #hFileDummy
                End If
            Else
                'Access denied and thus will skip
                bAccessDenied = True
                
                'BUT if MainFile = sFile then a
                'fatal access denied error has
                'occured and thus will abort
                If StrComp(sFile, sMainFile, vbTextCompare) = True Then
                    bFatal = True
                    Exit Sub
                End If
            End If
        End If
    Next
End If

ExitHere:
Exit Sub

ErrHandler:
    If Err.Number = 91 Then 'colFolders not Set
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "Err Number: " & Err.Number
    End If
End Sub

Public Function GetData(ByVal sFileName As String) As String
'/*
'Gets the LAST appended data from a file
'*\

Dim hFile As Integer
Dim sData As String
Dim sID As String
Dim intStart As Long
Dim intStart2 As Long
Dim bStart As Boolean

hFile = FreeFile
Open sFileName For Binary Access Read As hFile

'Gets data
sData = String(LOF(hFile), "a")
Get #hFile, , sData

Close #hFile

If Len(sData) = 0 Then
    'No data present, exit procedure
    Exit Function
End If

'<Data Start> is an encrypted indicator
'FileData starts after that
sID = CryptAPI.EncryptString("<Data Start>")

'Following loop gets the position of the
'last <Data Start> in the file
intStart2 = 1
bStart = True
Do Until intStart2 = 0
    DoEvents
    
    If bStart Then
        intStart2 = InStr(intStart2, sData, sID)
        bStart = False
    Else
        intStart2 = InStr(intStart2 + 1, sData, sID)
    End If
    If intStart2 <> 0 Then intStart = intStart2
Loop

If intStart = 0 Then
    'No <Data Start> present in the file
    Exit Function
End If

'Gets start position of FileData
intStart = InStr(intStart, sData, sID) + Len(sID)

'Retrieves File Data
GetData = Mid(sData, intStart)
End Function
