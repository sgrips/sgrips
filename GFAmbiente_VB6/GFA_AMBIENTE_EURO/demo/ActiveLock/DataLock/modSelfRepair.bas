Attribute VB_Name = "modSelfRepair"
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

Public Sub QuickScan()
'/*
'This procedure checks whether the main file has
'been corrupted or deleted. If it is, it will loop
'through all other backup files and update the main
'copy, if they too are not corrupted and are
'valid files
'*\


Dim sEXT As String
Dim sFileName As String
Dim sFolder As String

Dim sData As String
Dim bCached As Boolean

Dim intFolder As Integer
Dim intFile As Integer

On Error GoTo ErrHandler

'First checks whether the Main File variable
'has the path of the main file. If not, fills it in
If Len(sMainFile) = 0 Then
    If colFileNames.Count = 0 Then
        
        'Gets Windows Path
        If Len(sWinPath) = 0 Then
            sWinPath = WindowsPath
        End If
    
        'Gets the most used extension in the
        'windows folder
        sEXT = BestEXT(sWinPath)
        
        'Gets a file name for the Windows folder
        sMainFile = FileName(sWinPath, sEXT)
    Else
        'The first file in colFileNames will always
        'be the MainFile
        sMainFile = colFileNames(1)
    End If
End If

'Checks whether MainFile Exists and whether tampered with
'If not tampered or deleted, the procedure will Exit

If FileExists(sMainFile) Then
    If ReadAllowed(sMainFile) Then
        sData = GetData(sMainFile)
        If IsTampered(sData) Then
            If WriteAllowed(sMainFile) Then
                Kill sMainFile
            End If
        Else
            'File OK! Exit Procedure
            Exit Sub
        End If
    Else
        'Fatal Error. Since Read Not allowed
        'on the main file. Nothing can be done
        
        bFatal = True
        bAccessDenied = True
        Exit Sub
    End If
End If

'Checks whether filenames are present within
'colFileNames or not
If colFileNames.Count = 0 Then
    'Files are not present
    bCached = False
    Set colFileNames = New Collection
Else
    'File are present
    bCached = True
End If


'NOTE: Depending on whether the filenames are stored
'in colFileNames, one of the 2 loops below will
'be executed.

If Not bCached Then
    'If no folders are present in colFolders
    'then will GetFolders
    If colFolders.Count = 0 Then
        GetFolders
    End If
    
    'Note: Starts from 2 because Index 1
    'will always be the MainFile
    For intFolder = 2 To colFolders.Count
        DoEvents
    
        sFolder = colFolders(intFolder)
        
        intProgress = intFolder / colFolders.Count * 100
        
        'retrieves FileName for the passed folder
        sFileName = FileName(sFolder, BestEXT(sFolder))
        
        'Fills colFileNames for future use
        colFileNames.Add sFileName

        'The following code segment will check
        'whether the file exists and if tampered or
        'not. If not tampered and if exists, it
        'will replace the MainFile
        
        If FileExists(sFileName) Then
            'Checks whether Windows will allow
            'to read file
            If ReadAllowed(sFileName) Then
                
                'Gets file data
                sData = GetData(sFileName)
                
                'Checks if data is tampered
                If IsTampered(sData) Then
                    bTampered = True
                Else
                    'Checks whether Writing is allowed
                    'by windows on the main file
                    If Not WriteAllowed(sMainFile) Then
                        bAccessDenied = True
                        bFatal = True
                        
                        'Since writing to the MainFile isn't
                        'allowed, a fatal Error has occured
                        'and thus with exit the procedure
                        Exit Sub
                    End If
                    
                    'MainFile being replaced
                    CopyFile sFileName, sMainFile
                    
                    'Since replace is successful, the following
                    'variables will be set to false
                    bTampered = False
                    bAccessDenied = False
                    
                    'MainFile is replaced and will Exit sub
                    Exit Sub
                End If
            Else
                'Since read not allowed, will
                'skip this file
                bAccessDenied = True
            End If
        End If
    Next
Else

    'Note: Starts from 2 because Index 1
    'will always be the MainFile
    For intFile = 2 To colFileNames.Count
        DoEvents
    
        intProgress = intFile / colFileNames.Count * 100
        
        'Retrieves filename from collection
        sFileName = colFileNames(intFile)


        'The following code segment will check
        'whether the file exists and if tampered or
        'not. If not tampered and if exists, it
        'will replace the MainFile
        
        If FileExists(sFileName) Then
            'Checks whether Windows will allow
            'to read file
            If ReadAllowed(sFileName) Then
                
                'Gets file data
                sData = GetData(sFileName)
                
                'Checks if data is tampered
                If IsTampered(sData) Then
                    bTampered = True
                Else
                    'Checks whether Writing is allowed
                    'by windows on the main file
                    If Not WriteAllowed(sMainFile) Then
                        bAccessDenied = True
                        bFatal = True
                        
                        'Since writing to the MainFile isn't
                        'allowed, a fatal Error has occured
                        'and thus will exit the procedure
                        Exit Sub
                    End If
                    
                    'MainFile being replaced
                    CopyFile sFileName, sMainFile
                    
                    'Since replace is successful, the following
                    'variables will be set to false
                    bTampered = False
                    bAccessDenied = False
                    
                    'MainFile is replaced and will Exit sub
                    Exit Sub
                End If
            Else
                'Since read not allowed, will
                'skip this file
                bAccessDenied = True
            End If
        End If
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

Public Sub FullScan()
'/*
'This procedure first calls QuickScan to make sure
'the MainFile is OK. Then it will loop through all
'the backup files. IF anyone is missing or tampered
'with, it will be replaced with a fresh copy
'of the Main File
'*\


Dim intFile As Integer
Dim intFolder As Integer
Dim sFile As String
Dim sData As String
Dim bCached As Boolean
Dim sFolder As String

On Error GoTo ErrHandler

'Calls QuickScan to first check whether the MainFile
'is OK or not
QuickScan

If Not FileExists(sMainFile) Then Exit Sub

If bTampered Or bAccessDenied Or bFatal Then
    'This means that the MainFile is still
    'tampered with or doesn't exist and thus
    'cannot continue
    Exit Sub
End If

'Checks whether filenames are present within
'colFileNames or not
If colFileNames.Count = 0 Then
    'Files are not present
    bCached = False
    Set colFileNames = New Collection
Else
    'File are present
    bCached = True
End If

'NOTE: Depending on whether Data is present in
'colFileNames, one of the 2 loops beneath will
'be executed

If bCached Then
    'Note: Starts from Index 2 because Index 1
    'will always be the path to the MainFile
    For intFile = 2 To colFileNames.Count
        DoEvents
        
        sFile = colFileNames(intFile)
        
        intProgress = intFile / colFileNames.Count * 100
        If FileExists(sFile) Then
            'Checks whether Windows allows
            'Reading this file
            If ReadAllowed(sFile) Then
                'Gets data from file
                sData = GetData(sFile)
                
                'Checks whether data has been
                'tampered with
                If IsTampered(sData) Then
                    If WriteAllowed(sFile) Then
                        'data has been tampered
                        'and write allowed. it
                        'will be replaced by a
                        'good copy of the Main file
                        
                        CopyFile sMainFile, sFile
                    End If
                End If
            End If
        Else
            'if write allowed by Windows, a copy
            'of MainFile will be written to
            'sFile
            
            If WriteAllowed(sFile) Then
                CopyFile sMainFile, sFile
            End If
        End If
    Next
Else

    'Checks whether folders are present in
    'colFolders and if not, will get them
    If colFolders.Count = 0 Then
        GetFolders
    End If
    
    'Note: Starts from Index 2 because Index 1
    'will always be the path to the MainFile
    For intFolder = 2 To colFolders.Count
        DoEvents
        
        sFolder = colFolders(intFolder)
        
        intProgress = intFolder / colFolders.Count * 100
        
        'Retrieves filename for passed folder
        sFile = FileName(sFolder, BestEXT(sFolder))
        colFileNames.Add sFile
        
        If FileExists(sFile) Then
            'Checks whether read is allowed by
            'Windows
            If ReadAllowed(sFile) Then
                'Gets File data
                sData = GetData(sFile)
                
                'checks if file data is tampered
                If IsTampered(sData) Then
                    If WriteAllowed(sFile) Then
                        'tampered and write
                        'allowed. sFile is replaced
                        'with a copy of the Main file
                        CopyFile sMainFile, sFile
                    End If
                End If
            End If
        Else
            'If file not found and write allowed,
            'sFile will be replaced with a copy
            'of the Main File
            
            If WriteAllowed(sFile) Then
                CopyFile sMainFile, sFile
            End If
        End If
    Next
End If
    
ExitHere:
Exit Sub

ErrHandler:
    If Err.Number = 91 Then 'Object not set
        Resume Next
    Else
        MsgBox Err.Description, vbCritical, "Err Number: " & Err.Number
    End If
End Sub


Private Sub CopyFile(sFrom As String, sTo As String)
'/*
'This procedure will copy the data from one to
'another.
'*\

Dim sData As String
Dim sFolder As String
Dim sEXT As String
Dim sRandomFile As String
Dim sBinaryData As String
Dim hFile As Integer

'Gets FileData
sData = CryptAPI.EncryptString("<Data Start>") & GetData(sFrom)

'Gets random file for a folder
sFolder = Folder(sTo)
sEXT = FileEXT(sTo)
sRandomFile = RandomFile(sFolder, sEXT)

hFile = FreeFile
If ReadAllowed(sRandomFile) Then
    'Gets binary data from random file

    Open sRandomFile For Binary Access Read As hFile
    
    sBinaryData = String(LOF(hFile), "a")
    Get #hFile, , sBinaryData
    
    Close #hFile
Else
    'Nothing can be done
    Exit Sub
End If

'Puts the binary data along with the file data
'into the new file
Open sTo For Binary Access Write As hFile

Put #hFile, , sBinaryData
Put #hFile, , sData

Close #hFile
End Sub


