VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private bOpened As Boolean
Private sMainFile As String

Public Sub OpenSlot(sAppName As String, sPassword As String, sSlot As String)
Dim sWinPath As String

SetKey sAppName & sPassword & sSlot

sWinPath = WindowsPath
sMainFile = FileName(sWinPath, BestEXT(sWinPath))
End Sub
Public Property Get FatalError() As Boolean
FatalError = bFatal
End Property

Public Sub SetData(sData As String)
intProgress = 0

If DoScan = False Then
    UpdateFiles scramb(sData)
Else
    bFatal = True
End If
End Sub

Public Function RetrieveData() As String
Dim sData As String

intProgress = 0

If DoScan = False Then
    'for better security, change below to fullScan
    QuickScan
    If Not bFatal Then
        sData = GetData(sMainFile)
        sData = CryptAPI.DecryptString(sData)
        sData = Mid(sData, 9)
        RetrieveData = unscramb(sData)
    End If
Else
    bFatal = True
End If
End Function

Public Sub DoClean()
intProgress = 0
FullClean
End Sub


Public Property Get Progress() As Integer
'Sorry, couldn't be implemented as an event
Progress = intProgress
End Property
