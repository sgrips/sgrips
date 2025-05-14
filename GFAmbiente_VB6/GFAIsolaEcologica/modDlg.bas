Attribute VB_Name = "modDlg"
Option Explicit

Private Const cMaxPath = 1024
Private Const cMaxFile = 1024
Private Const sEmpty = ""

Public Const DT_Open = 1
Public Const DT_Save = 2
Public Const DT_Browse = 3

'Public Enum EErrorCommonDialog
'    eeBaseCommonDialog = 13450  ' CommonDialog
'End Enum

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    Flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" Alias "GetFileTitleA" (ByVal szFile As String, ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
'Public Const OFN_ENABLEHOOK = &H20
'Public Const OFN_ENABLETEMPLATE = &H40
'Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000

' Common dialog errors
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

'Public Enum EDialogError
Public Const CDERR_DIALOGFAILURE = &HFFFF
Public Const CDERR_GENERALCODES = &H0
Public Const CDERR_STRUCTSIZE = &H1
Public Const CDERR_INITIALIZATION = &H2
Public Const CDERR_NOTEMPLATE = &H3
Public Const CDERR_NOHINSTANCE = &H4
Public Const CDERR_LOADSTRFAILURE = &H5
Public Const CDERR_FINDRESFAILURE = &H6
Public Const CDERR_LOADRESFAILURE = &H7
Public Const CDERR_LOCKRESFAILURE = &H8
Public Const CDERR_MEMALLOCFAILURE = &H9
Public Const CDERR_MEMLOCKFAILURE = &HA
Public Const CDERR_NOHOOK = &HB
Public Const CDERR_REGISTERMSGFAIL = &HC
Public Const PDERR_PRINTERCODES = &H1000
Public Const PDERR_SETUPFAILURE = &H1001
Public Const PDERR_PARSEFAILURE = &H1002
Public Const PDERR_RETDEFFAILURE = &H1003
Public Const PDERR_LOADDRVFAILURE = &H1004
Public Const PDERR_GETDEVMODEFAIL = &H1005
Public Const PDERR_INITFAILURE = &H1006
Public Const PDERR_NODEVICES = &H1007
Public Const PDERR_NODEFAULTPRN = &H1008
Public Const PDERR_DNDMMISMATCH = &H1009
Public Const PDERR_CREATEICFAILURE = &H100A
Public Const PDERR_PRINTERNOTFOUND = &H100B
Public Const PDERR_DEFAULTDIFFERENT = &H100C

Public Const CFERR_CHOOSEFONTCODES = &H2000
Public Const CFERR_NOFONTS = &H2001
Public Const CFERR_MAXLESSTHANMIN = &H2002

Public Const FNERR_FILENAMECODES = &H3000
Public Const FNERR_SUBCLASSFAILURE = &H3001
Public Const FNERR_INVALIDFILENAME = &H3002
Public Const FNERR_BUFFERTOOSMALL = &H3003

Public Const CCERR_CHOOSECOLORCODES = &H5000

'Browse Directory
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer









Public Function DoCommonDialog(DialogType As Long, _
                               FileName As String, _
                      Optional FileTitle As String, _
                      Optional DlgTitle As String, _
                      Optional Owner As Long = -1, _
                      Optional Filter As String = "All (*.*)| *.*", _
                      Optional FilterIndex As Long = 1, _
                      Optional InitDir As String, _
                      Optional DefaultExt As String, _
                      Optional Flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String, b As Boolean, fileattr As Long
    Dim lpIDList As Long, tBrowseInfo As BrowseInfo

    If DialogType = DT_Browse Then
        If Owner <> -1 Then tBrowseInfo.hWndOwner = Owner
        tBrowseInfo.lpszTitle = lstrcat(DlgTitle, "")
        
        tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        lpIDList = SHBrowseForFolder(tBrowseInfo)

        b = lpIDList <> 0
        DoCommonDialog = b
        If b Then
           s = Space(cMaxPath)
           SHGetPathFromIDList lpIDList, s
           s = Left(s, InStr(s, vbNullChar) - 1)
           FileName = s
        End If
        Exit Function
    End If

    opfile.lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    opfile.Flags = Flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then opfile.hWndOwner = Owner
    ' InitDir can take initial directory string
    opfile.lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    opfile.lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    opfile.lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim i As Long
    ' Put double null at end
    s = Filter & vbNullChar & vbNullChar
    Do
        i = InStr(s, "|")
        If i = 0 Then i = InStr(s, ":")
        If i = 0 Then Exit Do
        Mid$(s, i, 1) = vbNullChar
    Loop
    opfile.lpstrFilter = s
    opfile.nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    
    On Error Resume Next
    fileattr = -1
    fileattr = GetAttr(FileName)
    On Error GoTo 0
    b = True
    If fileattr >= 0 Then
        If (fileattr And vbDirectory) > 0 Then
            s = FileName
            If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
            opfile.lpstrInitialDir = s
            opfile.lpstrFile = String$(cMaxPath, 0)
            opfile.lpstrFileTitle = String$(cMaxFile, 0)
            b = False
        End If
    End If
    If b Then
        s = FileName & String$(cMaxPath - Len(FileName), 0)
        opfile.lpstrFile = s
        s = FileTitle & String$(cMaxFile - Len(FileTitle), 0)
        opfile.lpstrFileTitle = s
    End If
    opfile.nMaxFile = cMaxPath
    opfile.nMaxFileTitle = cMaxFile
    ' All other fields set to zero
    
    Select Case DialogType
        Case DT_Open
            b = GetOpenFileName(opfile) <> 0
        Case DT_Save
            b = GetSaveFileName(opfile) <> 0
    End Select
        
    DoCommonDialog = b
    If b Then
        DoCommonDialog = True
        FileName = Left$(opfile.lpstrFile, InStr(opfile.lpstrFile, vbNullChar) - 1)
        FileTitle = Left$(opfile.lpstrFileTitle, InStr(opfile.lpstrFileTitle, vbNullChar) - 1)
        Flags = opfile.Flags
        ' Return the filter index
        FilterIndex = opfile.nFilterIndex
    End If
End Function


Public Sub TitleBar(frm As Form, ShowTitle As Boolean)
   Const GWW_ID = (-12)
   Const GWL_STYLE = (-16)
   Const WS_DLGFRAME = &H400000
   Const WS_SYSMENU = &H80000
   Const WS_MINIMIZEBOX = &H20000
   Const WS_MAXIMIZEBOX = &H10000


   Static Oldhmenu As Integer
   Static SavedStyle As Long
   Dim NewStyle As Long


   If ShowTitle Then
      'get the current style attributes
      NewStyle = GetWindowLong&(frm.hwnd, GWL_STYLE)
      'set only the attributes that were removed earlier
      NewStyle = NewStyle Or SavedStyle
      're-establish the menu
      If Oldhmenu <> 0 Then
         SetWindowWord frm.hwnd, GWW_ID, Oldhmenu
      End If
      'set the new style
      SetWindowLong frm.hwnd, GWL_STYLE, NewStyle
   Else
      'get the current style attributes
      NewStyle = GetWindowLong(frm.hwnd, GWL_STYLE)


      'determine whether the form has a dialog frame, a ControlBox,
      'a minimize button, or a maximize button and save this info.
      'for later use
      SavedStyle = 0
      SavedStyle = SavedStyle Or (NewStyle And WS_DLGFRAME)
      SavedStyle = SavedStyle Or (NewStyle And WS_SYSMENU)
      SavedStyle = SavedStyle Or (NewStyle And WS_MINIMIZEBOX)
      SavedStyle = SavedStyle Or (NewStyle And WS_MAXIMIZEBOX)


      'remove the attributes for a dialog frame, a ControlBox, a minimize
      'button and a maximize button
      NewStyle = NewStyle And Not WS_DLGFRAME
      NewStyle = NewStyle And Not WS_SYSMENU
      NewStyle = NewStyle And Not WS_MINIMIZEBOX
      NewStyle = NewStyle And Not WS_MAXIMIZEBOX


      'is there a menu associated with this form?
      Oldhmenu = GetWindowWord(frm.hwnd, GWW_ID)
      If Oldhmenu <> 0 Then
         'yes-zero the menu handle
         SetWindowWord frm.hwnd, GWW_ID, 0
      End If


      'set the new style
      SetWindowLong frm.hwnd, GWL_STYLE, NewStyle
   End If
End Sub


