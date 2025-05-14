VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMDIBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrExeName          As String
Private mfrmMDI              As MDIForm
Private mblnRunningIDE       As Boolean
Private mlnghModule          As Long
Private mlngForegroundBmpID  As Long
Private mlngBackgroundBmpID  As Long
Private mlngTransparentColor As Long

'
' API Stuff.
'
Private Const API_FALSE As Long = 0&
Private Const API_TRUE As Long = 1&
Private Const GW_HWNDNEXT As Long = 2&
Private Const GW_CHILD As Long = 5&
Private Const GWL_STYLE As Long = (-16&)
Private Const RASTERCAPS As Long = 38&
Private Const SIZEPALETTE As Long = 104&
Private Const RC_PALETTE As Long = &H100&
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000
Private Const SM_CXVSCROLL As Long = 2&
Private Const SM_CYHSCROLL As Long = 3&
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCPAINT As Long = &HEE0086
Private Const SRCAND As Long = &H8800C6
Private Const NOTSRCCOPY As Long = &H330008

Private Type RECT
    left   As Long
    top    As Long
    right  As Long
    bottom As Long
End Type

Private Type POINTL
    x As Long
    y As Long
End Type

Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC_BMP
    Size     As Long
    Type     As Long
    hBmp     As Long
    hPal     As Long
    Reserved As Long
End Type



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth&, ByVal nHeight&, ByVal nPlanes&, ByVal nBitCount&, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC&, ByVal nWidth&, ByVal nHeight&) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE256) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC&, ByVal crColor&) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC&, ByVal hObject&) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd&) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd&, ByVal hDC&) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC&, ByVal nSavedDC&) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject&) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC&, ByVal iCapabilitiy&) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC&, ByVal wStartIndex&, ByVal wNumEntries&, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC_BMP, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC&, ByVal hPalette&, ByVal bForceBackground&) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd&, lpRect As RECT) As Long
Private Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance&, ByVal lpBitmapName&) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd&, ByVal wCmd&) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName$) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule&) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex&) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd&, lpRect As RECT, ByVal bErase&) As Long


Private Function fCenterForegroundOnBackground(ByVal hWndTarget As Long, _
        ByVal lngForgroundResID As Long, _
        ByVal lngBackgroundResID As Long, _
        ByVal lngTransparentColor As Long, _
        ByRef stdPicture As stdPicture, _
        ByRef strError As String) As Boolean

Dim hOwnerDC           As Long
Dim hForgroundDC       As Long
Dim hBackgroundDC      As Long
Dim hStretchDC         As Long
Dim hBmpObj            As Long
Dim nClientWidth       As Long
Dim nClientHeight      As Long
Dim hBackgroundBmp     As Long
Dim hPrevBackgroundBmp As Long
Dim hForegroundBmp     As Long
Dim hPrevForegroundBmp As Long
Dim hPrevStretchBmp    As Long
Dim hShrinkDC          As Long
Dim hPrevBitmap        As Long
Dim nShrinkWidth       As Long
Dim nShrinkHeight      As Long
Dim rectStretch        As RECT
Dim rectForeground     As RECT
Dim ptForegroundSize   As POINTL
Dim ptBkgroundSize     As POINTL
Dim Bmp                As BITMAP
    
    
    On Error GoTo ErrorHandler
    fCenterForegroundOnBackground = False
    strError = ""
    
    '
    ' Create the DCs we will use.
    '
    hOwnerDC = GetDC(hWndTarget)
    hForgroundDC = CreateCompatibleDC(hOwnerDC)
    hBackgroundDC = CreateCompatibleDC(hOwnerDC)
    hStretchDC = CreateCompatibleDC(hOwnerDC)
    Call ReleaseDC(hWndTarget, hOwnerDC)
    Call SaveDC(hForgroundDC)
    Call SaveDC(hBackgroundDC)
    Call SaveDC(hStretchDC)
   

    '
    ' Load the foreground bitmap.
    '
    hForegroundBmp = LoadBitmapBynum(mlnghModule, lngForgroundResID)
    
    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hForegroundBmp, Len(Bmp), Bmp) Then
        With ptForegroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        strError = "Unable to retrieve foreground bitmap info."
        GoTo ErrorHandler
    End If

    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevForegroundBmp = SelectObject(hForgroundDC, hForegroundBmp)
    
    '
    ' Load the background bitmap.
    '
    hBackgroundBmp = LoadBitmapBynum(mlnghModule, lngBackgroundResID)

    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hBackgroundBmp, Len(Bmp), Bmp) Then
        With ptBkgroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        strError = "Unable to retrieve background bitmap info."
        GoTo ErrorHandler
    End If
    
    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevBackgroundBmp = SelectObject(hBackgroundDC, hBackgroundBmp)
  
       
    '
    ' Get the rect of the window.
    '
    Call GetClientRect(hWndTarget, rectStretch)
  

    With rectStretch
        '
        ' If either scrollbar is showing, add the width of the scrollbar to the client area dimentions.
        '
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_VSCROLL Then .right = .right + GetSystemMetrics(SM_CXVSCROLL)
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_HSCROLL Then .bottom = .bottom + GetSystemMetrics(SM_CYHSCROLL)
        
        '
        ' Store the client dimentions for later use.
        '
        nClientWidth = .right
        nClientHeight = .bottom
        
        '
        ' Create a new bitmap and select it into the stretch DC.
        '
        hPrevStretchBmp = SelectObject(hStretchDC, CreateCompatibleBitmap(hBackgroundDC, .right, .bottom))
    
        '
        ' Stretch the background bitmap onto the stretched DC.
        '
        Call StretchBlt(hStretchDC, 0&, 0&, .right, .bottom, hBackgroundDC, 0&, 0&, ptBkgroundSize.x, ptBkgroundSize.y, SRCCOPY)
    
        '
        ' If the overlay graphic doesn't fit in the client area......
        '
        If (.right < (ptForegroundSize.x + 20)) Or (.bottom < (ptForegroundSize.y + 20)) Then
            '
            ' Determine which dimention is the smallest and adjust the values accordingly.
            '
            Select Case True
                Case ((ptForegroundSize.x + 20) - .right) > ((ptForegroundSize.y + 20) - .bottom)
                    nShrinkWidth = .right - 20
                    nShrinkHeight = nShrinkWidth * (ptForegroundSize.y / ptForegroundSize.x)
                
                Case Else
                    nShrinkHeight = .bottom - 20
                    nShrinkWidth = nShrinkHeight * (ptForegroundSize.x / ptForegroundSize.y)
            End Select

            hShrinkDC = CreateCompatibleDC(hForgroundDC)
            SaveDC hShrinkDC
            
            '
            ' Create a new bitmap and select it into the shrink DC.
            '
            hPrevBitmap = SelectObject(hShrinkDC, CreateCompatibleBitmap(hForgroundDC, nShrinkWidth, nShrinkHeight))

            '
            ' Shrink the bitmap onto the shrink DC.
            '
            Call StretchBlt(hShrinkDC, 0&, 0&, nShrinkWidth, nShrinkHeight, hForgroundDC, 0&, 0&, ptForegroundSize.x, ptForegroundSize.y, SRCCOPY)

            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = nShrinkWidth
                .bottom = nShrinkHeight
            End With

            '
            ' Calculate the location for the foreground bitmap.
            '
            .left = (.right \ 2) - (nShrinkWidth \ 2)
            .right = .left + nShrinkWidth
            .top = (.bottom \ 2) - (nShrinkHeight \ 2)
            .bottom = .top + nShrinkHeight

            '
            ' Copy the graphic onto the background.
            '
            pTransparentBlt hStretchDC, rectStretch, hShrinkDC, rectForeground, lngTransparentColor

            '
            ' Clean up the shrink DC.
            '
            Call DeleteObject(SelectObject(hShrinkDC, hPrevBitmap))
            Call RestoreDC(hShrinkDC, (-1&))
            Call DeleteDC(hShrinkDC)
        Else
            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = ptForegroundSize.x
                .bottom = ptForegroundSize.y
            End With
            
            '
            ' Calculate the location for the foreground bitmap.
            '
            .left = (.right \ 2) - (ptForegroundSize.x \ 2)
            .right = .left + ptForegroundSize.x
            .top = (.bottom \ 2) - (ptForegroundSize.y \ 2)
            .bottom = .top + ptForegroundSize.y
            
            '
            ' Copy the graphic onto the background.
            '
            pTransparentBlt hStretchDC, rectStretch, hForgroundDC, rectForeground, lngTransparentColor
        End If
    End With
    
    '
    ' Create a picture object and return it to the caller.
    '
    Set stdPicture = fPictureFromDC(hStretchDC, 0, 0, nClientWidth, nClientHeight)
    
    
    fCenterForegroundOnBackground = True
    GoTo NormalExit
    
ErrorHandler:
    If strError = "" Then strError = Err.Description

NormalExit:
    On Error Resume Next
    '
    ' Delete the memory DCs used for holding the bitmaps.
    '
    Call DeleteObject(SelectObject(hForgroundDC, hPrevForegroundBmp))
    Call RestoreDC(hForgroundDC, (-1&))
    Call DeleteDC(hForgroundDC)
    
    Call DeleteObject(SelectObject(hBackgroundDC, hPrevBackgroundBmp))
    Call RestoreDC(hBackgroundDC, (-1&))
    Call DeleteDC(hBackgroundDC)
    
    Call DeleteObject(SelectObject(hStretchDC, hPrevStretchBmp))
    Call RestoreDC(hStretchDC, (-1&))
    Call DeleteDC(hStretchDC)
End Function
Public Function fCenterForeground(ByVal hWndTarget As Long, _
        ByVal nForgroundResID As Long, _
        ByVal nTransparentColor As Long, _
        ByRef stdPicture As stdPicture, _
        ByRef strError As String) As Boolean
  
Dim hOwnerDC           As Long
Dim hForgroundDC       As Long
Dim hBackgroundDC      As Long
Dim hStretchDC         As Long
Dim hBmpObj            As Long
Dim nClientWidth       As Long
Dim nClientHeight      As Long
Dim hBackgroundBmp     As Long
Dim hPrevBackgroundBmp As Long
Dim hForegroundBmp     As Long
Dim hPrevForegroundBmp As Long
Dim hPrevStretchBmp    As Long
Dim hShrinkDC          As Long
Dim hPrevBitmap        As Long
Dim nShrinkWidth       As Long
Dim nShrinkHeight      As Long
Dim blnUseBackground   As Boolean
Dim rectStretch        As RECT
Dim rectForeground     As RECT
Dim ptForegroundSize   As POINTL
Dim ptBkgroundSize     As POINTL
Dim Bmp                As BITMAP

    
    On Error GoTo ErrorHandler
    fCenterForeground = False
    strError = ""
    
    
    '
    ' Create the DCs we will use.
    '
    hOwnerDC = GetDC(hWndTarget)
    hForgroundDC = CreateCompatibleDC(hOwnerDC)
    hBackgroundDC = CreateCompatibleDC(hOwnerDC)
    hStretchDC = CreateCompatibleDC(hOwnerDC)
    Call SaveDC(hForgroundDC)
    Call SaveDC(hBackgroundDC)
    Call SaveDC(hStretchDC)
    
    
    '
    ' Load the foreground bitmap.
    '
    hForegroundBmp = LoadBitmapBynum(mlnghModule, nForgroundResID)
    
    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hForegroundBmp, Len(Bmp), Bmp) Then
        With ptForegroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        strError = "Unable to retrieve bitmap information."
        GoTo ErrorHandler
    End If
    
    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevForegroundBmp = SelectObject(hForgroundDC, hForegroundBmp)
   
    
    '
    ' Get the rect of the window.
    '
    Call GetClientRect(hWndTarget, rectStretch)
    
    With rectStretch
        '
        ' If either scrollbar is showing, add the width of the scrollbar to the client area dimentions.
        '
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_VSCROLL Then .right = .right + GetSystemMetrics(SM_CXVSCROLL)
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_HSCROLL Then .bottom = .bottom + GetSystemMetrics(SM_CYHSCROLL)
        
        '
        ' Store the client dimentions for later use.
        '
        nClientWidth = .right
        nClientHeight = .bottom
        
        '
        ' Create a new bitmap and select it into the stretch DC.
        '
        hPrevBackgroundBmp = SelectObject(hBackgroundDC, CreateCompatibleBitmap(hOwnerDC, .right, .bottom))
    
        '
        ' Copy the existing background bitmap onto the stretched DC.
        '
        Call BitBlt(hBackgroundDC, 0, 0, .right, .bottom, hOwnerDC, .left, .top, SRCCOPY)
        
        '
        ' Release the owner DC.
        '
        Call ReleaseDC(hWndTarget, hOwnerDC)
        
        '
        ' If the overlay graphic doesn't fit in the client area......
        '
        If (.right < (ptForegroundSize.x + 20)) Or (.bottom < (ptForegroundSize.y + 20)) Then
            '
            ' Determine which dimention is the smallest and adjust the values accordingly.....
            '
            Select Case True
                Case ((ptForegroundSize.x + 20) - .right) > ((ptForegroundSize.y + 20) - .bottom)
                    nShrinkWidth = .right - 20
                    nShrinkHeight = nShrinkWidth * (ptForegroundSize.y / ptForegroundSize.x)
                
                Case Else
                    nShrinkHeight = .bottom - 20
                    nShrinkWidth = nShrinkHeight * (ptForegroundSize.x / ptForegroundSize.y)
            End Select
    
            '
            ' Create a DC to shrink the foreground bitmap into.
            '
            hShrinkDC = CreateCompatibleDC(hForgroundDC)
            Call SaveDC(hShrinkDC)
            
            '
            ' Create a new bitmap and select it into the shrink DC.
            '
            hPrevBitmap = SelectObject(hShrinkDC, CreateCompatibleBitmap(hForgroundDC, nShrinkWidth, nShrinkHeight))
            
            '
            ' Shrink the bitmap onto the shrink DC.
            '
            Call StretchBlt(hShrinkDC, 0&, 0&, nShrinkWidth, nShrinkHeight, hForgroundDC, 0&, 0&, ptForegroundSize.x, ptForegroundSize.y, SRCCOPY)
            
            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = nShrinkWidth
                .bottom = nShrinkHeight
            End With
            
            '
            ' Calculate the location for the foreground bitmap.
            '
            .left = (.right \ 2) - (nShrinkWidth \ 2)
            .right = .left + nShrinkWidth
            .top = (.bottom \ 2) - (nShrinkHeight \ 2)
            .bottom = .top + nShrinkHeight
            
            '
            ' Copy the graphic onto the background using the transparent color.
            '
            pTransparentBlt hBackgroundDC, rectStretch, hShrinkDC, rectForeground, nTransparentColor
    
            '
            ' Clean up the shrink DC
            '
            Call DeleteObject(SelectObject(hShrinkDC, hPrevBitmap))
            Call RestoreDC(hShrinkDC, (-1&))
            Call DeleteDC(hShrinkDC)
        Else
            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = ptForegroundSize.x
                .bottom = ptForegroundSize.y
            End With
            
            '
            ' Calculate the location for the foreground bitmap.
            '
            .left = (.right \ 2) - (ptForegroundSize.x \ 2)
            .right = .left + ptForegroundSize.x
            .top = (.bottom \ 2) - (ptForegroundSize.y \ 2)
            .bottom = .top + ptForegroundSize.y
            
            '
            ' Copy the graphic onto the background using the transparent color.
            '
            pTransparentBlt hBackgroundDC, rectStretch, hForgroundDC, rectForeground, nTransparentColor
        End If
    End With
    
    '
    ' Create a picture object and return it to the caller.
    '
    Set stdPicture = fPictureFromDC(hBackgroundDC, 0, 0, nClientWidth, nClientHeight)

    fCenterForeground = True
    GoTo NormalExit
    
ErrorHandler:
    If strError = "" Then strError = Err.Description

NormalExit:
    On Error Resume Next
    '
    ' Delete the memory DCs used for holding the bitmaps.
    '
    Call DeleteObject(SelectObject(hForgroundDC, hPrevForegroundBmp))
    Call RestoreDC(hForgroundDC, (-1&))
    Call DeleteDC(hForgroundDC)
    
    Call DeleteObject(SelectObject(hBackgroundDC, hPrevBackgroundBmp))
    Call RestoreDC(hBackgroundDC, (-1&))
    Call DeleteDC(hBackgroundDC)
    
    Call DeleteObject(SelectObject(hStretchDC, hPrevStretchBmp))
    Call RestoreDC(hStretchDC, (-1&))
    Call DeleteDC(hStretchDC)
End Function
Private Function fTileForeground(ByVal hWndTarget As Long, _
        ByVal lngForgroundResID As Long, _
        ByVal lngTransparentColor As Long, _
        ByRef stdPicture As stdPicture, _
        ByRef strError As String) As Boolean


Dim hOwnerDC           As Long
Dim hForgroundDC       As Long
Dim hBackgroundDC      As Long
Dim hBmpObj            As Long
Dim nClientWidth       As Long
Dim nClientHeight      As Long
Dim rectStretch        As RECT
Dim hBackgroundBmp     As Long
Dim hPrevBackgroundBmp As Long
Dim hForegroundBmp     As Long
Dim hPrevForegroundBmp As Long
Dim hPrevStretchBmp    As Long
Dim hShrinkDC          As Long
Dim hPrevBitmap        As Long
Dim nShrinkWidth       As Long
Dim nShrinkHeight      As Long
Dim ptForegroundSize   As POINTL
Dim ptBkgroundSize     As POINTL
Dim Bmp                As BITMAP
Dim rectTile           As RECT
Dim rectForeground     As RECT

    On Error GoTo ErrorHandler
    fTileForeground = False
    strError = ""
    
    
    '
    ' Create the DCs we will use.
    '
    hOwnerDC = GetDC(hWndTarget)
    hForgroundDC = CreateCompatibleDC(hOwnerDC)
    hBackgroundDC = CreateCompatibleDC(hOwnerDC)
    Call SaveDC(hForgroundDC)
    Call SaveDC(hBackgroundDC)
    
    '
    ' Load the foreground bitmap.
    '
    hForegroundBmp = LoadBitmapBynum(mlnghModule, lngForgroundResID)
    
    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hForegroundBmp, Len(Bmp), Bmp) Then
        With ptForegroundSize
        .x = Bmp.bmWidth
        .y = Bmp.bmHeight
        End With
    Else
        strError = "Unable to retrieve foreground bitmap info."
        GoTo ErrorHandler
    End If
    
    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevForegroundBmp = SelectObject(hForgroundDC, hForegroundBmp)
      
    '
    ' Get the rect of the window.
    '
    Call GetClientRect(hWndTarget, rectStretch)
  

    With rectStretch
        '
        ' If either scrollbar is showing, add the width of the scrollbar to the client area dimentions.
        '
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_VSCROLL Then .right = .right + GetSystemMetrics(SM_CXVSCROLL)
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_HSCROLL Then .bottom = .bottom + GetSystemMetrics(SM_CYHSCROLL)
        
        '
        ' Store the client dimentions for later use.
        '
        nClientWidth = .right
        nClientHeight = .bottom
        
        '
        ' Clear the existing picture.
        '
        Set mfrmMDI.Picture = Nothing
    
        '
        ' Create a new bitmap and select it into the stretch DC.
        '
        hPrevBackgroundBmp = SelectObject(hBackgroundDC, CreateCompatibleBitmap(hOwnerDC, .right, .bottom))
        
        '
        ' Copy the existing background bitmap onto the stretched DC.
        '
        Call BitBlt(hBackgroundDC, 0, 0, .right, .bottom, hOwnerDC, .left, .top, SRCCOPY)
        '
        ' Release the owner DC.
        '
        Call ReleaseDC(hWndTarget, hOwnerDC)
        
        '
        ' If the overlay graphic doesn't fit in the client area.
        '
        If (.right < (ptForegroundSize.x + 20)) Or (.bottom < (ptForegroundSize.y + 20)) Then
            '
            ' Determine which dimention is the smallest and adjust the values accordingly.
            '
            Select Case True
                Case ((ptForegroundSize.x + 20) - .right) > ((ptForegroundSize.y + 20) - .bottom)
                    nShrinkWidth = .right - 20
                    nShrinkHeight = nShrinkWidth * (ptForegroundSize.y / ptForegroundSize.x)
                
                Case Else
                    nShrinkHeight = .bottom - 20
                    nShrinkWidth = nShrinkHeight * (ptForegroundSize.x / ptForegroundSize.y)
            End Select
            
            '
            ' Create a DC to shrink the foreground bitmap into.
            '
            hShrinkDC = CreateCompatibleDC(hForgroundDC)
            Call SaveDC(hShrinkDC)
            
            '
            ' create a new bitmap and select it into the shrink DC.
            '
            hPrevBitmap = SelectObject(hShrinkDC, CreateCompatibleBitmap(hForgroundDC, nShrinkWidth, nShrinkHeight))
            
            '
            ' shrink the bitmap onto the shrink DC.
            '
            Call StretchBlt(hShrinkDC, 0&, 0&, nShrinkWidth, nShrinkHeight, hForgroundDC, 0&, 0&, ptForegroundSize.x, ptForegroundSize.y, SRCCOPY)
            
            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = nShrinkWidth
                .bottom = nShrinkHeight
            End With
    
            '
            ' Calculate the location for the foreground bitmap.
            '
            .left = (.right \ 2) - (nShrinkWidth \ 2)
            .right = .left + nShrinkWidth
            .top = (.bottom \ 2) - (nShrinkHeight \ 2)
            .bottom = .top + nShrinkHeight
    
            '
            ' Copy the graphic onto the background using the transparent color.
            '
            Call pTransparentBlt(hBackgroundDC, rectStretch, hShrinkDC, rectForeground, lngTransparentColor)
            
            '
            ' Clean up the shrink DC.
            '
            Call DeleteObject(SelectObject(hShrinkDC, hPrevBitmap))
            Call RestoreDC(hShrinkDC, (-1&))
            Call DeleteDC(hShrinkDC)
        Else
            '
            ' Set the dimentions of the graphic.
            '
            With rectForeground
                .right = ptForegroundSize.x
                .bottom = ptForegroundSize.y
            End With
        
            '
            ' Copy the graphic onto the background.
            '
            pTransparentBlt hBackgroundDC, rectForeground, hForgroundDC, rectForeground, lngTransparentColor
    
            '
            ' Tile the images onto the new bitmap we create above.
            '
            With rectTile
                Do Until .top >= nClientHeight
                    .bottom = .top + ptForegroundSize.y
                
                    Do Until .left >= nClientWidth
                        .right = .left + ptForegroundSize.x
                        Call BitBlt(hBackgroundDC, .left, .top, .right, .bottom, hBackgroundDC, 0&, 0&, SRCCOPY)
                        .left = .right
                    Loop
                
                .left = 0
                .top = .bottom
                Loop
            End With
        End If
    End With
  
    '
    ' Create a picture object and return it to the caller.
    '
    Set stdPicture = fPictureFromDC(hBackgroundDC, 0, 0, nClientWidth, nClientHeight)

    fTileForeground = True
    GoTo NormalExit
    
ErrorHandler:
    If strError = "" Then strError = Err.Description

NormalExit:
    On Error Resume Next
    '
    ' Delete the memory DCs used for holding the bitmaps.
    '
    Call DeleteObject(SelectObject(hForgroundDC, hPrevForegroundBmp))
    Call RestoreDC(hForgroundDC, (-1&))
    Call DeleteDC(hForgroundDC)
    
    Call DeleteObject(SelectObject(hBackgroundDC, hPrevBackgroundBmp))
    Call RestoreDC(hBackgroundDC, (-1&))
    Call DeleteDC(hBackgroundDC)
End Function
Public Function CreateFormPicture(ByRef Frm As MDIForm, _
        ByVal lngAction As Long, _
        ByRef strError As String) As Boolean

Dim lnghClientArea As Long
Dim rc             As RECT
Dim stdPicture     As stdPicture

    On Error GoTo ErrorHandler
    CreateFormPicture = False
    strError = ""


    '
    ' Get the resource from the exe when running in the IDE.
    ' this is because the resource is not compiled into the module pointed to
    ' by App.hInstance when this code is run in the IDE.
    '
    If mblnRunningIDE Then
        mlnghModule = LoadLibrary(mstrExeName)
    Else
        mlnghModule = App.hInstance
    End If
    If mlnghModule = 0 Then
        strError = "Unable to find bitmap resources."
        GoTo ErrorHandler
    End If

    '
    ' In MDI parent forms the client area is a seperate child window.
    ' Get the hWnd of this window since VB does not expose it.
    '
    lnghClientArea = FindWindowEx(Frm.hWnd, 0&, "MDIClient", vbNullChar)
    
    '
    ' Invalidate the client rect for the image to be redrawn
    ' properly on some systems. First, get the rect to invalidate.
    '
    Call GetClientRect(lnghClientArea, rc)
    
    '
    ' Invalidate the client rect so it will be redrawn.
    '
    Call InvalidateRect(lnghClientArea, rc, API_TRUE)
    
    
    Select Case lngAction
        Case 0 'Background only.
            '
            ' Stretch the background bitmap to fit the form's client area.
            ' Pass the resource ID of the background bitmap.
            '
            If Not fStretchToBackground(lnghClientArea, mlngBackgroundBmpID, stdPicture, strError) Then
                GoTo ErrorHandler
            End If
        
        Case 1 'Foreground only.
            '
            ' Center the foreground bitmap on the default background.
            '
            ' Pass the resource ID of the foreground bitmap and the
            ' transparent color used for drawing the bitmap onto the
            ' background.
            '
            mfrmMDI.Picture = Nothing
            If Not fCenterForeground(lnghClientArea, mlngForegroundBmpID, _
                        mlngTransparentColor, stdPicture, strError) Then
                GoTo ErrorHandler
            End If
        
        Case 2 'Foreground and Background.
            '
            ' Center the foreground bitmap on the background bitmap.
            ' Pass the resource IDs of the foreground and background
            ' bitmaps and the color we want to be transparent when the
            ' foreground bitmap is drawn onto the background.
            '
            If Not fCenterForegroundOnBackground(lnghClientArea, mlngForegroundBmpID, _
                        mlngBackgroundBmpID, mlngTransparentColor, stdPicture, strError) Then
                GoTo ErrorHandler
            End If
        
        Case 3 'Tiled foreground.
            '
            ' Ttile the foreground bitmap.
            ' Pass both resource IDs and the transparent color.
            '
            If Not fTileForeground(lnghClientArea, mlngForegroundBmpID, _
                        mlngTransparentColor, stdPicture, strError) Then
                GoTo ErrorHandler
            End If
        
        Case 4 'Tiled foreground on background.
            '
            ' Tile the foreground bitmap onto the background.
            ' Pass both resource IDs and the transparent color.
            '
            If Not fTileForegroundOnBackground(lnghClientArea, mlngForegroundBmpID, _
                        mlngBackgroundBmpID, mlngTransparentColor, stdPicture, strError) Then
                GoTo ErrorHandler
            End If
    End Select
    
    mfrmMDI.Picture = stdPicture
    
    '
    ' Invalidate the client rect again for the redraw to occur.
    '
    Call InvalidateRect(lnghClientArea, rc, API_TRUE)

    CreateFormPicture = True
    GoTo NormalExit

ErrorHandler:

NormalExit:
    '
    ' Free the library if we are running in the IDE.
    '
    If mblnRunningIDE Then Call FreeLibrary(mlnghModule)

End Function
Private Function fPictureFromBitmap(ByVal hBmp As Long, ByVal hPal As Long) As stdPicture
Dim IPictureIID As GUID
Dim IPic        As IPicture
Dim tagPic      As PICTDESC_BMP
Dim lpGUID      As Long

    '
    ' Fill in the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    '
    With IPictureIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    '
    ' Set the properties on the picture object.
    '
    With tagPic
        .Size = Len(tagPic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    '
    ' Create a picture that will delete it's
    ' bitmap when it is finished with it.
    '
    Call OleCreatePictureIndirect(tagPic, IPictureIID, API_TRUE, IPic)

    '
    ' Return the picture to the caller.
    '
    Set fPictureFromBitmap = IPic
    
End Function
Private Sub pTransparentBlt(hDestDC As Long, lpDestRect As RECT, hSrcDC As Long, lpSrcRect As RECT, ByVal TransColor As Long)
Dim hInvDC         As Long
Dim hMaskDC        As Long
Dim hResultDC      As Long
Dim hInvBmp        As Long
Dim hMaskBmp       As Long
Dim hResultBmp     As Long
Dim hInvPrevBmp    As Long
Dim hMaskPrevBmp   As Long
Dim hDestPrevBmp   As Long
Dim nSrcWidth      As Long
Dim nSrcHeight     As Long
Dim nOriginalColor As Long

    With lpSrcRect
        nSrcWidth = .right - .left
        nSrcHeight = .bottom - .top
    End With
    
    '
    ' Create the mask and invert stage DCs and bitmaps.
    '
    hInvDC = CreateCompatibleDC(hDestDC)
    hMaskDC = CreateCompatibleDC(hDestDC)
    
    '
    ' Monochrome bitmaps for the masks.
    '
    hInvBmp = CreateBitmap(nSrcWidth, nSrcHeight, 1, 1, ByVal 0&)
    hMaskBmp = CreateBitmap(nSrcWidth, nSrcHeight, 1, 1, ByVal 0&)
    hInvPrevBmp = SelectObject(hInvDC, hInvBmp)
    hMaskPrevBmp = SelectObject(hMaskDC, hMaskBmp)
    
    '
    ' Create the DC and bitmap to hold the result.
    '
    hResultDC = CreateCompatibleDC(hDestDC)
    
    '
    ' Color bitmap for final result.
    '
    hResultBmp = CreateCompatibleBitmap(hDestDC, nSrcWidth, nSrcHeight)
    hDestPrevBmp = SelectObject(hResultDC, hResultBmp)
    
    '
    ' Create mask: set background color of source to transparent color.
    '
    nOriginalColor = SetBkColor(hSrcDC, TransColor)
    With lpSrcRect
        Call BitBlt(hMaskDC, 0, 0, nSrcWidth, nSrcHeight, hSrcDC, .left, .top, SRCCOPY)
    End With
    TransColor = SetBkColor(hSrcDC, nOriginalColor)
    
    '
    ' Create inverse of mask to AND w/ source & combine w/ background.
    '
    Call BitBlt(hInvDC, 0, 0, nSrcWidth, nSrcHeight, hMaskDC, 0, 0, NOTSRCCOPY)
    
    '
    ' Copy background bitmap to result & create final transparent bitmap.
    '
    With lpDestRect
        Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hDestDC, .left, .top, SRCCOPY)
        
        '
        ' AND mask bitmap w/ result DC to punch hole in the background by
        ' painting black area for non-transparent portion of source bitmap.
        '
        Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hMaskDC, 0, 0, SRCAND)
        
        '
        ' AND inverse mask w/ source bitmap to turn off bits associated
        ' with transparent area of source bitmap by making it black.
        '
        Call BitBlt(hSrcDC, 0, 0, nSrcWidth, nSrcHeight, hInvDC, 0, 0, SRCAND)
        
        '
        ' XOR result w/ source bitmap to make background show through.
        '
        Call BitBlt(hResultDC, 0, 0, nSrcWidth, nSrcHeight, hSrcDC, 0, 0, SRCPAINT)
        
        '
        ' Copy result to the dest DC.
        '
        Call BitBlt(hDestDC, .left, .top, nSrcWidth, nSrcHeight, hResultDC, 0, 0, SRCCOPY)
    End With
    
    '
    ' Clean up.
    '
    Call DeleteObject(SelectObject(hMaskDC, hMaskPrevBmp))
    Call DeleteObject(SelectObject(hInvDC, hInvPrevBmp))
    Call DeleteObject(SelectObject(hResultDC, hDestPrevBmp))
    
    Call DeleteDC(hMaskDC)
    Call DeleteDC(hInvDC)
    Call DeleteDC(hResultDC)
End Sub
Private Function fPictureFromDC(ByVal hDCSrc As Long, ByVal nLeft As Long, ByVal nTop As Long, ByVal nWidth As Long, ByVal nHeight As Long) As stdPicture
Dim hDCMemory       As Long
Dim hBmp            As Long
Dim hBmpPrev        As Long
Dim hPal            As Long
Dim hPalPrev        As Long
Dim fHasPalette     As Long
Dim nPaletteEntries As Long
Dim LogPal          As LOGPALETTE256

    '
    ' Create the DC and bitmap we will use.
    '
    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, nWidth, nHeight)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    '
    ' Determine whether or not the video supports 256 color palettes.
    '
    nPaletteEntries = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    fHasPalette = GetDeviceCaps(hDCSrc, RASTERCAPS) And RC_PALETTE

    '
    ' If this is 256 color video, we need to create and
    ' add a palette to our DC.
    '
    If fHasPalette And (nPaletteEntries = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        Call GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        Call RealizePalette(hDCMemory)
    End If

    '
    ' Copy the passed image into the local DC.
    '
    Call BitBlt(hDCMemory, 0, 0, nWidth, nHeight, hDCSrc, nLeft, nTop, vbSrcCopy)

    '
    ' Get the bitmap from the DC.
    '
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    '
    ' If we created a palette, get it from the DC.
    '
    If fHasPalette And (nPaletteEntries = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    '
    ' Clean up.
    '
    Call DeleteDC(hDCMemory)

    ' create a picture from the bitmap and return it to the caller
    Set fPictureFromDC = fPictureFromBitmap(hBmp, hPal)
    
End Function
Private Function fStretchToBackground(ByVal hWndTarget As Long, _
        ByVal nBackgroundResID As Long, _
        ByRef stdPicture As stdPicture, _
        ByRef strError As String) As Boolean

Dim hOwnerDC           As Long
Dim hBackgroundDC      As Long
Dim hStretchDC         As Long
Dim hBackgroundBmp     As Long
Dim hPrevBackgroundBmp As Long
Dim hPrevStretchBmp    As Long
Dim hBmpObj            As Long
Dim ptBkgroundSize     As POINTL
Dim Bmp                As BITMAP
Dim rectStretch        As RECT

    On Error GoTo ErrorHandler
    fStretchToBackground = False
    strError = ""

    '
    ' Stretch the background image onto the client area.
    '
    '
    ' Create the DCs.
    '
    hOwnerDC = GetDC(hWndTarget)
    hBackgroundDC = CreateCompatibleDC(hOwnerDC)
    hStretchDC = CreateCompatibleDC(hOwnerDC)
    Call ReleaseDC(hWndTarget, hOwnerDC)
    Call SaveDC(hBackgroundDC)
    Call SaveDC(hStretchDC)
    

    '
    ' Load the background bitmap.
    '
    hBackgroundBmp = LoadBitmapBynum(mlnghModule, nBackgroundResID)

    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hBackgroundBmp, Len(Bmp), Bmp) Then
        With ptBkgroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        Call MsgBox("Unable to retrieve bitmap info.", vbExclamation, "Cannot Format Background")
    End If

    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevBackgroundBmp = SelectObject(hBackgroundDC, hBackgroundBmp)
    
    '
    ' Get the rect of the window.
    '
    Call GetClientRect(hWndTarget, rectStretch)


    With rectStretch
        '
        ' If either scrollbar is showing, add the width of the scrollbar
        ' to the client area dimentions.
        '
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_VSCROLL Then .right = .right + GetSystemMetrics(SM_CXVSCROLL)
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_HSCROLL Then .bottom = .bottom + GetSystemMetrics(SM_CYHSCROLL)
        
        '
        ' Create a new bitmap and select it into the stretch DC.
        '
        hPrevStretchBmp = SelectObject(hStretchDC, CreateCompatibleBitmap(hBackgroundDC, .right, .bottom))
        
        '
        ' Stretch the background bitmap onto the stretched DC.
        '
        Call StretchBlt(hStretchDC, 0&, 0&, .right, .bottom, hBackgroundDC, 0&, 0&, ptBkgroundSize.x, ptBkgroundSize.y, SRCCOPY)
        
        '
        ' Create a picture object and return it to the caller.
        '
        Set stdPicture = fPictureFromDC(hStretchDC, 0, 0, .right, .bottom)
    End With
    
    fStretchToBackground = True
    GoTo NormalExit
    
ErrorHandler:
    If strError = "" Then strError = Err.Description

NormalExit:
    On Error Resume Next
    '
    ' Clean up.
    '
    Call DeleteObject(SelectObject(hBackgroundDC, hPrevBackgroundBmp))
    Call RestoreDC(hBackgroundDC, (-1&))
    Call DeleteDC(hBackgroundDC)
    
    Call DeleteObject(SelectObject(hStretchDC, hPrevStretchBmp))
    Call RestoreDC(hStretchDC, (-1&))
    Call DeleteDC(hStretchDC)
End Function
Private Function fTileForegroundOnBackground(ByVal hWndTarget As Long, _
        ByVal lngForgroundResID As Long, _
        ByVal lngBackgroundResID As Long, _
        ByVal lngTransparentColor As Long, _
        ByRef stdPicture As stdPicture, _
        ByRef strError As String) As Boolean
  
Dim hOwnerDC           As Long
Dim hForgroundDC       As Long
Dim hBackgroundDC      As Long
Dim hStretchDC         As Long
Dim hBmpObj            As Long
Dim nClientWidth       As Long
Dim nClientHeight      As Long
Dim hBackgroundBmp     As Long
Dim hPrevBackgroundBmp As Long
Dim hForegroundBmp     As Long
Dim hPrevForegroundBmp As Long
Dim hPrevStretchBmp    As Long
Dim Bmp                As BITMAP
Dim rectStretch        As RECT
Dim rectForeground     As RECT
Dim rectTile           As RECT
Dim ptForegroundSize   As POINTL
Dim ptBkgroundSize     As POINTL
  
    On Error GoTo ErrorHandler
    fTileForegroundOnBackground = False
    strError = ""
  
    
    '
    ' Create the DCs we will use.
    '
    hOwnerDC = GetDC(hWndTarget)
    hForgroundDC = CreateCompatibleDC(hOwnerDC)
    hBackgroundDC = CreateCompatibleDC(hOwnerDC)
    hStretchDC = CreateCompatibleDC(hOwnerDC)
    Call ReleaseDC(hWndTarget, hOwnerDC)
    Call SaveDC(hForgroundDC)
    Call SaveDC(hBackgroundDC)
    Call SaveDC(hStretchDC)


    '
    ' Load the foreground bitmap.
    '
    hForegroundBmp = LoadBitmapBynum(mlnghModule, lngForgroundResID)
    
    '
    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hForegroundBmp, Len(Bmp), Bmp) Then
        With ptForegroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        Call MsgBox("Unable to retrieve foreground bitmap info.", vbExclamation, "Cannot Format Background")
    End If
  
    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevForegroundBmp = SelectObject(hForgroundDC, hForegroundBmp)
    
    '
    ' Load the background bitmap.
    '
    hBackgroundBmp = LoadBitmapBynum(mlnghModule, lngBackgroundResID)

    ' Get the bitmap info so that we can get the size of the bitmap.
    '
    If GetObjectAPI(hBackgroundBmp, Len(Bmp), Bmp) Then
        With ptBkgroundSize
            .x = Bmp.bmWidth
            .y = Bmp.bmHeight
        End With
    Else
        Call MsgBox("Unable to retrieve background bitmap info.", vbExclamation, "Cannot Format Background")
    End If

    '
    ' Save the previous bitmap when we select the new one into the DC.
    '
    hPrevBackgroundBmp = SelectObject(hBackgroundDC, hBackgroundBmp)
  
        
    '
    ' Get the rect of the window.
    '
    Call GetClientRect(hWndTarget, rectStretch)
  

    With rectStretch
        '
        ' If either scrollbar is showing, add the width of the scrollbar to the client area dimentions.
        '
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_VSCROLL Then .right = .right + GetSystemMetrics(SM_CXVSCROLL)
        If GetWindowLong(hWndTarget, GWL_STYLE) And WS_HSCROLL Then .bottom = .bottom + GetSystemMetrics(SM_CYHSCROLL)
        
        '
        ' Store the client dimentions for later use.
        '
        nClientWidth = .right
        nClientHeight = .bottom
        
        '
        ' Create a new bitmap and select it into the stretch DC.
        '
        hPrevStretchBmp = SelectObject(hStretchDC, CreateCompatibleBitmap(hBackgroundDC, .right, .bottom))
        
        '
        ' Stretch the background bitmap onto the stretched DC.
        '
        Call StretchBlt(hStretchDC, 0&, 0&, .right, .bottom, hBackgroundDC, 0&, 0&, ptBkgroundSize.x, ptBkgroundSize.y, SRCCOPY)
        
        '
        ' Delete the background bitmap.
        '
        DeleteObject SelectObject(hBackgroundDC, hPrevBackgroundBmp)
        
        '
        ' Create a new bitmap to hold the tiled images.
        '
        hPrevBackgroundBmp = SelectObject(hBackgroundDC, CreateCompatibleBitmap(hStretchDC, .right, .bottom))

        '
        ' Set the dimentions of the graphic.
        '
        With rectForeground
            .right = ptForegroundSize.x
            .bottom = ptForegroundSize.y
        End With
        
        '
        ' Tile the images onto the new bitmap we create above.
        '
        With rectTile
            Do Until .top >= nClientHeight
                .bottom = .top + ptForegroundSize.y
                
                Do Until .left >= nClientWidth
                    .right = .left + ptForegroundSize.x
                    Call BitBlt(hBackgroundDC, .left, .top, .right, .bottom, hForgroundDC, 0&, 0&, SRCCOPY)
                    .left = .right
                Loop
        
                .left = 0
                .top = .bottom
            Loop
        End With
    End With
    
    '
    ' Copy the tiled bitmap onto the background image.
    '
    Call pTransparentBlt(hStretchDC, rectStretch, hBackgroundDC, rectStretch, lngTransparentColor)
    
    '
    ' Create a picture object and return it to the caller.
    '
    Set stdPicture = fPictureFromDC(hStretchDC, 0, 0, nClientWidth, nClientHeight)
    
    fTileForegroundOnBackground = True
    GoTo NormalExit
    
ErrorHandler:
    If strError = "" Then strError = Err.Description

NormalExit:
    On Error Resume Next
    '
    ' Delete the memory DCs used for holding the bitmaps.
    '
    Call DeleteObject(SelectObject(hForgroundDC, hPrevForegroundBmp))
    Call RestoreDC(hForgroundDC, (-1&))
    Call DeleteDC(hForgroundDC)
    
    Call DeleteObject(SelectObject(hBackgroundDC, hPrevBackgroundBmp))
    Call RestoreDC(hBackgroundDC, (-1&))
    Call DeleteDC(hBackgroundDC)
    
    Call DeleteObject(SelectObject(hStretchDC, hPrevStretchBmp))
    Call RestoreDC(hStretchDC, (-1&))
    Call DeleteDC(hStretchDC)
End Function
Public Property Let MDIForm(ByVal Frm As MDIForm)

    Set mfrmMDI = Frm

End Property
Public Property Let ExeName(ByVal strValue As String)

    mstrExeName = strValue

End Property

Public Property Let ForegroundBitmapID(ByVal lngValue As Long)

    mlngForegroundBmpID = lngValue

End Property

Public Property Let BackgroundBitmapID(ByVal lngValue As Long)

    mlngBackgroundBmpID = lngValue

End Property


Public Property Let TransparentColor(ByVal lngValue As Long)

    mlngTransparentColor = lngValue

End Property



Public Property Let RunningInIDE(ByVal blnValue As Boolean)

    mblnRunningIDE = blnValue

End Property


