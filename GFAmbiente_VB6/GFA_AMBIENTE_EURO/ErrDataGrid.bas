Attribute VB_Name = "ErrDataGRid"
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Nopaint As Boolean
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum DirectionFill
    Vertical
    Horizontal
End Enum

Private Declare Function GetWindowText& Lib "user32" _
          Alias "GetWindowTextA" _
          (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long)

Private Declare Function GetWindowTextLength& Lib "user32" _
          Alias "GetWindowTextLengthA" _
          (ByVal hwnd As Long)



Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Const LF_FACESIZE = 32

Private Const OPAQUE = 2


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
        lfFaceName As String * LF_FACESIZE '(1 To LF_FACESIZE) As Byte
End Type
Private Const TRANSPARENT = 1

Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_ULTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_REGULAR = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_DEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_ULTRABOLD = 800
Private Const FW_HEAVY = 900
Private Const FW_BLACK = 900
Private Const PROOF_QUALITY = 2

Public lWidth As Long
Public lHeight As Long

Sub rotate(pic As PictureBox, s As String, size As Long)
    Dim rc As RECT
    Dim hfnt As Long, hfntPrev As Long
    Dim lpszRotate As String
    Dim plf As LOGFONT
    Dim l As Long
    Dim angle As Integer
    
    
    lpszRotate = s
    plf.lfFaceName = pic.FontName & Chr$(0)
    plf.lfHeight = (size * pic.Height) / lHeight
    plf.lfWeight = FW_NORMAL
    plf.lfItalic = True
    plf.lfQuality = PROOF_QUALITY
    l = GetClientRect(pic.hwnd, rc)
     
    l = SetBkMode(pic.hdc, TRANSPARENT)
     
    angle = 900

    plf.lfEscapement = angle
    hfnt = CreateFontIndirect(plf)
    hfntPrev = SelectObject(pic.hdc, hfnt)
    l = TextOut(pic.hdc, (rc.Right / 2) - pic.TextWidth("X"), rc.Bottom, lpszRotate, Len(lpszRotate))
    l = SelectObject(pic.hdc, hfntPrev)
    l = DeleteObject(hfnt)

 
 
    l = SetBkMode(pic.hdc, OPAQUE)

End Sub


Public Function GetCellText(GridControl As DataGrid) As String
      'This function takes a DataGrid control and checks its
      'edit status. The hWndEditor property is only valid
      'when the grid is being edited.

      Dim st As String
      Dim ch As Long

      If GridControl.EditActive Then
          ch = GetWindowTextLength(GridControl.hWndEditor) + 1
          st = Space(ch)

          ch = GetWindowText(GridControl.hWndEditor, st, ch)

          GetCellText = Trim(st)

      End If

End Function

Public Sub DrawBackgroundPattern(hwnd As Long, ByVal Direction As DirectionFill)

  Dim hdc As Long
  Dim rectFill As RECT     ' Rectangle for filling band
  Dim rectClient As RECT        'Rectangle for entire client area
  Dim fStep As Double
  Dim hBrush As Long
  Dim iOnBand As Integer   'Loop index
  Dim l As Long
   
  

  Nopaint = True
  hdc = GetDC(hwnd)

  ' How large is the area you need to fill?
  l = GetClientRect(hwnd, rectClient)

  'Determine how large each band should be in order to cover the
  'client with 256 bands (one for every color intensity level)
  
    If Direction = Horizontal Then
        fStep = rectClient.Bottom / 256
    Else
        fStep = rectClient.Right / 256
    End If
    
   ' Start filling bands
    Do While iOnBand < 256
        ' Set the location of the current band
        If Direction = Horizontal Then
            l = SetRect(rectFill, 0, iOnBand * fStep, rectClient.Right + 1, (iOnBand + 1) * fStep)
        Else
            l = SetRect(rectFill, iOnBand * fStep, 0, (iOnBand + 1) * fStep, rectClient.Bottom + 1)
        End If
        iOnBand = iOnBand + 1
        'Create a brush with the appropriate color for this band
        hBrush = CreateSolidBrush(RGB(Abs(255 - iOnBand), Abs(255 - iOnBand), Abs(255 - iOnBand)))
        
        
        'Fill the rectangle
        l = FillRect(hdc, rectFill, hBrush)
        
        'Get rid of the brush you created
        l = DeleteObject(hBrush)
    Loop

  ' Give back the DC
  l = ReleaseDC(hwnd, hdc)

    Nopaint = False

End Sub



