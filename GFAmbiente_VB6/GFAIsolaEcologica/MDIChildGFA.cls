VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDIChildGFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetParent Lib "user32" _
        (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Const GWL_STYLE = (-16)
Const WS_POPUP = &H80000000


Private frmMDI As MDIForm
Private hWndMDIChildForm As Long

Public Function ShowChildForm(ByVal hWndMDIForm As Long, ByVal hWndMDIChildForm As Long, MDI As Object)
    Dim ret As Long
    Dim CurrentStyle As Long
    
    
    Set frmMDI = MDI
    
    ret = SetParent(frmContainer.hwnd, hWndMDIForm)
    ret = SetParent(hWndMDIChildForm, frmContainer.hwnd)
    
    
    CurrentStyle = GetWindowLong(frmContainer.hwnd, GWL_STYLE)
    ret = SetWindowLong(frmContainer.hwnd, GWL_STYLE, CurrentStyle Or WS_POPUP)
    frmContainer.Show
    
    Refresh
    
End Function

Public Sub Refresh()
    
    If frmMDI Is Nothing Then Exit Sub
    
    frmContainer.Move 10, 600, frmMDI.ScaleWidth, frmMDI.ScaleHeight
    frmContainer.WindowState = vbMaximized
End Sub

Public Sub CloseW()
    Dim ret As Long
    ret = SetParent(frmContainer.hwnd, 0)
    ret = SetParent(hWndMDIChildForm, 0)
End Sub



