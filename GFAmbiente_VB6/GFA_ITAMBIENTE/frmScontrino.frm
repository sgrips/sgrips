VERSION 5.00
Begin VB.Form frmScontrino 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scontrino"
   ClientHeight    =   7575
   ClientLeft      =   14955
   ClientTop       =   3855
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   105
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   6240
      Width           =   3615
   End
End
Attribute VB_Name = "frmScontrino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Type RECT
left As Long
top As Long
Right As Long
Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1

Private Sub Command1_Click()
    Me.Visible = False
End Sub

Public Sub showMe()
    Picture1.AutoRedraw = True
    
    
    Me.Move Screen.width - Me.width, Screen.Height - 1000 - 800, Me.width, 1000
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE + SWP_NOACTIVATE
    
    Picture1.top = Me.ScaleHeight - 200
    Picture1.Height = 10000
    Picture1.FontName = "ARIAL"
    Picture1.fontSize = 10

End Sub

Sub CancellaScontrino()
    Picture1.Cls
    Picture1.CurrentX = 0
    Picture1.CurrentY = 0
    
    Picture1.top = Me.ScaleHeight - 200
End Sub

Sub PrintText(strText As String, left As Long, top As Long, center As Boolean)
    Dim rct As RECT
    Dim p As Long
    Dim HeightScontrino As Long
    
    HeightScontrino = Picture1.TextHeight(strText)
    
    Me.top = Me.top - Picture1.TextHeight(strText)
    Me.Height = Me.Height + Picture1.TextHeight(strText)
    

'    If Picture1.TextWidth(strText) > Picture1.width Then
'        Picture1.CurrentY = Picture1.CurrentY + Picture1.TextHeight(strText)
 '       Picture1.top = Picture1.top - Picture1.TextHeight(strText) * 2
 '   Else
    Picture1.top = HeightScontrino
  '  End If
    
    
    'usage
    With rct
        .left = Picture1.ScaleX(left, vbTwips, vbPixels)
        .Right = Picture1.ScaleX(Picture1.ScaleWidth, vbTwips, vbPixels)
        .top = Picture1.ScaleY(top, vbTwips, vbPixels)
        .Bottom = .top + Picture1.TextHeight(strText)
    End With
    
    
    p = DT_WORDBREAK
    If center Then p = p Or DT_CENTER
    
    DrawText Picture1.hdc, strText, -1, rct, p
    
    Picture1.CurrentX = 0
    Picture1.CurrentY = Picture1.CurrentY + Picture1.TextHeight(strText)
    
    'If Picture1.TextWidth(strText) > Picture1.width Then
    '    Picture1.CurrentY = Picture1.CurrentY + Picture1.TextHeight(strText)
    'End If
    
End Sub


Public Sub AddPesata(Des As String, qty As String)
    
    Picture1.fontSize = 11
    Picture1.ForeColor = vbBlack
    PrintText Des, Picture1.CurrentX, Picture1.CurrentY, False
    
    
    Picture1.ForeColor = vbGreen
    Picture1.fontSize = 18
    PrintText CStr(qty), Picture1.CurrentX, Picture1.CurrentY, True
    
End Sub

