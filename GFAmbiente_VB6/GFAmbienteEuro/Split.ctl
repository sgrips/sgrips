VERSION 5.00
Begin VB.UserControl SgripsSplit 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "Split.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Split.ctx":0014
   Begin VB.PictureBox pic 
      BackColor       =   &H00C0E0FF&
      Height          =   3615
      Left            =   960
      ScaleHeight     =   3555
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   0
      Width           =   100
   End
End
Attribute VB_Name = "SgripsSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private pStartMove As Boolean
'Property Variables:
Dim m_leftObject As String
'Default Property Values:
Const m_def_leftObject = "0"

Private mMe As Object
Private pLeft As Object
Private pRight As Object

Private posX As Integer

Private Sub UserControl_Initialize()
    pic.Left = 0
    pic.Top = 0
    pStartMove = False

End Sub


Public Property Set LeftObj(Obj As Object)
    Set pLeft = Obj
End Property

Public Property Set RightObj(Obj As Object)
    Set pRight = Obj
End Property




Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pStartMove = True
    posX = X
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim posX As Long
    
    pic.MousePointer = vbSizeWE
    
    posX = (Sgn(X - posX)) * Abs(X - posX)
    If pStartMove And mMe.Left + posX > mMe.Width And mMe.Left + posX < pRight.Container.ScaleWidth - 2 * mMe.Width Then
        mMe.Left = mMe.Left + posX
        resizeobject

    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pStartMove = False
End Sub

Private Sub UserControl_Resize()
    If Not mMe Is Nothing Then
        If Not pLeft Is Nothing Then
            mMe.Left = pLeft.Left + pLeft.Width
            mMe.Top = pLeft.Top
            mMe.Height = pLeft.Height
            pic.Height = pLeft.Height
            resizeobject
        End If
    Else
        UserControl.Width = pic.Width
        pic.Height = UserControl.Height
    End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_leftObject = PropBag.ReadProperty("leftObject", m_def_leftObject)
End Sub

Public Sub attachControl(Left As Object, Right As Object)
    Dim c  As Object
    
    Set pLeft = Left
    Set pRight = Right
    For Each c In UserControl.ParentControls
        If TypeOf c Is SgripsSplit Then
            Set mMe = c
            Exit For
        End If
    Next
    
    pLeft.Left = 0
    resizeobject
End Sub

Private Sub resizeobject()
    pLeft.Height = mMe.Height
    pLeft.Top = mMe.Top
    pLeft.Width = mMe.Left - pLeft.Left
    
    pRight.Height = mMe.Height
    pRight.Top = mMe.Top
    pRight.Left = mMe.Left + mMe.Width
    
    pRight.Width = pRight.Container.ScaleWidth - pLeft.Width - pLeft.Left - mMe.Width
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("leftObject", m_leftObject, m_def_leftObject)
End Sub

'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get leftObject() As String
Attribute leftObject.VB_Description = "Returns/sets the index of the currently selected item in the control."
Attribute leftObject.VB_ProcData.VB_Invoke_Property = ";List"
    leftObject = m_leftObject
End Property

Public Property Let leftObject(ByVal New_leftObject As String)
    m_leftObject = New_leftObject
    PropertyChanged "leftObject"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_leftObject = m_def_leftObject


End Sub



Public Sub Refresh()
    UserControl_Resize
End Sub

