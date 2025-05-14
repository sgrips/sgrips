VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frmMessageBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Non mostrare più al prossimo riavvio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Si"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   120
      Picture         =   "frmMessageBox.frx":1CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3585
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pConferma As VbMsgBoxResult
Private pFlgMostraMsg As Boolean

Public Function showMe(message As String, settingKey As String) As VbMsgBoxResult
    lblInfo.Caption = message
    
    If settingKey = vbNullString Then Check1.enabled = False
    
    Me.Show vbModal
    If settingKey <> vbNullString Then SaveSetting App.Title, "IMPOSTAZIONI", settingKey, pFlgMostraMsg
    
    showMe = pConferma
End Function

Private Sub Check1_Click()
    pFlgMostraMsg = Check1.Value = 0
End Sub

Private Sub cmd_Click(Index As Integer)
    pConferma = IIf(Index = 0, VbMsgBoxResult.vbNo, VbMsgBoxResult.vbYes)
    Unload Me
End Sub

Private Sub Form_Load()
    pFlgMostraMsg = True
End Sub

