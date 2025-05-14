VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ActiveLock"
   ClientHeight    =   4785
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4710
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3302.69
   ScaleMode       =   0  'User
   ScaleWidth      =   4422.935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtLicense 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmAbout.frx":030A
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "All rights reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "DataLock Copyright (c) 2002 Apyl Software && Systems Ltd"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "ActiveLock Copyright (c) 1998-2002, Nelson Ferraz"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   112.686
      X2              =   4282.077
      Y1              =   2733.261
      Y2              =   2733.261
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "August / 2002"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4282.077
      Y1              =   1024.973
      Y2              =   1024.973
   End
   Begin VB.Label lblTitle 
      Caption         =   "lblTitle"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   4282.077
      Y1              =   1035.326
      Y2              =   1035.326
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   112.686
      X2              =   4282.077
      Y1              =   2733.261
      Y2              =   2733.261
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  lblTitle.Caption = App.Title & " version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
