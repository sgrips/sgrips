VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveLock 1.86 KeyGenerator"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5895
      TabIndex        =   21
      Top             =   2685
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Key"
      Height          =   375
      Left            =   4620
      TabIndex        =   20
      Top             =   2685
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select HashAlgorithm"
      Height          =   2160
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB2"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA2"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties and LiberationKey"
      Height          =   2940
      Left            =   30
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Generate key File"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   2490
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   4
         Left            =   1695
         TabIndex        =   0
         Top             =   345
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   1
         Left            =   1695
         TabIndex        =   2
         Top             =   1125
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   2
         Left            =   1695
         TabIndex        =   3
         Top             =   1650
         Width           =   2535
      End
      Begin VB.TextBox Text 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   0
         Left            =   1695
         TabIndex        =   1
         Top             =   765
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "SoftwareName:"
         Height          =   255
         Left            =   150
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SoftwarePassword:"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label lblLiberationKeyLenght 
         Alignment       =   1  'Right Justify
         Caption         =   "LiberationKeyLength:"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SoftwareCode:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "LiberationKey:"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2025
         Width           =   1455
      End
   End
   Begin VB.Label lblPanel 
      Caption         =   "lblPanel"
      Height          =   255
      Left            =   540
      TabIndex        =   19
      Top             =   3945
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "Form1.frx":0442
      Top             =   3075
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":0884
      Height          =   495
      Left            =   570
      TabIndex        =   16
      Top             =   3180
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************** _
  * This KeyGenerator for Activelock17b is made by                      * _
  * Mads Moen                                                           * _
  * My e-mailadress: mads@industri.no                                   * _
  * My homepage: http://www.industri.no                                 * _
  * Code is made from Mr. Nelson Correa de Toledo Ferraz`s posting to   * _
  * http://groups.yahoo.com/group/activelock Sat Apr 20, 2002           * _
  * Feel free to reuse and improve my KeyGenerator                      * _
  ***********************************************************************
Option Explicit

Dim mi_HashAlgorithm As Integer ' the "mi_" prefix stands for:
                                ' scope = Module, type = Integer

Private Sub Command1_Click()
    Dim F1 As Variant
    'Making some variables
    Dim strKey As String
    Dim strsoftwarename As String
    Dim strSoftwareCode As String
    Dim strsoftwarepassword As String
    Dim intLiberationKeyLength As Integer
    
    'Catching the Softwarename, make sure you type it exactly as you _
     set it in ActiveLock propertypages for your software
    strsoftwarepassword = Me.Text(0)
    
     'Catching the Softwarename, make sure you type it exactly as you _
     set it in ActiveLock propertypages for your software
    strsoftwarename = Me.Text(4)
    
      
    'Catching the length of LiberationKey, make sure you _
     type it exactly as you set it in ActiveLock propertypages _
     for your software
    intLiberationKeyLength = Me.Text(1)
    
    'Here you type or paste your softwarecode under runtime _
     Remember "exactly" is very important.
    strSoftwareCode = Me.Text(2)

    'Here we call for Function "hash"
    strKey = hash(strSoftwareCode & strsoftwarepassword & strsoftwarename)
    
    'Here we get the proper part of the generated LiberationKey based _
     on what we type in for LiberationKeyLength under runtime.
    strKey = UCase(Left(strKey, intLiberationKeyLength))
    
    'Printing the key to the textbox.
    Text(3).Text = strKey
    
    'Selecting the LiberatingKey in the textbox, now you just copy it _
     and place it where you decide
    Text(3).SetFocus
    Text(3).SelStart = 0
    Text(3).SelLength = Len(Text(3).Text)
    
    If Check1.Value = 1 Then
        F1 = FreeFile
        Open App.Path & "\" & Text(2).Text & ".key" For Output As #F1
            Print #F1, Text(3).Text
        Close #F1
    End If
        
    
    'Save default settings for next session
    'SaveSetting "KeyGenerator", "Default", "SoftwareName", strsoftwarepassword
    'SaveSetting "KeyGenerator", "Default", "LiberationKeyLength", intLiberationKeyLength
    'SaveSetting "KeyGenerator", "Default", "HashAlgorithm", mi_HashAlgorithm
End Sub
Private Function hash(ByVal strHashThis As String)

'Here we call for LiberatingKeys by several functions made by ActiveLock _
 author based on the algorithm we choose

Select Case mi_HashAlgorithm
    Case 0
        hash = SHA1AA1Hash(strHashThis)
        
    Case 1
        hash = SHA1AA2Hash(strHashThis)
    Case 2
        hash = MD5AA1Hash(strHashThis)
    Case 3
        hash = MD5AA2Hash(strHashThis)
    Case 4
        hash = MD5AB1Hash(strHashThis)
    Case 5
        hash = MD5AB2Hash(strHashThis)
    Case Else
        hash = SHA1AA1Hash(strHashThis)  'Default
End Select

End Function

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Dim strsoftwarepassword As String
Dim intLiberationKeyLength As Integer, intHashAlgorithm As Integer

'Making some text in the stausbarpanel..
lblPanel = "KeyGenerator for ActiveLock 1.7.2, Made by " & App.Comments & " @ http://www.industri.no"

'Get default values from previous session
'strsoftwarepassword = GetSetting("KeyGenerator", "Default", "SoftwareName", "My SoftWareName")
'intLiberationKeyLength = GetSetting("KeyGenerator", "Default", "LiberationKeyLength", 16)
'intHashAlgorithm = GetSetting("KeyGenerator", "Default", "HashAlgorithm", 0)

intLiberationKeyLength = 16
'Load some standard text
Me.Text(0).Text = strsoftwarepassword
Me.Text(1).Text = intLiberationKeyLength
optHashAlgorithm(intHashAlgorithm) = True

End Sub

Private Sub optHashAlgorithm_Click(Index As Integer)
'Setting conditions for the select event in Function "hash"
mi_HashAlgorithm = optHashAlgorithm(Index).Index
End Sub

Private Sub optHashAlgorithm_KeyPress(Index As Integer, KeyAscii As Integer)
'With this code EnterKey act like TAB key in the _
 optHashAlgorithm - optionbuttonarray
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optHashAlgorithm_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Calling proceduer under Command1 button when _
 a option in the optHashAlgorithm array is picked with spacebar

Command1_Click
End Sub

Private Sub optHashAlgorithm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Calling proceduer under Command1 button when _
 a option in the optHashAlgorithm array is picked with mousepointer

Command1_Click
End Sub

Private Sub Text_GotFocus(Index As Integer)
Dim I
'Selecting the all the text in textboxarray while _
 the box is getting focus
For I = 0 To Text.Count - 1
  Text(I).SelStart = 0
  Text(I).SelLength = Len(Text(I).Text)
Next I
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
'With this code EnterKey act like TAB key in the Text - textboxarray
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
