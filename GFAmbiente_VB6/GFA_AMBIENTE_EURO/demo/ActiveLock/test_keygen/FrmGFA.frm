VERSION 5.00
Begin VB.Form frmGFAKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GFAmbiente Generazione Chiave"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FrmGFA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select HashAlgorithm"
      Height          =   2160
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA2"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB2"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Genera File"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Genera Chiave"
      Default         =   -1  'True
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chiudi"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   $"FrmGFA.frx":0442
      Height          =   735
      Left            =   630
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmGFA.frx":04D6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Chiave Sblocco:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Codice Software:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   195
      Width           =   1215
   End
End
Attribute VB_Name = "frmGFAKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mi_HashAlgorithm As Integer ' the "mi_" prefix stands for:
                                ' scope = Module, type = Integer


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Text(3).Text = getKeyRegisterGFA(Me.Text(2))
End Sub

Public Function getKeyRegisterGFA(codSoftware As String) As String
    Dim F1 As Variant
    'Making some variables
    Dim strKey As String
    Dim strsoftwarename As String
    Dim strSoftwareCode As String
    Dim strsoftwarepassword As String
    Dim intLiberationKeyLength As Integer
    
    'Catching the Softwarename, make sure you type it exactly as you _
     set it in ActiveLock propertypages for your software
    strsoftwarepassword = "GFAmbiente"
    
     'Catching the Softwarename, make sure you type it exactly as you _
     set it in ActiveLock propertypages for your software
    strsoftwarename = "Bollettazione"
    
      
    'Catching the length of LiberationKey, make sure you _
     type it exactly as you set it in ActiveLock propertypages _
     for your software
    intLiberationKeyLength = 16
    
    'Here you type or paste your softwarecode under runtime _
     Remember "exactly" is very important.
    strSoftwareCode = codSoftware

    'Here we call for Function "hash"
    strKey = hash(strSoftwareCode & strsoftwarepassword & strsoftwarename)
    
    'Here we get the proper part of the generated LiberationKey based _
     on what we type in for LiberationKeyLength under runtime.
    strKey = UCase(Left(strKey, intLiberationKeyLength))
    
    'Printing the key to the textbox.
    getKeyRegisterGFA = strKey
    
    'Selecting the LiberatingKey in the textbox, now you just copy it _
     and place it where you decide
    If Me.Visible Then
        Text(3).SetFocus
        Text(3).SelStart = 0
        Text(3).SelLength = Len(Text(3).Text)
    
        If Check1.Value = 1 Then
            F1 = FreeFile
            Open App.Path & "\" & Text(2).Text & ".gfa" For Output As #F1
                Print #F1, Text(3).Text
            Close #F1
        End If
            
        
        'Save default settings for next session
        'SaveSetting "KeyGenerator", "Default", "SoftwareName", strsoftwarepassword
        'SaveSetting "KeyGenerator", "Default", "LiberationKeyLength", intLiberationKeyLength
        'SaveSetting "KeyGenerator", "Default", "HashAlgorithm", mi_HashAlgorithm
    End If
End Function



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

