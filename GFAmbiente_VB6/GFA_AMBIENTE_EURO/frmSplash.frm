VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4290
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   7785
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Societ‡Prodotto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   3000
      End
      Begin VB.Label lblEuro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ä"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   285
      End
      Begin VB.Image imgLogo 
         Height          =   3120
         Left            =   -360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label lblCompany 
         Caption         =   "Societ‡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   3120
         Width           =   3735
      End
      Begin VB.Label lblWarning 
         Caption         =   "Avviso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versione"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   5
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Piattaforma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6030
         TabIndex        =   6
         Top             =   2160
         Width           =   1590
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Prodotto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licenza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    
    lblCompanyProduct = App.CompanyName
    lblLicenseTo = ""
    lblPlatform = "Windows NT 4.0/98/2000/XP"
    lblCopyright = App.LegalCopyright
    lblCompany = App.CompanyName & App.Comments
    lblWarning = "Per le informazioni sul Copyright vedere la finestra Informazioni su"
    lblVersion.Caption = "Versione " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
