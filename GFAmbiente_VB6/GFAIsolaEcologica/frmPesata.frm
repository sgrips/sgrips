VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{55E28A67-2857-49A5-8466-5D35ED0A061F}#1.0#0"; "PlexaNetKProj1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesata 
   Caption         =   "Gestione Pesate"
   ClientHeight    =   10080
   ClientLeft      =   3885
   ClientTop       =   3015
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin PlexaNetKProj1.PlexaNetKX PlexaNetKX1 
      Height          =   855
      Left            =   12720
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
      Object.Visible         =   -1  'True
      AutoScroll      =   0   'False
      AutoSize        =   0   'False
      AxBorderStyle   =   1
      Caption         =   "PlexaNetKX"
      Color           =   -16777201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      KeyPreview      =   0   'False
      PixelsPerInch   =   96
      PrintScale      =   1
      Scaled          =   -1  'True
      DropTarget      =   0   'False
      HelpFile        =   ""
      ScreenSnap      =   0   'False
      SnapBuffer      =   10
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Server          =   ""
      Port            =   ""
      AcknowledgedMode=   0   'False
      SingleDeviceToConnect=   255
   End
   Begin VB.PictureBox PicTotali 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   102
      Top             =   0
      Width           =   8535
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Index           =   3
         Left            =   5760
         TabIndex        =   106
         Top             =   120
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Totale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3960
         TabIndex        =   105
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Index           =   1
         Left            =   2160
         TabIndex        =   104
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantita Totale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   103
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ricerca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   8880
      Picture         =   "frmPesata.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Scon&trino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   9840
      Picture         =   "frmPesata.frx":06EE
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Continua"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   11760
      Picture         =   "frmPesata.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   9840
      Picture         =   "frmPesata.frx":1674
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ri&stampa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   10800
      Picture         =   "frmPesata.frx":1E9F
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Esporta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   10800
      Picture         =   "frmPesata.frx":245F
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Stampa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   12720
      Picture         =   "frmPesata.frx":2C41
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Stampa"
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Peso Auto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   11760
      Picture         =   "frmPesata.frx":3201
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Rileva Pesata"
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   12720
      Picture         =   "frmPesata.frx":3C55
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Chiudi Finestra"
      Top             =   8160
      Width           =   975
   End
   Begin VB.PictureBox picContainer 
      Height          =   2535
      Index           =   1
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   7995
      TabIndex        =   63
      Top             =   6960
      Width           =   8055
      Begin VB.TextBox TxtMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdRicercaUtenteMov 
         Height          =   480
         Index           =   0
         Left            =   1560
         Picture         =   "frmPesata.frx":4207
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   240
         Width           =   480
      End
      Begin VB.TextBox TxtMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   6240
         TabIndex        =   77
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         ItemData        =   "frmPesata.frx":4971
         Left            =   6240
         List            =   "frmPesata.frx":4982
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         ItemData        =   "frmPesata.frx":49BC
         Left            =   120
         List            =   "frmPesata.frx":49CF
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cmbTargaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2280
         TabIndex        =   74
         Text            =   "Combo3"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbTargaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2280
         TabIndex        =   73
         Text            =   "Combo3"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdRicercaUtenteMov 
         Height          =   480
         Index           =   1
         Left            =   1560
         Picture         =   "frmPesata.frx":4A1B
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   960
         Width           =   480
      End
      Begin VB.TextBox TxtMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   71
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkEscludiCancellati 
         Caption         =   "Escludi Cancellati"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13320
         TabIndex        =   70
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         ItemData        =   "frmPesata.frx":5185
         Left            =   10320
         List            =   "frmPesata.frx":518D
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         ItemData        =   "frmPesata.frx":5198
         Left            =   2280
         List            =   "frmPesata.frx":51A5
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         ItemData        =   "frmPesata.frx":51CF
         Left            =   8400
         List            =   "frmPesata.frx":51E0
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         ItemData        =   "frmPesata.frx":521A
         Left            =   10560
         List            =   "frmPesata.frx":5222
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox chkIngombranti 
         Caption         =   "Ingombranti"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13320
         TabIndex        =   65
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cmbRicercaMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         ItemData        =   "frmPesata.frx":522D
         Left            =   6360
         List            =   "frmPesata.frx":5235
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1800
         Width           =   3855
      End
      Begin MSComctlLib.ListView lvwMovimenti 
         Height          =   1095
         Left            =   120
         TabIndex        =   80
         Top             =   2880
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   81
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   46268417
         CurrentDate     =   39318
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   82
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   111542273
         CurrentDate     =   39318
      End
      Begin MSComctlLib.ImageCombo ImgCombo2 
         Height          =   360
         Left            =   8280
         TabIndex        =   83
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList1"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Inzio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4200
         TabIndex        =   98
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Fine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   4200
         TabIndex        =   97
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Utente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   96
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rifiuto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   6240
         TabIndex        =   95
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stato Movimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   6240
         TabIndex        =   94
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flag Bonus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   93
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Targa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   2280
         TabIndex        =   92
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipologia Rifiuto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   34
         Left            =   8280
         TabIndex        =   91
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Targa Traspor."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   2280
         TabIndex        =   90
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trasportatore"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Tariffa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   10320
         TabIndex        =   88
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Utenza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   2280
         TabIndex        =   87
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registrato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   8400
         TabIndex        =   86
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Causale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   10560
         TabIndex        =   85
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Comune"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   6360
         TabIndex        =   84
         Top             =   1560
         Width           =   1350
      End
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   2
      Left            =   8280
      ScaleHeight     =   3015
      ScaleWidth      =   6735
      TabIndex        =   44
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtBadgeAccessi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   100
         Text            =   "8881234567890123"
         Top             =   2400
         Width           =   5775
      End
      Begin VB.CommandButton cmbRicerca 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   5880
         Picture         =   "frmPesata.frx":5240
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   2400
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwAccessi 
         Height          =   1095
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MMM/yy HH:mm:ss"
         Format          =   133562371
         CurrentDate     =   39318
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   47
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MMM/yy HH:mm:ss"
         Format          =   133562371
         CurrentDate     =   39318
      End
      Begin VB.Label lblBadgeAccessi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Badge - F3 Ingresso Manuale - F4 Uscita Manuale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   101
         Top             =   2160
         Width           =   4110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Inzio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Fine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   3000
         TabIndex        =   48
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.Timer tmrAccessi 
      Interval        =   1000
      Left            =   12720
      Top             =   3360
   End
   Begin VB.Timer tmrGeneraImgRD 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13200
      Top             =   3360
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   11280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   11280
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   10440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   0
      Left            =   600
      ScaleHeight     =   7215
      ScaleWidth      =   11655
      TabIndex        =   8
      Top             =   480
      Width           =   11655
      Begin VB.PictureBox picImageRD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   965
         Left            =   3120
         ScaleHeight     =   960
         ScaleWidth      =   1335
         TabIndex        =   29
         Top             =   2640
         Width           =   1335
         Begin VB.CheckBox optImgRD 
            Height          =   960
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picSearchRD 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4680
         ScaleHeight     =   975
         ScaleWidth      =   3495
         TabIndex        =   30
         Top             =   2640
         Width           =   3495
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   2
            Left            =   0
            TabIndex        =   31
            Top             =   300
            Width           =   1815
         End
         Begin MSComctlLib.ImageCombo imgCombo1 
            Height          =   450
            Left            =   1920
            TabIndex        =   32
            Top             =   300
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   794
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImageList       =   "ImageList1"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipologia Rifiuto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   7
            Left            =   1920
            TabIndex        =   34
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ricerca Tip.Rifiuto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.Frame fmeTrasportatore 
         Caption         =   "Trasportatore"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1200
         Left            =   240
         TabIndex        =   18
         Top             =   3585
         Width           =   10335
         Begin VB.CommandButton cmbRicerca 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   1
            Left            =   1440
            Picture         =   "frmPesata.frx":59AA
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   570
            Width           =   495
         End
         Begin VB.ComboBox cmbTarga 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   585
            Width           =   2175
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Codice del trasportatore che ha portato il rifiuto"
            Top             =   585
            Width           =   1215
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   4320
            TabIndex        =   5
            ToolTipText     =   "Inserire il numero formulario di ingresso"
            Top             =   600
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpickFormulario 
            Height          =   450
            Left            =   6720
            TabIndex        =   19
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   794
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   133627905
            CurrentDate     =   39352
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   1560
            TabIndex        =   28
            Top             =   0
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Targa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   33
            Left            =   2040
            TabIndex        =   23
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   26
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Formulario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   17
            Left            =   6720
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num.Formulario"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   24
            Left            =   4320
            TabIndex        =   20
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.Frame fmeUtente 
         Caption         =   "Utente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1875
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   4935
         Begin VB.CommandButton cmbRicerca 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   1440
            Picture         =   "frmPesata.frx":6114
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   420
            Width           =   495
         End
         Begin VB.TextBox txtBadge 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   27
            Text            =   "8881234567890123"
            Top             =   1200
            Width           =   4575
         End
         Begin VB.ComboBox cmbTarga 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   2160
            TabIndex        =   1
            Text            =   "Combo3"
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   0
            ToolTipText     =   "Codice dell'utente rilevato"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Targa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   25
            Left            =   2160
            TabIndex        =   17
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Badge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   10
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame fmeTotali 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2295
         Left            =   2640
         TabIndex        =   11
         Top             =   4800
         Width           =   8415
         Begin VB.Frame frmAccessi 
            Caption         =   "Totale Accessi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   5040
            TabIndex        =   42
            Top             =   720
            Width           =   3135
            Begin VB.Image imgTrafficLightRed 
               Height          =   1140
               Left            =   240
               Picture         =   "frmPesata.frx":687E
               Top             =   240
               Width           =   1140
            End
            Begin VB.Image imgTrafficLightGreen 
               Height          =   1140
               Left            =   240
               Picture         =   "frmPesata.frx":80A2
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label lblTotaleAccessi 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1125
               Left            =   1920
               TabIndex        =   43
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lblErrorUtenzaDomestica 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UTENZA NON DOMESTICA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   120
               TabIndex        =   51
               Top             =   1080
               Visible         =   0   'False
               Width           =   3675
            End
         End
         Begin VB.TextBox txtInfoUser 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1080
            Width           =   4575
         End
         Begin MSComctlLib.ListView lvwTotMovimenti 
            Height          =   495
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Tot.Pesate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Tot.Bonus"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Tot.Punti"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Max Al Giorno"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Max Al Anno"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblErrorMaxUtentiNelIsola 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERATO LIMITE MASSIMO UTENZE PIATTAFORMA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   7515
         End
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Peso rilevato"
         Top             =   2702
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvwImmobili 
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   5040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblCambiaPesa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambia Pesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1830
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   2445
         Width           =   1140
      End
      Begin VB.Label lblErrorRifiuto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATTENZIONE  !!! SELEZIONARE UN RIFIUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1320
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FATTURABILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   35
         Left            =   9240
         TabIndex        =   38
         Top             =   2040
         Width           =   2010
      End
      Begin VB.Label lblInfoCodCer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   340
         Left            =   240
         TabIndex        =   37
         Top             =   2050
         Width           =   11055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Immobili"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantità"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2450
         Width           =   690
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registrazione Movimento"
            Key             =   "aqpeso"
            Object.ToolTipText     =   "Aquisisce il  peso e le informazioni utente dalla pesa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Elenco Movimenti"
            Key             =   "elpesate"
            Object.ToolTipText     =   "Visualizza l'elenco delle pesate"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Elenco Accessi"
            Key             =   "Accessi"
            Object.ToolTipText     =   "Visualizza le elenco degli utenti presenti nell'isola ecologica"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListPesata 
      Left            =   14400
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesata.frx":9A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesata.frx":A46F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13800
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesata.frx":AEC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesata.frx":B63F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEsportazione 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Avanzamento Esportazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   25
      Top             =   10200
      Width           =   5340
   End
End
Attribute VB_Name = "frmPesata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHdrIcons As New cLVHeaderSortIcons
Private m_esportazione As Boolean
Private m_arr_optImgRD As New ColImagesRD

Private offsetX As Single
Private offsetY As Single
Private m_lastCheckImgRD As Long

Private Const CMD_ESCI = 0
Private Const CMD_LETTURAQUANTITA = 1
Private Const CMD_SALVAESTAMPA = 2
Private Const CMD_RISTAMPAMOV = 3
Private Const CMD_ESPORTAMOV = 4
Private Const CMD_CANCELLAMOV = 5
Private Const CMD_RICERCAMOV = 6
Private Const CMD_SALVAMOV = 7
Private Const CMD_MOSTRASCONTRINO = 8

Private ButtonsRicercaMov As String
Private ButtonsRicercaacc As String

Private Const LBL_TRASPOR_RAGIONESOCIALE = 28
Private Const LBL_QUANTITA = 5
Private Const LBL_IMMOBILI = 6
Private Const LBL_FATTURABILE = 35

Private Const TXT_COD_UTENTE = 0
Private Const TXT_QUANTITA = 1
Private Const TXT_NRFORMULARIO = 4
Private Const TXT_FINDCODCER = 2
Private Const TXT_COD_TRASPORTATORE = 3

Private Const TXT_FINDCODCERMOV = 1

Private Const CB_FINDSTRING As Long = &H14C

Private Esci As Boolean
Private msgPesa As messaggioPesa
Private Canceling As Boolean
Private Deleting As Boolean

Private mTipoUtenza As String
Private mForzaturaMovimento As Boolean
Private rsCodiciCER  As ADODB.Recordset
Private rsCliente As ADODB.Recordset
Private mTotResiduoFatt As Double
Private sWhereRifiuti As String
Private midMovimentiDaStampare As String
Private midDocumentoPesata As Long
Private mnrDocumentoPesata As Long
Private mlastUtente As Long
Private mlastBadge As String
Private mlastBadgeFromDeviceAccess As String
Private mIsOspiteFromDeviceAccess As Boolean
Private mIsUtenteDomestico As Boolean

Private m_codComune As String
Private m_indirizzo As String
Private m_ragioneSociale As String
Private m_badge As String
Private m_pivaCodFisc As String
Private m_badgeDTini As String
Private m_badgeDTend As String
Private m_badgeProg As String
Private m_noEvent As Boolean
Private mPesoLettoDaBilancia As Boolean
Private mRichiestaPesoInCorso As Boolean
Private rsTotaleAccessi As ADODB.Recordset
Private rsTotaleDaPesare As ADODB.Recordset


Enum ProprietaCER
    CorrispettivoBonus
    CorrispettivoBonusPunti
    UnitaDiMisura
    FlgAbilitatoFatturazione
    qtyMaxFatturabile
    qtyMaxGGUD
    qtyMaxYYUD
    qtyMaxGGUA
    qtyMaxYYUA
    flgAmmessoPerUD
    flgAmmessoPerUA
End Enum

Private Enum typeForzatura
    Forzato
    NonPresente
    NonForzatoPresente
End Enum

Private Enum typeAccessiRichiesta
    Ingresso
    Uscita
    Pulizia
    UtenteGiaRilevato
    UtenteGiaRilevatoInGiornata
    DisponibileDopoPesata
End Enum

Private Enum typeTrafficlightState
    Verde
    Rosso
End Enum

Private Enum typeAlertMessage
    WarningUtenzaNonDomestica
    ErrorRifiuto
    ErrorMaxUtentiNelIsola
    InfoCodCer
End Enum


Private Sub cmbRicercaMov_Change(Index As Integer)
    On Error GoTo cmbRicercaMov_ChangeERR
    If Index = 0 Then
        chkEscludiCancellati.enabled = False
        If cmbRicercaMov(Index).ItemData(cmbRicercaMov(Index).ListIndex) = -1 Then
            chkEscludiCancellati.enabled = True
        End If
    End If
Exit Sub
cmbRicercaMov_ChangeERR:
    ScriviLog TPL_ERRORE, App.Title, "cmbRicercaMov_Change", Err.Description

End Sub

Private Sub cmbRicercaMov_Click(Index As Integer)
    cmbRicercaMov_Change Index
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim idMovimento As Long
    Dim idAlibiMemory As String
    Dim Stampa As Boolean
    
    On Error GoTo cmd_ClickERR
    
    Select Case Index
        Case CMD_ESCI
            Esci = True
            
            If MDIMain.Timer1.enabled = True Then
                 StampaMovimenti (True)
                 If MDIMain.Timer1.enabled = True Then Exit Sub
            End If
            
            Unload Me
        Case CMD_MOSTRASCONTRINO
            frmScontrino.Visible = Not frmScontrino.Visible
            Me.SetFocus
        Case CMD_LETTURAQUANTITA
            If cmd(CMD_LETTURAQUANTITA).Caption = "&Annulla" Then
                cmd(CMD_LETTURAQUANTITA).Caption = "&Peso Auto"
                cmd(CMD_LETTURAQUANTITA).Picture = imgListPesata.ListImages(1).Picture
                cmd(CMD_ESCI).enabled = True
                picContainer(0).enabled = True
                
                cmd(CMD_SALVAESTAMPA).enabled = True
                cmd(CMD_SALVAMOV).enabled = True
                cmd(CMD_ESCI).enabled = True
                
                Esci = True
                
                mRichiestaPesoInCorso = False
                
                applicaRestrizioniUtente
            Else
                Esci = False
                If Not getCurrentPesa.ModalitaLetturaPesoAutomatica Then
                    picContainer(0).enabled = False
                    cmd(CMD_ESCI).enabled = False
                    cmd(CMD_SALVAESTAMPA).enabled = False
                    cmd(CMD_SALVAMOV).enabled = False
                End If
                
                cmd(CMD_LETTURAQUANTITA).Caption = "&Annulla"
                cmd(CMD_LETTURAQUANTITA).Picture = imgListPesata.ListImages(2).Picture
                
                
                letturaPeso gPesaIsolaCod
            End If
            
        Case CMD_SALVAESTAMPA, CMD_SALVAMOV
            
            'stampa il documento se esistono stampe in attessa e non vi è una pesata
            If Index = CMD_SALVAESTAMPA And MDIMain.Timer1.enabled = True Then
                If imgCombo1.SelectedItem Is Nothing Then
                    StampaMovimenti True
                    frmScontrino.CancellaScontrino
                    If frmScontrino.Visible Then cmd_Click (CMD_MOSTRASCONTRINO)
                    ResettaDatiUtenteDopoStampa
                    Exit Sub
                ElseIf (Mid$(imgCombo1.SelectedItem.key, 5, 2) = "00") Or Len(imgCombo1.SelectedItem.key) < 6 Then
                    StampaMovimenti True
                    If frmScontrino.Visible Then frmScontrino.CancellaScontrino
                    cmd_Click (CMD_MOSTRASCONTRINO)
                    ResettaDatiUtenteDopoStampa
                    Exit Sub
                End If
            End If
                    
            cmd(CMD_ESCI).enabled = False
            cmd(CMD_SALVAESTAMPA).enabled = False
            cmd(CMD_SALVAMOV).enabled = False
            TabStrip.enabled = False
            
            If (validaDati(idAlibiMemory)) Then
                
                If (salvaPesata(mForzaturaMovimento, idMovimento, idAlibiMemory)) Then
                    
                    frmScontrino.AddPesata imgCombo1.Text, txt(TXT_QUANTITA).Text + " " + getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura)
                    
                    midMovimentiDaStampare = midMovimentiDaStampare & idMovimento & ","
                    If Index = CMD_SALVAESTAMPA Then
                        Label1(LBL_FATTURABILE).Visible = False
                        DoEvents
                        frmScontrino.CancellaScontrino
                        If frmScontrino.Visible Then cmd_Click (CMD_MOSTRASCONTRINO)
                        StampaMovimenti True
                    Else
                        MDIMain.attivaMultiPesata (True)
                        'blocco campi per evitare l'inserimento di dati diversi dalla prima pesata escluso il Peso e il Tipo rifiuto
                        txt(TXT_COD_TRASPORTATORE).Locked = True
                        txt(TXT_NRFORMULARIO).Locked = True
                        cmbTarga(1).Locked = True
                        cmbTarga(0).Locked = True
                        lvwImmobili.enabled = False
                    End If
                                        
                    ResettaDatiUtenteDopoStampa
                End If
            End If
            
            cmd(CMD_ESCI).enabled = True
            cmd(CMD_SALVAESTAMPA).enabled = True
            cmd(CMD_SALVAMOV).enabled = True
            TabStrip.enabled = True
        Case CMD_RICERCAMOV
            If TabStrip.Tag = 1 Then
            CaricaMovimenti
            Else
            CaricaAccessi
            End If
            
        Case CMD_ESPORTAMOV
            lblEsportazione.Visible = True
            ProgressBar1.Visible = True
            Esportazione
            lblEsportazione.Visible = False
            ProgressBar1.Visible = False
        Case CMD_CANCELLAMOV
        
            
            If Not lvwMovimenti.SelectedItem Is Nothing Then
                If MsgBox("Attenzione si vuole veramente cancellare il movimento ?" & vbCrLf & "Se il movimento è già contabilizzato non può essere cancellato", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    
                    gDBC.BeginTrans
                    
                    gDBC.Execute "Update Movimenti set stato = '2' where stato ='0' and id=" & Mid$(lvwMovimenti.SelectedItem.key, 2)
                    
                    gDBC.CommitTrans
                    cmd_Click CMD_RICERCAMOV
                End If
            End If
                
            
        Case CMD_RISTAMPAMOV
            
            If Not lvwMovimenti.SelectedItem Is Nothing Then
                Dim idMovimenti  As String
                idMovimenti = getIdMovimentiFromDocumentId(Mid$(lvwMovimenti.SelectedItem.key, 2))
                
                If lvwMovimenti.SelectedItem.ListSubItems("Causale") = "Uscita" Then
                    ReportPesataUscita.Stampa idMovimenti, pfrmPesata, True
                
                Else
                    If lvwMovimenti.SelectedItem.ListSubItems("NumeroFormulario") <> "" Then
                        ReportPesataConTrasportatore.Stampa idMovimenti, pfrmPesata, True
                        Set ReportPesataConTrasportatore = Nothing
                    Else
                        ReportPesata.Stampa idMovimenti, pfrmPesata, True
                        Set ReportPesata = Nothing
                    End If
                    
                    If gStampaScontrinoNrCopie > 0 Then
                        'se attiva la stampa scontrino la faccio
                        StampaScontrino True, lvwMovimenti.SelectedItem.ListSubItems(2).Text, idMovimenti, lvwMovimenti.SelectedItem.ListSubItems(1).Text
                    End If
                End If
            End If
    End Select
    Exit Sub
cmd_ClickERR:
     ScriviLog TPL_ERRORE, App.Title, "cmd_Click", Err.Description

End Sub

Private Sub ResettaDatiUtenteDopoStampa()
    On Error GoTo ResettaDatiUtenteDopoStampaERR


    Dim codImmobile As String
      mTotResiduoFatt = 0
      
      'imposta campi form con i default
      txt(TXT_QUANTITA).Text = vbNullString
      txt(TXT_FINDCODCER).Text = vbNullString
      optImgRD(m_lastCheckImgRD).value = 0
      Set imgCombo1.SelectedItem = Nothing
      imgCombo1.SelectedItem = imgCombo1.ComboItems.item(1)
      
      If lvwImmobili.ListItems.Count > 0 Then codImmobile = lvwImmobili.SelectedItem.key
      ShowMessage InfoCodCer, vbNullString
      
      getTotaliPerCodiceCer (0)
      
      caricaDatiUtente txt(TXT_COD_UTENTE)
      
    
      Label1(LBL_FATTURABILE).Visible = False
      If lvwImmobili.ListItems.Count > 0 Then
          If (lvwImmobili.ListItems(codImmobile).Checked = False) Then
              lvwImmobili.ListItems(codImmobile).Checked = True
              lvwImmobili_ItemCheck lvwImmobili.ListItems(codImmobile)
          End If
      End If
      If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!", vbExclamation, App.Title
      
      MDIMain.refreshForm
      
      'mi riposiziono sul codbadge
      DoEvents
      If Not getCurrentPesa.BadgeInStringaPesata Then txtBadge.SetFocus
Exit Sub
ResettaDatiUtenteDopoStampaERR:
         ScriviLog TPL_ERRORE, App.Title, "ResettaDatiUtenteDopoStampa", Err.Description

End Sub

Public Sub StampaMovimenti(Cancel As Boolean, Optional ByVal codUtente As String = vbNullString)
    On Error GoTo StampaMovimentiERR
    
    If MDIMain.Timer1.enabled = True And Cancel And gStampaPesataNrCopie > 0 Then
        If MsgBox("Attenzione, si vuole procedere con la stampa dei movimenti per l'utente:" & m_ragioneSociale, vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    End If
    
    
    'aggiorna riferimento sulla tabella documenti
    AggiornaRiferimentoInDocumento midDocumentoPesata, mnrDocumentoPesata, midMovimentiDaStampare, "I"
    
    If (txt(TXT_COD_TRASPORTATORE).enabled And txt(TXT_COD_TRASPORTATORE).Text <> "") Then
        ReportPesataConTrasportatore.Stampa midMovimentiDaStampare, pfrmPesata, CBool(gStampaPesataAnteprima), gStampaPesataNrCopie
    Else
        ReportPesata.Stampa midMovimentiDaStampare, pfrmPesata, CBool(gStampaPesataAnteprima), gStampaPesataNrCopie
    End If
    
    Unload ReportPesata
    Unload ReportPesataConTrasportatore
    
    If gStampaScontrinoNrCopie > 0 Then
        'se attiva la stampa scontrino la faccio
        StampaScontrino MDIMain.Timer1.enabled = True And Cancel, m_badge, midMovimentiDaStampare, m_ragioneSociale
    End If
    
    MDIMain.attivaMultiPesata (False)
    
    'Ri-Attivo i campi
    txt(TXT_COD_TRASPORTATORE).Locked = False
    txt(TXT_NRFORMULARIO).Locked = False
    cmbTarga(1).Locked = False
    cmbTarga(0).Locked = False
    lvwImmobili.enabled = True
    
    'CLEAR fields trasportatore
    txt(TXT_COD_TRASPORTATORE).Text = vbNullString
    txt(TXT_NRFORMULARIO).Text = vbNullString
    
    'cancello le variabili di appoggio
    GestioneAccessi mlastBadge, False, False, False, DisponibileDopoPesata
    mlastUtente = -1
    mlastBadge = vbNullString
    midDocumentoPesata = 0
    mnrDocumentoPesata = 0
    midMovimentiDaStampare = vbNullString
    txt(TXT_COD_UTENTE).Text = codUtente
    
    caricaDatiUtente codUtente
        
    caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE)
    
    caricaImmobili codUtente
    
    'annullo lettura pesata
    If Not getCurrentPesa.ModalitaLetturaPesoAutomatica And cmd(CMD_LETTURAQUANTITA).Caption = "&Annulla" Then cmd_Click (CMD_LETTURAQUANTITA)
    DoEvents
Exit Sub
StampaMovimentiERR:
    ScriviLog TPL_ERRORE, App.Title, "StampaMovimenti", Err.Description
    
End Sub

Public Sub StampaScontrino(Cancel As Boolean, codBadge As String, midMovimentiDaStampare As String, m_ragioneSociale As String)
        
    On Error GoTo StampaScontrinoERR

    If Cancel Then
        If MsgBox("Attenzione, si vuole procedere con la stampa dello scontrino per l'utente:" & m_ragioneSociale, vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    End If
    
    
    ReportScontrino.Stampa midMovimentiDaStampare, codBadge, pfrmPesata, CBool(gStampaScontrinoAnteprima), gStampaScontrinoNrCopie
         
    
    Unload ReportScontrino
    Exit Sub
StampaScontrinoERR:
    ScriviLog TPL_ERRORE, App.Title, "StampaScontrino", Err.Description

End Sub




Private Sub Esportazione()
    On Error GoTo EsportazioneERR
    
    Dim item As ListItem
    Dim ItemH As ColumnHeader
    Dim ItemS As ListSubItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim mChecked As Long
    
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gMovimentiFile & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
   
    If lvwMovimenti.ListItems.Count = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = lvwMovimenti.ListItems.Count
    ProgressBar1.value = 0
    
    For Each ItemH In lvwMovimenti.ColumnHeaders
        txtStream.Write ItemH.Text
        If ItemH.Index < lvwMovimenti.ColumnHeaders.Count Then
            txtStream.Write gMovimentiSeparatore
        End If
    Next
    txtStream.Write vbCrLf
    
    For Each item In lvwMovimenti.ListItems
        ProgressBar1.value = ProgressBar1.value + 1
        txtStream.Write Chr$(34) & item.Text & Chr$(34)
        txtStream.Write gMovimentiSeparatore
        For Each ItemS In item.ListSubItems
            
            txtStream.Write Chr$(34) & IIf(ItemS.Tag <> "", ItemS.Tag, ItemS.Text) & Chr$(34)
            If ItemS.Index < item.ListSubItems.Count Then
                txtStream.Write gMovimentiSeparatore
            End If

        Next
        txtStream.Write vbCrLf
    Next
    
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    
    fso.MoveFile fileEXP, gPercorsoEsp & "\" & gMovimentiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
    
    
    Set fso = Nothing
    
    
    
    CaricaMovimenti
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title
    Exit Sub
EsportazioneERR:
    ScriviLog TPL_ERRORE, App.Title, "Esportazione", Err.Description

End Sub

Private Function validaDati(ByRef idAlibiMemory As String) As Boolean
On Error GoTo validaDatiERR
    

    Dim campoInErrore As Boolean
    Dim MaxGGUD As Double
    Dim MaxGGUA As Double
    Dim MaxYYUA As Double
    Dim MaxYYUD As Double
    Dim MaxFatturabileUA As Double
    Dim TotMovGG As Double
    Dim TotMovYY As Double
    Dim TotMovYYPerImmobile As Double
    Dim TotResiduo As Double
    Dim quantita As Double
    Dim rs As ADODB.Recordset
    Dim ColorError As Long
    
    validaDati = False
    campoInErrore = False
    
    txt(TXT_COD_UTENTE).ToolTipText = ""
    txt(TXT_QUANTITA).ToolTipText = ""
    imgCombo1.ToolTipText = ""
    lvwImmobili.ToolTipText = ""
    
    dtpickFormulario.ToolTipText = ""
    cmbTarga(1).ToolTipText = ""
    cmbTarga(0).ToolTipText = ""
    
    txtBadge.ToolTipText = "Valido dal " & m_badgeDTini & " al " & m_badgeDTend
    txt(TXT_COD_UTENTE).BackColor = -2147483643
    'txt(TXT_QUANTITA).BackColor = -2147483643
    txt(TXT_QUANTITA).BackColor = vbBlack
    imgCombo1.BackColor = -2147483643
    lvwImmobili.BackColor = -2147483643
    txtBadge.BackColor = -2147483643
    
    dtpickFormulario.CalendarBackColor = -2147483643
    cmbTarga(1).BackColor = -2147483643
    cmbTarga(0).BackColor = -2147483643
    lblErrorRifiuto.Visible = False
    lblErrorUtenzaDomestica.Visible = False
    lblErrorMaxUtentiNelIsola.Visible = False
    
    ColorError = &H80C0FF
    
    If (txt(TXT_COD_TRASPORTATORE).enabled And txt(TXT_COD_TRASPORTATORE).Text <> vbNullString) Then
    'validazioni su trasportatore
        If (txt(TXT_NRFORMULARIO).Text = "") Then
            txt(TXT_NRFORMULARIO).ToolTipText = "Attenzione numero formulario è obbligatorio"
            txt(TXT_NRFORMULARIO).BackColor = ColorError
            txt(TXT_NRFORMULARIO).SetFocus
            
            campoInErrore = True
        End If
        
        If (IsNull(dtpickFormulario.value)) Then
            dtpickFormulario.ToolTipText = "Attenzione la data formulario è obbligatoria"
            dtpickFormulario.CalendarBackColor = ColorError
            dtpickFormulario.SetFocus
            
            campoInErrore = True
        End If

        If (cmbTarga(1).Text = "") Then
            cmbTarga(1).ToolTipText = "Attenzione la targa è obbligatoria"
            cmbTarga(1).BackColor = ColorError
            cmbTarga(1).SetFocus
            
            campoInErrore = True
        Else
            Set rs = gDBC.Execute("select codutente,codSoggetto from automezzi where codSoggetto='U' and  codutente<>" & txt(TXT_COD_TRASPORTATORE) & " and targa='" & cmbTarga(1).Text & "'")
            If Not rs.EOF Then
                cmbTarga(1).ToolTipText = "Attenzione la targa è già presente ed assegnata al utente:" & rs(0) & " tipo:" & rs(1)
                cmbTarga(1).BackColor = ColorError
                cmbTarga(1).SetFocus
                
                campoInErrore = True
            End If
            rs.Close
        End If
    End If
    
    If txt(TXT_COD_UTENTE) = "" Then
        txt(TXT_COD_UTENTE).ToolTipText = "Attenzione codice utente obbligatorio"
        txt(TXT_COD_UTENTE).BackColor = ColorError
        txt(TXT_COD_UTENTE).SetFocus

        campoInErrore = True
    End If
    
    If (cmbTarga(0).Text <> "") Then
        Set rs = gDBC.Execute("select codutente,codSoggetto from automezzi where codSoggetto='U' and  codutente<>" & txt(TXT_COD_UTENTE) & " and targa='" & cmbTarga(0).Text & "'")
        If Not rs.EOF Then
            cmbTarga(0).ToolTipText = "Attenzione la targa è già presente ed assegnata al utente:" & rs(0) & " tipo:" & rs(1)
            cmbTarga(0).BackColor = ColorError
            cmbTarga(0).SetFocus
                
            campoInErrore = True
        End If
        rs.Close
    End If
    
    If Format$(Date, "YYYY/MM/DD") < Format$(m_badgeDTini, "YYYY/MM/DD") Then
        txtBadge.ToolTipText = txtBadge.ToolTipText & " [badge non ancora attivo]"
        txtBadge.BackColor = ColorError
        
        If campoInErrore = False Then txt(TXT_COD_UTENTE).SetFocus
        campoInErrore = True
    End If
    
    If Format$(Date, "YYYY/MM/DD") > Format$(m_badgeDTend, "YYYY/MM/DD") Then
    
        txtBadge.ToolTipText = txtBadge.ToolTipText & "[badge scaduto]"
        txtBadge.BackColor = ColorError
        
        If campoInErrore = False Then txt(TXT_COD_UTENTE).SetFocus
        campoInErrore = True
    End If
    
    If imgCombo1.SelectedItem Is Nothing Then
        imgCombo1.ToolTipText = "Attenzione il tipo rifiuto è obbligatorio"
        imgCombo1.BackColor = ColorError
        
        If campoInErrore = False And picSearchRD.Visible Then imgCombo1.SetFocus
        campoInErrore = True
        ShowMessage ErrorRifiuto, imgCombo1.ToolTipText
    ElseIf Mid$(imgCombo1.SelectedItem.key, 5, 2) = "00" Or Len(imgCombo1.SelectedItem.key) < 6 Then
            imgCombo1.ToolTipText = "Attenzione il tipo rifiuto non può essere una categoria o una sottogategoria"
            ShowMessage ErrorRifiuto, imgCombo1.ToolTipText
            imgCombo1.BackColor = ColorError
            If campoInErrore = False Then imgCombo1.SetFocus
            campoInErrore = True
    End If
    
    If lvwImmobili.Tag = vbNullString Then
        lvwImmobili.ToolTipText = "Attenzione l'immobile è obbligatorio"
        lvwImmobili.BackColor = ColorError
        If campoInErrore = False Then lvwImmobili.SetFocus
        campoInErrore = True
    End If
    
    If CDbl("0" & txt(TXT_QUANTITA)) <= 0 Then
        txt(TXT_QUANTITA).ToolTipText = "Attenzione la quantità deve essere > 0"
        txt(TXT_QUANTITA).BackColor = ColorError
        If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
        campoInErrore = True
    ElseIf campoInErrore = False Then
        
        
        'verifico se il peso è stabile prima di acquisire il peso
        If Not msgPesa.PesoStabile Then
            MsgBox "Attenzione il peso non è stabile, attendere che si stabilizzi !!"
            If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
            campoInErrore = True
        Else
            'verifico se abilitato alla registrazione movimento
            If mTipoUtenza = "D" And getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.flgAmmessoPerUD) = False Then
                txt(TXT_QUANTITA).ToolTipText = "Attenzione l'Utenza Domestica non è abilitata a portare il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text
                txt(TXT_QUANTITA).Text = 0
                txt(TXT_QUANTITA).BackColor = ColorError
                MsgBox txt(TXT_QUANTITA).ToolTipText, vbExclamation + vbOKOnly, App.Title
                If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
                campoInErrore = True
                
            
            ElseIf mTipoUtenza <> "D" And getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.flgAmmessoPerUA) = False Then
                txt(TXT_QUANTITA).ToolTipText = "Attenzione l'Utenza Attività non è abilitata a portare il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text
                txt(TXT_QUANTITA).Text = 0
                txt(TXT_QUANTITA).BackColor = ColorError
                MsgBox txt(TXT_QUANTITA).ToolTipText, vbExclamation + vbOKOnly, App.Title
                If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
                campoInErrore = True
            
            Else
                MaxGGUD = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxGGUD)
                MaxYYUD = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxYYUD)
     
                MaxGGUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxGGUA)
                MaxYYUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxYYUA)
                
                
                MaxFatturabileUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxFatturabile)
                
                quantita = Round(getNumber(txt(TXT_QUANTITA).Text), 2)
                TotMovYY = getTotaleMovimentiAnnoPerCER(imgCombo1.SelectedItem.key, Mid$(lvwImmobili.Tag, 2))
                TotMovGG = getTotaleMovimentiNelPeriodoPerCER(imgCombo1.SelectedItem.key, Mid$(lvwImmobili.Tag, 2))
                
                If mTipoUtenza = "D" Then
                    'limite quantita annuo UD
                    If MaxYYUD > 0 And (MaxYYUD < TotMovYY + quantita) Then
                        txt(TXT_QUANTITA).ToolTipText = "Attenzione si è superato il limite annuo di:" & MaxYYUD
                        TotResiduo = IIf(TotMovYY >= MaxYYUD, 0, TotMovYY + quantita - MaxYYUD)
                        TotMovGG = MaxGGUD + TotResiduo
                    End If
                    'limite quantita giornaliero UD
                    If MaxGGUD > 0 And (MaxGGUD < TotMovGG + quantita) Then
                        txt(TXT_QUANTITA).ToolTipText = txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Attenzione si è superato il limite " & IIf(gQtaMaxGiornaliera, "giornaliero", "settimanale") & " di:" & MaxGGUD
                        TotResiduo = IIf(TotMovGG >= MaxGGUD, 0, TotMovGG + quantita - MaxGGUD)
                    End If
                Else
                    'limite quantita annuo UA
                    If MaxYYUA > 0 And (MaxYYUA < TotMovYY + quantita) Then
                        txt(TXT_QUANTITA).ToolTipText = "Attenzione si è superato il limite annuo di:" & MaxYYUA
                         TotResiduo = IIf(TotMovYY >= MaxYYUA, 0, TotMovYY + quantita - MaxYYUA)
                        TotMovGG = MaxGGUA + TotResiduo
                    End If
                    
                    'limite quantita giornaliero UA
                    If MaxGGUA > 0 And (MaxGGUA < TotMovGG + quantita) Then
                        txt(TXT_QUANTITA).ToolTipText = txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Attenzione si è superato il limite " & IIf(gQtaMaxGiornaliera, "giornaliero", "settimanale") & " di:" & MaxGGUA
                        TotResiduo = IIf(TotMovGG >= MaxGGUA, 0, TotMovGG + quantita - MaxGGUA)
                    End If
                    
                    'Fatturazione
                    mTotResiduoFatt = 0
                    If getPropietaCER(imgCombo1.SelectedItem.Tag, FlgAbilitatoFatturazione) = True Then
                    
                        If gflgFatturazioneConKD = 0 Then
                            'limite quantita Fatturazione UA
                            If MaxFatturabileUA < TotMovYY + quantita Then
                                 mTotResiduoFatt = IIf(TotMovYY >= MaxFatturabileUA, quantita, TotMovYY + quantita - MaxFatturabileUA)
                            End If
                        Else
                             'limite quantita Fatturazione UA
                            MaxFatturabileUA = lvwImmobili.SelectedItem.ListSubItems("MaxQtaFatt")
                            TotMovYYPerImmobile = getTotaleMovimentiAnnoPerImmobile(Mid$(lvwImmobili.Tag, 2))
                            
                            If MaxFatturabileUA < TotMovYYPerImmobile + quantita Then
                                 mTotResiduoFatt = IIf(TotMovYYPerImmobile >= MaxFatturabileUA, quantita, TotMovYYPerImmobile + quantita - MaxFatturabileUA)
                            End If
                        End If
                        
                        Label1(LBL_FATTURABILE).Visible = mTotResiduoFatt > 0

                    End If
                    
                End If
                mForzaturaMovimento = False
                If txt(TXT_QUANTITA).ToolTipText <> "" Then
                    If (MsgBox(txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Si vuole comunque forzare il movimento?", vbQuestion + vbYesNo, App.Title) = vbNo) Then
                        txt(TXT_QUANTITA).Text = 0
                        If (TotResiduo > 0) Then
                            MsgBox "Occore effettuare una altra pesata togliendo la quantità in eccesso pari a:" & TotResiduo, vbExclamation + vbOKOnly, App.Title
                        Else
                            MsgBox "Non è possibile effetuare altre registrazioni su questo immobile per il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text, vbExclamation + vbOKOnly, App.Title
                        End If
                        
                        txt(TXT_QUANTITA).BackColor = ColorError
                        If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
                        campoInErrore = True
                    Else
                        mForzaturaMovimento = True
                    End If
                End If
            End If
        End If
    End If
    
    
    validaDati = Not campoInErrore
    
    
    'se richiesto ALIBI MEMORY WRITE e sono in peso proveniente da bilancia, mando il messaggio di scrittura
    If validaDati And mPesoLettoDaBilancia Then
        If getCurrentPesa.CmdScritturaAlibiM <> "" Then
            InvioComandoAllaBilancia cmd_ScritturaAlibi, True, idAlibiMemory
            mPesoLettoDaBilancia = False
'            If Not InvioComandoAllaBilancia(cmd_ScritturaAlibi, True, idAlibiMemory) Then
'                validaDati = False
'            Else
'                'resetto la modalità da PESA
'                mPesoLettoDaBilancia = False
'            End If
        End If
    End If
Exit Function
validaDatiERR:
    ScriviLog TPL_ERRORE, App.Title, "validaDati", Err.Description

    
End Function

Private Function getMaxQuantitaPerFatturaImmobile(pCodComune As String, pCodUtente As Long, _
                                                  pIndirizzo As String, pNumCivico As String, pCodTariffa As String) As Double
On Error GoTo getMaxQuantitaPerFatturaImmobileERR
    Dim cmd As New ADODB.Command
    cmd.CommandText = "SELECT Immobili.CodComune,Immobili.CodUtente, Immobili.Indirizzo, Immobili.NumCivico, Immobili.CodTariffa, Sum([Superficie]*[tariffe].[kd]) AS SommaSup " & _
                      "FROM Immobili INNER JOIN Tariffe ON Immobili.CodTariffa = Tariffe.CodTariffa " & _
                      "Where Immobili.TipoUtenza='N' and DATE() between DataInizioValidita and DataFineValidita and immobili.Tipo<>'Z' " & _
                      "GROUP BY  Immobili.CodComune,Immobili.CodUtente, Immobili.Indirizzo, Immobili.NumCivico, Immobili.CodTariffa " & _
                      "HAVING  Immobili.CodComune='" & pCodComune & "'" & _
                      " and Immobili.CodUtente=" & pCodUtente & _
                      " and Immobili.Indirizzo='" & Replace(pIndirizzo, "'", "''") & "'" & _
                      " and Immobili.NumCivico='" & pNumCivico & "'" & _
                      " and Immobili.CodTariffa='" & pCodTariffa & "'"

    cmd.ActiveConnection = gDBC
    
    Dim rs  As ADODB.Recordset
    Set rs = cmd.Execute
    
    If Not rs.EOF Then
        getMaxQuantitaPerFatturaImmobile = rs("SommaSup").value
    Else
        getMaxQuantitaPerFatturaImmobile = -1
    End If
    
    
    rs.Close
    Set rs = Nothing
        
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
Exit Function
getMaxQuantitaPerFatturaImmobileERR:
      ScriviLog TPL_ERRORE, App.Title, "getMaxQuantitaPerFatturaImmobileERR", Err.Description
End Function

Private Function salvaPesata(forzatura As Boolean, idMovimento As Long, idAlibiMemory As String) As Boolean
    On Error GoTo salvaPesataERR
    Dim codImmobile As Long
    Dim codUtente As Long
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codCer As String
    Dim DataMovimento As String
    Dim quantita As Double
    Dim numeroPezzi As Double
    Dim BonusDovuto As Double
    Dim PremioDovuto As Double
    Dim nrMov As Long
    Dim idRegistro As Long
    Dim nrRegistro As Long
    Dim manuale As Boolean
    Dim registrato As Boolean
    Dim codSubCer As String
    
    On Error GoTo salvaPesataERR
    
    gDBC.BeginTrans
    
    
    nrMov = getNumeratore(enumEntita.MOVIMENTI, False)
    salvaPesata = False
    codImmobile = lvwImmobili.SelectedItem.Text
    codUtente = txt(TXT_COD_UTENTE).Text
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codSubCer = Right$(codCer, 2)
    codCer = Mid$(codCer, 5, 2)
    
    DataMovimento = IIf(CLng(msgPesa.Data) <= 0, Now, msgPesa.Data)
    quantita = Round(getNumber(txt(TXT_QUANTITA).Text), 2)
    numeroPezzi = 0
    
    Dim bonusAttivo As Boolean
    Dim premioAttivo As Boolean
    
    BounsAttivoEPremioAttivo codUtente, bonusAttivo, premioAttivo
    
    If bonusAttivo Then
        BonusDovuto = getNumber(getPropietaCER(imgCombo1.SelectedItem.Tag, CorrispettivoBonus))
        BonusDovuto = Round(BonusDovuto * quantita, 4)
    End If
    
    If premioAttivo Then
        PremioDovuto = getNumber(getPropietaCER(imgCombo1.SelectedItem.Tag, CorrispettivoBonusPunti))
        PremioDovuto = Round(PremioDovuto * quantita, 4)
    End If
    
    manuale = IIf(CLng(msgPesa.Data) <= 0, True, False)
    registrato = IIf(txt(TXT_NRFORMULARIO).Text <> "", True, False)
    
    'salva documento e riferimento al movimento
    If mnrDocumentoPesata = 0 Then
        mnrDocumentoPesata = getNumeratore(enumEntita.DOCUMENTO_PESATA_IN, False)
        'aggiorna riferimento sulla tabella documenti
        AggiornaRiferimentoInDocumento midDocumentoPesata, mnrDocumentoPesata, midMovimentiDaStampare, "I"
    End If
    
    
    idMovimento = creaMovimento(nrMov, codImmobile, codUtente, GetCurrentComune!codistat, codCategoria, codSottoCategoria, codCer, _
      codSubCer, DataMovimento, quantita, "0", manuale, frmLogin.codUser, _
        BonusDovuto, PremioDovuto, getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura), _
        cmbTarga(0).Text, txt(TXT_NRFORMULARIO).Text, Format$(dtpickFormulario.value, "YYYY-MM-DD"), _
         forzatura, cmbTarga(1).Text, CLng(IIf(txt(TXT_COD_TRASPORTATORE).Text = "", -1, _
         txt(TXT_COD_TRASPORTATORE).Text)), "00", registrato, gPesaIsolaCod, midDocumentoPesata, idAlibiMemory)


    salvaTarga txt(TXT_COD_UTENTE).Text, cmbTarga(0).Text, "U"

    'crea record per fatturazione
    If mTotResiduoFatt > 0 Then
        CreaRecordPerFatturazione mTotResiduoFatt, DataMovimento, codCategoria & codSottoCategoria & codCer & codSubCer, codImmobile, codUtente, txt(TXT_NRFORMULARIO).Text, idMovimento
    End If

    
    If txt(TXT_NRFORMULARIO).Text <> "" Then
    
        nrRegistro = getNumeratore(REGISTRO, False)
        
        idRegistro = creaVoceDiRegistro(nrRegistro, "C", Format$(DataMovimento, "yyyy/MM/dd"), txt(TXT_NRFORMULARIO).Text, Format$(dtpickFormulario.value, "YYYY-MM-DD") _
            , codCategoria, codSottoCategoria, codCer, "", getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura), _
                quantita, "", 0, "", Now, frmLogin.codUser, "")
    
    
        creaLotto idRegistro, codCategoria, codSottoCategoria, codCer, _
                  quantita, 0, "I", 0, Now, frmLogin.codUser, Now, frmLogin.codUser
    
    End If
    
   
    gDBC.CommitTrans
    
    salvaPesata = True
    Exit Function
salvaPesataERR:
    ScriviLog TPL_ERRORE, App.Title, "salvaPesata", Err.Description

    gDBC.RollbackTrans
    
End Function



Private Sub CreaRecordPerFatturazione(quantita As Double, DataMovimento As String, codCer As String, codImmobile As Long, codUtente As Long, nrFormulario As String, idRifMovimento As Long)
    Dim cmd As ADODB.Command
    Dim P As Parameter
    
    On Error GoTo CreaRecordPerFatturazioneERR
    
    'se utenza Non domestica verifico se il totale delle pesate annuali supera il max
    'se si viene fatturata la differenza
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO Fatturazione (CODCOMUNE, PROGRESSIVO_INVIO, COD_ANA,  " & _
            "COD_IMMOBILE, COD_RIFIUTO, QUANTITA, DATA_MOVIMENTO, NUMERO_FORMULARIO, STATO_INVIO, DATA_SCARICO, FLAG, [NOTE],idRifMovimento) " & _
            "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?)"
            
            
    cmd.Parameters.Append cmd.CreateParameter("CODCOMUNE", adVarChar, adParamInput, 6)
    cmd.Parameters.Append cmd.CreateParameter("PROGRESSIVO_INVIO", adInteger, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("COD_ANA", adInteger, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("COD_IMMOBILE", adInteger, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("COD_RIFIUTO", adVarChar, adParamInput, 8)
    
    
    Set P = cmd.CreateParameter("QUANTITA", adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6
    cmd.Parameters.Append P
    
    cmd.Parameters.Append cmd.CreateParameter("DATA_MOVIMENTO", adDate, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("NUMERO_FORMULARIO", adVarChar, adParamInput, 15)
    cmd.Parameters.Append cmd.CreateParameter("STATO_INVIO", adVarChar, adParamInput, 1)
    cmd.Parameters.Append cmd.CreateParameter("DATA_SCARICO", adDate, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("FLAG", adVarChar, adParamInput, 1)
    cmd.Parameters.Append cmd.CreateParameter("NOTE", adVarChar, adParamInput, 50)
    cmd.Parameters.Append cmd.CreateParameter("idRifMovimento", adInteger, adParamInput)
    
    Set cmd.ActiveConnection = gDBC
   
    
    cmd.Parameters("CODCOMUNE").value = GetCurrentComune!codistat
    cmd.Parameters("PROGRESSIVO_INVIO").value = getNumeratore(FATTURE, False)
    cmd.Parameters("COD_ANA").value = codUtente
    cmd.Parameters("COD_IMMOBILE").value = codImmobile
    cmd.Parameters("COD_RIFIUTO").value = codCer
    
    cmd.Parameters("QUANTITA").value = quantita
    cmd.Parameters("DATA_MOVIMENTO").value = DataMovimento
    cmd.Parameters("NUMERO_FORMULARIO").value = nrFormulario
    cmd.Parameters("STATO_INVIO").value = 5
    cmd.Parameters("DATA_SCARICO").value = Null
    cmd.Parameters("FLAG").value = Null
    cmd.Parameters("NOTE").value = Null
    cmd.Parameters("idRifMovimento").value = idRifMovimento
    
    
    cmd.Execute
    Exit Sub
CreaRecordPerFatturazioneERR:
    ScriviLog TPL_ERRORE, App.Title, "CreaRecordPerFatturazione", Err.Description

End Sub


Private Sub cmbRicerca_Click(Index As Integer)
On Error GoTo cmbRicerca_ClickERR

    Select Case Index
        Case 0
            frmRicercaUtente.schowMe txt(TXT_COD_UTENTE), "U"
            caricaDatiUtente txt(TXT_COD_UTENTE).Text
        Case 1
            frmRicercaUtente.schowMe txt(TXT_COD_TRASPORTATORE), "T"
            caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE).Text
        Case 2
            frmRicercaUtente.schowMe txtBadgeAccessi, "U", True
            If Trim$(txtBadgeAccessi.Text) <> vbNullString Then
               PlexaNetKX1_OnCodeEnter 2, 100, 1, txtBadgeAccessi.Text
               txtBadgeAccessi.Text = vbNullString
            End If
    End Select
    Exit Sub
cmbRicerca_ClickERR:
    ScriviLog TPL_ERRORE, App.Title, "cmbRicerca_Click", Err.Description

End Sub



 
Private Sub caricaDatiTrasportatore(CodTrasportatore As String)
    Dim rs As Recordset
    On Error GoTo caricaDatiTrasportatoreERR
    
    
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
    
    If CodTrasportatore <> vbNullString Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.LockType = adLockBatchOptimistic
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Source = "select *,comuni.comune from Trasportatori LEFT JOIN Comuni ON Trasportatori.CodComuneSedeLegale = Comuni.CODISTAT where idTrasportatore=" & CodTrasportatore
        
        Set rs.ActiveConnection = gDBC
        rs.Open
        Set rs.ActiveConnection = Nothing
                
        If rs.EOF Then
            MsgBox "caricaDatiTrasportatore:" & "Attenzione Trasportatore:" & CodTrasportatore & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
        Else
            
            Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = rs!ragioneSociale
            If nullToBlank(rs!CodFiscale) = vbNullString Then
                Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption & "p.Iva " & nullToBlank(rs!PartitaIva)
            Else
                Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption & "cod.Fisc " & nullToBlank(rs!CodFiscale)
            End If
            
            DoEvents
            
            Set rsCliente = rs.Clone
            
        End If
        rs.Close
    End If
    
    'carica combo targhe
    cmbTarga(1).Clear
    
    If CodTrasportatore <> "" Then
        If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
            DevEnvMDB.rscmdAutomezzi.Close
        End If
        DevEnvMDB.cmdAutomezzi CodTrasportatore, "T"
        Do While Not DevEnvMDB.rscmdAutomezzi.EOF
            cmbTarga(1).AddItem DevEnvMDB.rscmdAutomezzi!Targa
            DevEnvMDB.rscmdAutomezzi.MoveNext
        Loop
        If cmbTarga(1).ListCount > 0 Then cmbTarga(1).ListIndex = 0
    End If
    
    Set rs = Nothing
    Exit Sub
caricaDatiTrasportatoreERR:
    ScriviLog TPL_ERRORE, App.Title, "caricaDatiTrasportatore", Err.Description

End Sub


Private Sub caricaDatiUtente(ByVal codUtente As String)
    Dim rs As Recordset
On Error GoTo caricaDatiUtenteERR
    
    Label3(1).Caption = 0
    
    txtBadge.Text = ""
    txtInfoUser.Text = ""
    
        
    m_badge = vbNullString
    m_pivaCodFisc = vbNullString
    m_pivaCodFisc = vbNullString
    m_badgeDTini = vbNullString
    m_badgeDTend = vbNullString
    m_badgeProg = vbNullString
    
    mTipoUtenza = ""

    If codUtente = vbNullString Then
    
        Exit Sub
    End If
    
    If (mlastUtente <> -1 And mlastUtente <> codUtente) And midMovimentiDaStampare <> vbNullString Then
        MsgBox "Attenzione L'utente è cambiato, verrà ora avviata la stampa movimenti", vbExclamation + vbOKOnly, App.Title
        StampaMovimenti False, codUtente
    End If
    
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Source = "select utenti.*,codBadge,DataBadgeInizio,DataBadgefine,ProgressivoBadge,comuni.comune, " & _
                "(badge.DataBadgeFine - date()) as ScadenzaGG from ((Utenti  LEFT JOIN Comuni ON Utenti.CodComuneResidenza = Comuni.CODISTAT) " & _
                "inner join badge on badge.codUtente = utenti.codutente and badge.codcomune = utenti.codcomune) where badge.default=true and utenti.codUtente=" & codUtente
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
            
    If rs.EOF Then
        MsgBox "Attenzione CodUtente:" & codUtente & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
    Else
        
        
        
        m_ragioneSociale = IIf(nullToBlank(rs!ragioneSociale) = vbNullString, nullToBlank(rs!Nome) & " " & nullToBlank(rs!Cognome), nullToBlank(rs!ragioneSociale))
        m_codComune = fCapitalize(nullToBlank(rs!Comune.value))
        m_indirizzo = fCapitalize(nullToBlank(rs!IndirizzoResidenza) & ", " & nullToBlank(rs!NumCivicoResidenza))
        
        
        m_badge = nullToBlank(rs!codBadge)
        m_pivaCodFisc = nullToBlank(rs!CodFiscale)
        If m_pivaCodFisc = vbNullString Then m_pivaCodFisc = nullToBlank(rs!PartitaIva)
        m_badgeDTini = nullToBlank(rs!dataBadgeInizio)
        m_badgeDTend = nullToBlank(rs!DataBadgeFine)
        m_badgeProg = nullToBlank(rs!progressivobadge)
        
        Set rsCliente = rs.Clone
                
        txtInfoUser.Text = m_ragioneSociale & vbCrLf & m_indirizzo & IIf(m_codComune = vbNullString, "", " ," & m_codComune) & " ," & _
                       IIf(nullToBlank(rs!CodFiscale) = vbNullString, "p.Iva ", "cod.Fisc ") & m_pivaCodFisc & " badge prog." & m_badgeProg
                
    End If
    
    
    txtBadge.Text = m_badge
    txtBadge.ToolTipText = "Valido dal " & m_badgeDTini & " al " & m_badgeDTend
    rs.Close
    
    caricaImmobili codUtente
    
    DoEvents
    
    'carica combo targhe
    cmbTarga(0).Clear
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    
    DevEnvMDB.cmdAutomezzi codUtente, "U"
    Do While Not DevEnvMDB.rscmdAutomezzi.EOF
        cmbTarga(0).AddItem DevEnvMDB.rscmdAutomezzi!Targa
        DevEnvMDB.rscmdAutomezzi.MoveNext
    Loop
    
    
    mlastUtente = codUtente
    mlastBadge = m_badge
    
    If cmbTarga(0).ListCount > 0 Then cmbTarga(0).ListIndex = 0
    Exit Sub
caricaDatiUtenteERR:
    ScriviLog TPL_ERRORE, App.Title, "caricaDatiUtente", Err.Description

End Sub

Private Sub getTotaliPerCodiceCer(codImmobile As Long)
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim UnitaMisura As String
    Dim totQTA As Double
    Dim totBONUS As Double
    Dim totPUNTI As Double
    
    
    On Error GoTo getTotaliPerCodiceCerERR

    
    lvwTotMovimenti.ListItems.Clear
    
    If codImmobile = 0 Then Exit Sub
    
    Set rs = gDBC.Execute("SELECT Movimenti.Stato, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer,Movimenti.CodSubCer,Movimenti.UnitaMisura, CodiciCerComunali.Descrizione, Sum(Movimenti.BonusDovuto) AS Bonus,Sum(Movimenti.PuntiDovuti) AS Punti," & _
        "Sum(Movimenti.Quantita) AS Quantita, " & _
        "IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxGGUD],[CodiciCer].[qtyMaxGGUA]) AS qtyMaxGG, IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxAAUD],[CodiciCer].[qtyMaxAAUA]) AS qtyMaxAA " & _
        "FROM ((Movimenti " & _
        "INNER JOIN CodiciCerComunali ON (Movimenti.CodCategoria = CodiciCerComunali.CodCategoria) AND (Movimenti.CodSubCer = CodiciCerComunali.CodSubCer) AND (Movimenti.CodComune = CodiciCerComunali.CodComune) AND (Movimenti.CodCer = CodiciCerComunali.CodCer) AND (Movimenti.CodSottoCategoria = CodiciCerComunali.CodSottoCategoria)) " & _
        "INNER JOIN CodiciCer ON (CodiciCer.CodCategoria = CodiciCerComunali.CodCategoria) AND  (CodiciCer.CodCer = CodiciCerComunali.CodCer) AND (CodiciCer.CodSottoCategoria = CodiciCerComunali.CodSottoCategoria)) " & _
        "INNER JOIN Immobili ON Movimenti.CodImmobile = Immobili.CodImmobile and Movimenti.CodUtente = Immobili.CodUtente " & _
        "WHERE Movimenti.codImmobile = " & codImmobile & " and  Movimenti.DataMovimento Between dateadd('d', datediff('d',  date()  ,year(date()) &'-01-01' ) , date()) +'00.00.00' And Now() AND Movimenti.Stato='0' and (codCausale='00'" & _
        " or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
        "GROUP BY Movimenti.Stato, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer, Movimenti.CodsubCer,Movimenti.UnitaMisura, CodiciCerComunali.Descrizione, IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxGGUD],[CodiciCer].[qtyMaxGGUA]), IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxAAUD],[CodiciCer].[qtyMaxAAUA]), FlgBonus, flgBonusAPunti " & _
        "ORDER BY FlgBonus,flgBonusAPunti ")
    
    Set item = lvwTotMovimenti.ListItems.Add(, , "")
    item.ForeColor = vbRed
    
    item.ListSubItems.Add , , "Tot.Quantità"
    item.ListSubItems(item.ListSubItems.Count).ForeColor = vbRed
    item.ListSubItems.Add , , "Tot.Bonus (€)"
    item.ListSubItems(item.ListSubItems.Count).ForeColor = vbRed
    item.ListSubItems.Add , , "Tot.Punti"
    item.ListSubItems(item.ListSubItems.Count).ForeColor = vbRed
    'item.ListSubItems.Add , , "Max al Giorno"
    'item.ListSubItems(item.ListSubItems.Count).ForeColor = vbRed
    'item.ListSubItems.Add , , "Max al Anno"
    'item.ListSubItems(item.ListSubItems.Count).ForeColor = vbRed
    
    totQTA = 0
    totBONUS = 0
    Dim subitem As ListSubItem
    
    Do While Not rs.EOF
        Set item = lvwTotMovimenti.ListItems.Add(, "_" & rs!codCategoria & rs!codSottoCategoria & rs!codCer & rs!codSubCer & rs!UnitaMisura, rs!Descrizione & " (" & rs!UnitaMisura & ")")
        
        totQTA = totQTA + rs!quantita
        totBONUS = totBONUS + rs!Bonus
        totPUNTI = totPUNTI + rs!Punti
        
        item.ListSubItems.Add , , rs!quantita
        item.ListSubItems.Add , , rs!Bonus
        item.ListSubItems.Add , , rs!Punti
        'item.ListSubItems.Add , , nullToBlank(rs!qtyMaxGG)
        'item.ListSubItems.Add , , nullToBlank(rs!qtyMaxAA)
        rs.MoveNext
    Loop
    
    Label3(1).Caption = FormatNumber(totQTA, 2, vbUseDefault, vbUseDefault, vbTrue) & " Kg"
'    Label3(3).Caption = FormatNumber(totBONUS, 2, vbUseDefault, vbUseDefault, vbTrue) & " €"
'    Label3(4).Caption = FormatNumber(totPUNTI, 2, vbUseDefault, vbUseDefault, vbTrue)
        
    
    rs.Close
    
    AllargaLeColonneLVW lvwTotMovimenti, False
    Exit Sub
getTotaliPerCodiceCerERR:
        ScriviLog TPL_ERRORE, App.Title, "getTotaliPerCodiceCer", Err.Description

End Sub

Function fCapitalize(Stringa As String) As String
    Dim s() As String
    Dim i  As Long
    
    s = Split(Stringa, " ")
    For i = LBound(s) To UBound(s)
        s(i) = Trim$(s(i))
        If s(i) <> vbNullString And Len(s(i)) > 1 Then
            Mid(s(i), 1, 1) = UCase$(Mid$(s(i), 1, 1))
            Mid(s(i), 2) = LCase$(Mid$(s(i), 2))
        End If
    Next
        
    fCapitalize = Join(s, " ")
End Function

Private Sub cmdRicercaUtenteMov_Click(Index As Integer)
    Dim utente As TextBox
    Dim codUte As String
    
    On Error GoTo cmdRicercaUtenteMov_ClickERR
    
    
    If Index = 0 Then
        Set utente = TxtMov(0)
        codUte = "U"
    Else
        Set utente = TxtMov(2)
        codUte = "T"
    End If
    
    If (cmdRicercaUtenteMov(0).Tag = "") Then
        frmRicercaUtente.schowMe utente, codUte
    End If
    cmdRicercaUtenteMov(0).Tag = ""
   
    'carica combo targhe
    cmbTargaMov(Index).Clear
    If (utente = "") Then Exit Sub
    
    cmbTargaMov(Index).AddItem "Tutte"
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    
    DevEnvMDB.cmdAutomezzi utente, codUte
    Do While Not DevEnvMDB.rscmdAutomezzi.EOF
        cmbTargaMov(Index).AddItem DevEnvMDB.rscmdAutomezzi!Targa
        DevEnvMDB.rscmdAutomezzi.MoveNext
    Loop

    cmbTargaMov(Index).ListIndex = 0
    Exit Sub
cmdRicercaUtenteMov_ClickERR:
    ScriviLog TPL_ERRORE, App.Title, "cmdRicercaUtenteMov_Click", Err.Description

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Not getCurrentPesa.BadgeInStringaPesata Then txtBadge.SetFocus
End Sub
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As ADODB.Recordset

On Error GoTo Form_KeyDownERR

    If KeyCode = vbKeyF2 Then
        Dim codUtente As Integer
        codUtente = CInt(Int((15740 * Rnd()) + 10000))
        Set rs = gDBC.Execute("Select CodFiscale,PartitaIva from utenti where codutente=" & codUtente)
        If Not rs.EOF Then
            PlexaNetKX1_OnCodeEnter 2, 100, 1, IIf(nullToBlank(rs!CodFiscale) = vbNullString, nullToBlank(rs!PartitaIva), nullToBlank(rs!CodFiscale))
        End If
        rs.Close
    ElseIf KeyCode = vbKeyF3 Then 'fotocellula
        PlexaNetKX1_OnDigitalLevelRead 1, 1, 0
    ElseIf KeyCode = vbKeyF4 Then 'uscita
        PlexaNetKX1_OnDigitalLevelRead 1, 2, 0
    End If
    Exit Sub
Form_KeyDownERR:
    ScriviLog TPL_ERRORE, App.Title, "Form_KeyDown", Err.Description

End Sub

Private Sub imgCombo1_Click()
On Error GoTo imgCombo1_ClickERR


    
    'se viene selezionato un codice CER imposto unità di misura
    If Not imgCombo1.SelectedItem Is Nothing Then
        If (Mid$(imgCombo1.SelectedItem.key, 5, 2) <> "00") And Len(imgCombo1.SelectedItem.key) > 6 Then
            Label1(LBL_QUANTITA) = "Quantità (" & getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura) & ")"
            
            'Label1(LBL_FATTURABILE).Visible = mTipoUtenza = "N" And CBool(getPropietaCER(imgCombo1.SelectedItem.Tag, FlgAbilitatoFatturazione))
            'Label1(LBL_FATTURABILE).Visible = False
            
            ShowMessage InfoCodCer, imgCombo1.SelectedItem.Text

        End If
    End If
Exit Sub
imgCombo1_ClickERR:
    ScriviLog TPL_ERRORE, App.Title, "imgCombo1_Click", Err.Description

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    
    imgTrafficLightGreen.Visible = True
    imgTrafficLightRed.Visible = False
    
    mlastUtente = -1
    mlastBadge = vbNullString
    txtBadge.Locked = getCurrentPesa.BadgeInStringaPesata
       ShowMessage InfoCodCer, vbNullString
    
    
    ButtonsRicercaMov = CMD_ESCI & "," & CMD_RISTAMPAMOV & "," & CMD_ESPORTAMOV & "," & CMD_CANCELLAMOV & "," & CMD_RICERCAMOV & ","
    ButtonsRicercaacc = CMD_ESCI & "," & CMD_RICERCAMOV & ","
    
    txtBadge.MaxLength = gLunghezzaBadge
    txtBadgeAccessi.MaxLength = 16
    txtBadgeAccessi.Text = ""
    
    Label1(LBL_FATTURABILE).Visible = False
    lblEsportazione.Visible = False
    ProgressBar1.Visible = False
    applicaRestrizioniUtente
    
    For i = picContainer.LBound To picContainer.UBound
        picContainer(i).Visible = False
    Next i

    caricaComboRifiuti imgCombo1, Combo2(0), True
    caricaComboRifiuti ImgCombo2, Combo2(1), False
    
    Set rs = CreateObject("ADODB.Recordset")
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from causali"
    rs.Open
    Do While Not rs.EOF
        cmbRicercaMov(5).AddItem rs!codCausale & " - " & rs!Descrizione
        rs.MoveNext
    Loop
    
    CaricaListaComuni
    
    dtpickFormulario.value = Now
    
    cmbRicercaMov(0).ListIndex = 1
    cmbRicercaMov(1).ListIndex = 0
    cmbRicercaMov(2).ListIndex = 0
    cmbRicercaMov(3).ListIndex = 0
    cmbRicercaMov(5).ListIndex = 0
    
    cmbTarga(0).Text = ""
    'cmbTarga(1).Text = ""
    cmbTargaMov(0).AddItem "Tutte"
    cmbTargaMov(0).ListIndex = 0
    cmbTargaMov(1).AddItem "Tutte"
    cmbTargaMov(1).ListIndex = 0
    
    Set rs = gDBC.Execute("select * from tariffe order by codTariffa")
    Do While Not rs.EOF
        cmbRicercaMov(2).AddItem (rs!codTariffa & " - " & rs!Descrizione)
        rs.MoveNext
    Loop
    
    
    txtBadge.Text = ""
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
    
    lvwImmobili.ColumnHeaders.Add , "cod", "Cod.Imm."
    lvwImmobili.ColumnHeaders.Add , , "Tipo"
    lvwImmobili.ColumnHeaders(lvwImmobili.ColumnHeaders.Count).width = 0
    lvwImmobili.ColumnHeaders.Add , , "Comune"
    lvwImmobili.ColumnHeaders.Add , "Tipo Utenza", "Tpo Ute"
    lvwImmobili.ColumnHeaders.Add , "MaxQuantFatt", "Max Quant.Fatt"
    lvwImmobili.ColumnHeaders.Add , , "Indirizzo"
    'lvwImmobili.ColumnHeaders.Add , , "Sub.Alt."
    'lvwImmobili.ColumnHeaders.Add , , "Nr.Int."
    lvwImmobili.ColumnHeaders.Add , , "Cod.Tar."
    lvwImmobili.ColumnHeaders.Add , , "Tariffa"
    
    
    lvwAccessi.ColumnHeaders.Clear
    lvwAccessi.ColumnHeaders.Add , , "Utente"
    lvwAccessi.ColumnHeaders.Add , , "Badge"
    lvwAccessi.ColumnHeaders.Add , , "Cod.Fiscale/PIVA"
    lvwAccessi.ColumnHeaders.Add , , "Data Ora Ingresso"
    lvwAccessi.ColumnHeaders.Add , , "Data Ora Uscita"
    lvwAccessi.ColumnHeaders.Add , , "Ospite"
    lvwAccessi.ColumnHeaders.Add , , "UND"
    lvwAccessi.ColumnHeaders.Add , , "Pesata"
    lvwAccessi.ColumnHeaders.Add , , "Stato"
    
    lvwMovimenti.ColumnHeaders.Clear
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Utente"
    lvwMovimenti.ColumnHeaders.Add , , "Utente"
    lvwMovimenti.ColumnHeaders.Add , , "Badge"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Imm."
    lvwMovimenti.ColumnHeaders.Add , , "Immobile"
    lvwMovimenti.ColumnHeaders.Add , , "Tip.Ute."
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Tar."
    lvwMovimenti.ColumnHeaders.Add , , "Targa"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Cer"
    lvwMovimenti.ColumnHeaders.Add , , "Tipo Rif."
    lvwMovimenti.ColumnHeaders.Add , , "Nr. Mov."
    lvwMovimenti.ColumnHeaders.Add , , "Data Mov."
    lvwMovimenti.ColumnHeaders.Add , , "Quantità"
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "U.M."
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "Bonus €"
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "Punti"
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , "Causale", "Causale"
    lvwMovimenti.ColumnHeaders.Add , , "Id.Ali.Mem"
    lvwMovimenti.ColumnHeaders.Add , , "Reg."
    lvwMovimenti.ColumnHeaders.Add , , "Stato"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Trasp."
    lvwMovimenti.ColumnHeaders.Add , , "Trasp."
    lvwMovimenti.ColumnHeaders.Add , , "Targa Trasp."
    lvwMovimenti.ColumnHeaders.Add , "NumeroFormulario", "Nr.Formulario"
    lvwMovimenti.ColumnHeaders.Add , , "Data Formulario"
    lvwMovimenti.ColumnHeaders.Add , , "Tipo"
    lvwMovimenti.ColumnHeaders.Add , , "Utente Ins."
    lvwMovimenti.ColumnHeaders.Add , , "Data Ins."

    
    
    DTPicker(0).value = Year(Now) & "/" & Right$(0 & (Month(Now)), 2) & "/" & 1
    '
    DTPicker(1).value = DateAdd("d", -Day(DateAdd("m", 1, Now)), DateAdd("m", 1, Now))
        
    cmbRicercaMov_Change 0
    
    cmbRicercaMov(4).Clear
    cmbRicercaMov(4).AddItem ("Tutti")
    cmbRicercaMov(4).AddItem ("Da Registrare")
    cmbRicercaMov(4).AddItem ("Registrati")
    cmbRicercaMov(4).ListIndex = 0


    If gPuntiAssegnatiPer > 0 Then
        cmbRicercaMov(3).AddItem GetTipiAnagrafe
    End If
    
    lblCambiaPesa.Visible = getCountPese > 1

    frmAccessi.Visible = gActiveAccess = 1
    tmrAccessi.enabled = gActiveAccess = 1
    
    'comando contatore accessi all'isola
    If gActiveAccess Then
        
        MDIMain.SbarraCollegata False, "Sbarra"
        
        tmrAccessi.Interval = 1000 * gSecRefreshAccessMonitor
        
        Set rsTotaleAccessi = CreateObject("ADODB.Recordset")
        rsTotaleAccessi.LockType = adLockBatchOptimistic
        rsTotaleAccessi.CursorLocation = adUseClient
        rsTotaleAccessi.CursorType = adOpenStatic
        Set rsTotaleAccessi.ActiveConnection = gDBC
        
        rsTotaleAccessi.Source = "Select count(*) from RilevazioneAccessi where DataOraUscita is null"
        rsTotaleAccessi.Open
        
        
        Set rsTotaleDaPesare = CreateObject("ADODB.Recordset")
        rsTotaleDaPesare.LockType = adLockBatchOptimistic
        rsTotaleDaPesare.CursorLocation = adUseClient
        rsTotaleDaPesare.CursorType = adOpenStatic
        Set rsTotaleDaPesare.ActiveConnection = gDBC
        
        rsTotaleDaPesare.Source = "Select count(*) from RilevazioneAccessi where DataOraUscita is null and DaPesare=true"
        rsTotaleDaPesare.Open
        
        TrafficLight Rosso
        
        'connessione al rilevatore accessi
        PlexaNetKX1.Port = gAccessDeviceCOMName
        PlexaNetKX1.Connect
    End If
    
End Sub

Private Sub CaricaListaComuni()
    Dim Command As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Integer

On Error GoTo CaricaListaComuniERR


    cmbRicercaMov(6).Clear

    Set Command = CreateObject("ADODB.Command")
    Set Command.ActiveConnection = gDBC
    
    Command.CommandText = "select * from comuni where logo<>''"
    Set rs = Command.Execute
    
    cmbRicercaMov(6).AddItem "000000 - Tutti"
    cmbRicercaMov(6).ListIndex = 0
    i = 0
    Do While Not rs.EOF
        i = i + 1
        cmbRicercaMov(6).AddItem rs.Fields("CodIstat") & " - " & rs.Fields("Comune")
        If InStr(1, cmbRicercaMov(6).List(cmbRicercaMov(6).ListCount - 1), gCODISTAT) > 0 Then
             cmbRicercaMov(6).ListIndex = i
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
Exit Sub
CaricaListaComuniERR:
    ScriviLog TPL_ERRORE, App.Title, "CaricaListaComuni", Err.Description

End Sub

Private Sub CaricaMovimenti()
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim where As String
    Dim codCer As String
    Dim codSottoCategoria As String
    Dim codCategoria As String
    Dim codSubCer As String
    Dim qty As Double

On Error GoTo CaricaMovimentiERR

    
    where = ""
    Screen.MousePointer = vbHourglass
    
    If TxtMov(0).Text <> "" Then
        If where <> "" Then where = where & " and "
        where = "Immobili.CodUtente =" & TxtMov(0)
    End If
    
    If TxtMov(2).Text <> "" Then
        If where <> "" Then where = where & " and "
        where = "Movimenti.CodTrasportatore =" & TxtMov(2)
    End If
    
    
    If cmbRicercaMov(0).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
            
        where = where & "Movimenti.Stato ='" & cmbRicercaMov(0).ItemData(cmbRicercaMov(0).ListIndex) & "'"
    ElseIf (chkEscludiCancellati.enabled And chkEscludiCancellati.value = 1) Then
        If where <> "" Then where = where & " and "
        
        where = where & " Movimenti.Stato <>'2'"
    End If
    
    If (chkIngombranti.value = 1) Then
        If where <> "" Then where = where & " and "
        
        where = where & " CodiciCerComunali.flgIngombrante =true"
    End If
    
    
    If cmbRicercaMov(5).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.codCausale ='" & left$(cmbRicercaMov(5).Text, 2) & "'"
    End If
    
    If cmbRicercaMov(1).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        
        If cmbRicercaMov(1).ListIndex = 1 Then
            where = where & "CodiciCerComunali.FlgBonus =true"
        ElseIf cmbRicercaMov(1).ListIndex = 2 Then
            where = where & "CodiciCerComunali.flgBonusAPunti =true"
        ElseIf cmbRicercaMov(1).ListIndex = 3 Then
            where = where & "CodiciCerComunali.flgBonusAPunti =true and CodiciCerComunali.FlgBonus =true"
        Else
            where = where & "CodiciCerComunali.flgBonusAPunti =false and CodiciCerComunali.FlgBonus =false"
        End If
    End If
    
    If cmbRicercaMov(3).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Utenti.TipoAnagrafe ='" & left$(cmbRicercaMov(3).Text, 1) & "'"
    End If
    
    If cmbRicercaMov(2).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Immobili.CodTariffa ='" & left$(cmbRicercaMov(2).Text, 5) & "'"
    End If
    
    If cmbRicercaMov(4).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.flgRegistrato =" & IIf(cmbRicercaMov(4).ListIndex = 1, 0, -1)
    End If
    
    If cmbRicercaMov(6).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.codComune ='" & left$(cmbRicercaMov(6).Text, 6) & "'"
    End If

    
    If Not ImgCombo2.SelectedItem Is Nothing Then
        If (Mid$(ImgCombo2.SelectedItem.key, 5, 2) <> "00") And Len(ImgCombo2.SelectedItem.key) >= 6 Then
            
            codCer = Mid$(ImgCombo2.SelectedItem.key, 3)
            codCategoria = left$(codCer, 2)
            codSottoCategoria = Mid$(codCer, 3, 2)
            codSubCer = Right$(codCer, 2)
            codCer = Mid$(codCer, 5, 2)
            
            
            If where <> "" Then where = where & " and "
            where = where & "Movimenti.CodCer ='" & codCer & "' and Movimenti.CodSottoCategoria='" & codSottoCategoria & "' and Movimenti.CodCategoria='" & codCategoria & "'"
        End If
    End If
    
    If cmbTargaMov(0).Text <> "Tutte" Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.Targa like '" & cmbTargaMov(0).Text & "%'"
    End If
    
    If cmbTargaMov(1).Text <> "Tutte" Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.TargaTrasportatore like '" & cmbTargaMov(1).Text & "%'"
    End If
    
    
    If Not IsNull(DTPicker(0).value) Then
        If where <> "" Then where = where & " and "
        where = where & "dataMovimento >=#" & Format$(DTPicker(0).value, "yyyy-mm-dd") & " 00.00.00#"
    End If
    
    If Not IsNull(DTPicker(1).value) Then
        If where <> "" Then where = where & " and "
        where = where & " dataMovimento <=#" & Format$(DTPicker(1).value, "yyyy-mm-dd") & " 23.59.59#"
    End If
    
    
    
    Set rs = gDBC.Execute("SELECT trasportatori.RagioneSociale AS RagioneSocialeT, badge.codBadge, Utenti.Cognome, Utenti.Nome, Movimenti.CodImmobile, " & _
                            "Utenti.RagioneSociale, Movimenti.id, Movimenti.nrMovimento, Comuni.Comune, Immobili.Indirizzo, Immobili.NumCivico, " & _
                            "Immobili.Subalterno, Immobili.NumInterno, Immobili.Scala, Immobili.Piano, Immobili.CodComune, Immobili.CodTariffa, " & _
                            "Immobili.TipoUtenza, (CodiciCerComunali.CodCategoria+CodiciCerComunali.CodSottoCategoria+CodiciCerComunali.CodCer) AS CodiceCer, " & _
                            "CodiciCerComunali.Descrizione AS Rifiuto, Movimenti.Manuale, CodiciCerComunali.FlgBonus, Movimenti.DataMovimento, Movimenti.NrMovimento, " & _
                            "Movimenti.targa, Movimenti.CodTrasportatore, Movimenti.TargaTrasportatore, Movimenti.NumeroFormulario, Movimenti.DataFormulario, " & _
                            "Movimenti.Quantita, Movimenti.unitamisura, Movimenti.BonusDovuto, Movimenti.PuntiDovuti, Movimenti.CodUser, Movimenti.DataIns, " & _
                            "Movimenti.idAlibiMemory, StatoMovimento.descrizione AS descrizioneStato, Movimenti.Stato AS Stato, Immobili.CodUtente, Causali.Descrizione AS CausaleDes " & _
                            " , Movimenti.flgRegistrato " & _
                            "FROM (((((((Movimenti " & _
                            "LEFT JOIN Immobili ON (Movimenti.CodComune = Immobili.CodComune) AND (Movimenti.CodUtente = Immobili.CodUtente) AND (Movimenti.CodImmobile = Immobili.CodImmobile)) " & _
                            "INNER JOIN CodiciCerComunali ON (Movimenti.CodCategoria = CodiciCerComunali.CodCategoria) AND (Movimenti.CodSottoCategoria = CodiciCerComunali.CodSottoCategoria) AND (Movimenti.CodCer = CodiciCerComunali.CodCer) AND " & _
                            "(Movimenti.CodComune = CodiciCerComunali.CodComune) AND (Movimenti.CodsubCer = CodiciCerComunali.CodSubCer)) " & _
                            "INNER JOIN StatoMovimento ON Movimenti.Stato = StatoMovimento.codStato) LEFT JOIN Comuni ON Immobili.CodComune = Comuni.CODISTAT) " & _
                            "LEFT JOIN Utenti ON (Movimenti.CodComune = Utenti.CodComune) AND (Movimenti.CodUtente = Utenti.CodUtente)) " & _
                            "LEFT JOIN badge ON (Movimenti.CodComune = badge.CodComune) AND (Movimenti.CodUtente = badge.CodUtente)) " & _
                            "LEFT JOIN trasportatori ON Movimenti.CodTrasportatore = trasportatori.idTrasportatore) " & _
                            "INNER JOIN Causali ON Movimenti.codCausale = Causali.codCausale " & _
                            "WHERE " & where & " order by Immobili.codUtente,Movimenti.DataMovimento Desc, iif (Utenti.Cognome = null ,Utenti.Cognome + Utenti.Nome, Utenti.RagioneSociale) ")
                                        


    lvwMovimenti.ListItems.Clear
    
    Do While Not rs.EOF
        Set item = lvwMovimenti.ListItems.Add(, "_" & rs!id, nullToBlank(rs!codUtente))
        

        item.ListSubItems.Add , , IIf(nullToBlank(rs!Cognome) = "", nullToBlank(rs!ragioneSociale), nullToBlank(rs!Cognome) & " " & nullToBlank(rs!Nome))
        item.ListSubItems.Add , , nullToBlank(rs!codBadge)
        item.ListSubItems.Add , , nullToBlank(rs!codImmobile)
        item.ListSubItems.Add , , IIf(IsNull(rs!Indirizzo), vbNullString, rs!Indirizzo & "," & nullToBlank(rs!NumCivico) & " Sub:" & nullToBlank(rs!Subalterno) & " Num.Int:" & nullToBlank(rs!NumInterno))
        item.ListSubItems.Add , , nullToBlank(rs!TipoUtenza)
        item.ListSubItems.Add , , nullToBlank(rs!codTariffa)
        item.ListSubItems.Add , , nullToBlank(rs!Targa)
        item.ListSubItems.Add , , rs!codiceCer
        item.ListSubItems.Add , , rs!Rifiuto
        item.ListSubItems.Add , , rs!nrMovimento
        item.ListSubItems.Add , , Format$(rs!DataMovimento, "dd-mmm-yyyy hh:nn:ss")
        item.ListSubItems(item.ListSubItems.Count).Tag = Format$(rs!DataMovimento, "dd/mm/yyyy hh:nn:ss")
        qty = rs!quantita.value
        item.ListSubItems.Add , , FormatNumber(qty, 2, , , vbFalse)
        item.ListSubItems(item.ListSubItems.Count).Tag = getNumber(qty)
        item.ListSubItems.Add , , rs!UnitaMisura
        item.ListSubItems.Add , , rs!BonusDovuto
        item.ListSubItems.Add , , rs!PuntiDovuti
        item.ListSubItems(item.ListSubItems.Count).Tag = getNumber(rs!BonusDovuto)
        item.ListSubItems.Add , "Causale", rs!CausaleDes
        item.ListSubItems.Add , , nullToBlank(rs!idAlibiMemory)
        item.ListSubItems.Add , "flgRegistrato", IIf(rs!flgRegistrato = -1, "Si", "No")
        item.ListSubItems.Add , "Stato", rs!descrizioneStato
        
        
        item.ListSubItems.Add , , nullToBlank(rs!CodTrasportatore)
        item.ListSubItems.Add , , nullToBlank(rs!RagioneSocialeT)
        item.ListSubItems.Add , , nullToBlank(rs!TargaTrasportatore)
        item.ListSubItems.Add , "NumeroFormulario", nullToBlank(rs!NumeroFormulario)
        item.ListSubItems.Add , , nullToBlank(rs!DataFormulario)
        item.ListSubItems.Add , , IIf(rs!manuale = -1, "Manuale", "Automatico")
        item.ListSubItems.Add , , rs!codUser
        item.ListSubItems.Add , , Format$(rs!dataIns, "dd-mmm-yyyy hh:nn:ss")
        item.ListSubItems(item.ListSubItems.Count).Tag = Format$(rs!dataIns, "dd/mm/yyyy hh:nn:ss")
        
        If (rs!Stato = "2") Then coloraRiga item, vbRed
        
        rs.MoveNext
    Loop
    
    rs.Close
    AllargaLeColonneLVW lvwMovimenti, False
    Screen.MousePointer = vbDefault
    Exit Sub
CaricaMovimentiERR:
    ScriviLog TPL_ERRORE, App.Title, "CaricaMovimenti", Err.Description

End Sub

Private Sub CaricaAccessi()
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim where As String
   
    On Error GoTo CaricaAccessiERR


    
    where = ""
    Screen.MousePointer = vbHourglass
    
    
    
    If Not IsNull(DTPicker(0).value) Then
        If where <> "" Then where = where & " and "
        where = where & "DataOraIngresso >=#" & Format$(DTPicker(2).value, "yyyy-mm-dd HH:mm:ss") & "#"
    End If
    
    If Not IsNull(DTPicker(1).value) Then
        If where <> "" Then where = where & " and "
        where = where & " DataOraIngresso <=#" & Format$(DTPicker(3).value, "yyyy-mm-dd HH:mm:ss") & " #"
    End If
    
    
    
    Set rs = gDBC.Execute("SELECT Utenti.codUtente,RilevazioneAccessi.UtenzaDomestica,RilevazioneAccessi.DaPesare ,RilevazioneAccessi.codbadge as codbadge1, " & _
                    "Badge.codbadge,Utenti.RagioneSociale, Utenti.Nome, Utenti.Cognome, Utenti.CodFiscale, Utenti.PartitaIva, " & _
                    "RilevazioneAccessi.DataOraIngresso , RilevazioneAccessi.DataOraUscita, RilevazioneAccessi.Ospite,RilevazioneAccessi.Stato " & _
                    "FROM (RilevazioneAccessi LEFT OUTER JOIN Badge ON RilevazioneAccessi.CodBadge = Badge.CodBadge and RilevazioneAccessi.codcomune = Badge.codcomune) " & _
                    "LEFT OUTER JOIN Utenti ON (Badge.CodUtente = Utenti.CodUtente) AND (Badge.CodComune = Utenti.CodComune) " & _
                    "WHERE " & where & " order by RilevazioneAccessi.DataOraIngresso desc")
                    
                                        


    lvwAccessi.ListItems.Clear
    Dim ragioneSociale As String
    
    Do While Not rs.EOF
        ragioneSociale = IIf(nullToBlank(rs!Cognome) = "", nullToBlank(rs!ragioneSociale), nullToBlank(rs!Cognome) & " " & nullToBlank(rs!Nome))
        If Trim(ragioneSociale) = vbNullString Then ragioneSociale = "Ospite"
        
        Set item = lvwAccessi.ListItems.Add(, , ragioneSociale)

        item.ListSubItems.Add , , IIf(nullToBlank(rs!codBadge) = vbNullString, rs!codbadge1, nullToBlank(rs!codBadge))
        item.ListSubItems.Add , , IIf(nullToBlank(rs!CodFiscale) = "", nullToBlank(rs!PartitaIva), nullToBlank(rs!CodFiscale))
        item.ListSubItems.Add , , Format$(rs!DataOraIngresso, "dd/MM/yy HH:mm:ss")
        item.ListSubItems.Add , , Format$(rs!DataOraUscita, "dd/MM/yy HH:mm:ss")
        item.ListSubItems.Add , , IIf(rs!ospite, "Si", "No")
        item.ListSubItems.Add , , IIf(rs!UtenzaDomestica, "No", "Si")
        item.ListSubItems.Add , , IIf(rs!DaPesare, "Si", "No")
        
        Select Case rs!Stato
            Case "0"
                item.ListSubItems.Add , , "Uscito"
            Case "1"
                item.ListSubItems.Add , , "Esportato"
            Case "2"
                item.ListSubItems.Add , , "Entrato"
            Case "3"
                item.ListSubItems.Add , , "Rifiutato"
        End Select
        
        
            
        rs.MoveNext
    Loop
    
    rs.Close
    AllargaLeColonneLVW lvwAccessi, False
    Screen.MousePointer = vbDefault
    Exit Sub
CaricaAccessiERR:
        ScriviLog TPL_ERRORE, App.Title, "CaricaAccessi", Err.Description
End Sub


Private Sub caricaImmobili(codUtente As String)
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    
    On Error GoTo caricaImmobiliERR
    lvwImmobili.ListItems.Clear
    
    If codUtente = vbNullString Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    Set rs = gDBC.Execute("SELECT Comuni.comune, Tariffe.Descrizione AS Tariffa, Immobili.* " & _
            "FROM (immobili INNER JOIN Comuni ON immobili.CodComune = Comuni.CODISTAT) " & _
            "INNER JOIN Tariffe ON immobili.CodTariffa = Tariffe.CodTariffa " & _
            "WHERE  DATE() between DataInizioValidita and DataFineValidita and   immobili.Tipo<>'Z' AND immobili.[codComune]='" & GetCurrentComune("CodIstat").value & "' AND immobili.[codUtente]=" & codUtente & _
            " ORDER BY " & IIf(gflgOrderByImmobiliUD = 1, "immobili.FlgUtenzaDomesticaResidente", "immobili.TipoUtenza Desc,immobili.Superficie Desc"))

    Do While Not rs.EOF
        Set item = lvwImmobili.ListItems.Add(, "_" & rs!codImmobile, rs!codImmobile)
        item.ListSubItems.Add , rs!tipo.Name, rs!tipo.value
        item.ListSubItems.Add , rs!Comune.Name, rs!Comune.value
        item.ListSubItems.Add , rs!TipoUtenza.Name, rs!TipoUtenza.value
        item.ListSubItems.Add , "MaxQtaFatt", getMaxQuantitaPerFatturaImmobile(rs!codComune.value, _
                                              rs!codUtente.value, rs!Indirizzo.value, nullToBlank(rs!NumCivico.value), _
                                              rs!codTariffa.value)
        item.ListSubItems.Add , rs!Indirizzo.Name, rs!Indirizzo.value & "," & rs!NumCivico.value
'       item.ListSubItems.Add , rs!Subalterno.Name, nullToBlank(rs!Subalterno.Value)
'       item.ListSubItems.Add , rs!NumInterno.Name, nullToBlank(rs!NumInterno.Value)
        item.ListSubItems.Add , rs!codTariffa.Name, nullToBlank(rs!codTariffa.value)
        item.ListSubItems.Add , rs!Tariffa.Name, nullToBlank(rs!Tariffa.value)
        
        
        If ((rs!flgUtenzaDomesticaResidente And gflgOrderByImmobiliUD = 1) Or lvwImmobili.ListItems.Count = 1) Then
            lvwImmobili.Tag = ""
            lvwImmobili_ItemCheck item
            
            item.Checked = True
            lvwImmobili.Tag = item.key
        End If
        
        rs.MoveNext
    Loop
    
    AllargaLeColonneLVW lvwImmobili, False
    Screen.MousePointer = vbDefault
    lvwImmobili.ColumnHeaders(2).width = 0
    

    Dim lastKeyImg As String
    If m_lastCheckImgRD > 0 And optImgRD.Count > m_lastCheckImgRD + 1 Then
        If optImgRD(m_lastCheckImgRD).value = 1 Then
            lastKeyImg = m_arr_optImgRD(m_lastCheckImgRD + 1).key
            optImgRD(m_lastCheckImgRD).value = 0
        End If
    End If
    
    caricaComboRifiuti imgCombo1, Combo2(0), True
    
    If lastKeyImg <> vbNullString Then
        On Error Resume Next
        m_arr_optImgRD(lastKeyImg).LinkOBJ.value = 1
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
caricaImmobiliERR:
    ScriviLog TPL_ERRORE, App.Title, "caricaImmobili", Err.Description

End Sub



Private Sub caricaComboRifiuti(imgCombo As ImageCombo, cmb As ComboBox, caricaImmagini As Boolean)
    Dim primaVolta As Boolean
    Dim keyCat As String
    Dim keySottCat As String
    Dim key As String
    Dim keyPrefisso As String
    Dim sWhere As String
    
    On Error GoTo caricaComboRifiutiERR
    
    
    'filtro su combo rifiuti
    If imgCombo.Name = "imgCombo1" Then
        sWhere = ""
        If txt(TXT_FINDCODCER).Text <> "" Then
            If IsNumeric(txt(TXT_FINDCODCER)) Then
                sWhere = " Codice like '" & txt(TXT_FINDCODCER).Text & "%'"
                caricaImmagini = False
            Else
                sWhere = " Descrizione like '" & txt(TXT_FINDCODCER).Text & "%'"
                caricaImmagini = False
            End If
        End If
        sWhere = IIf(sWhere <> "", sWhere & " and ", sWhere) & IIf(mTipoUtenza = "D", "  flgAmmessoPerUD= true ", " flgAmmessoPerUA= true ")
        
        Dim i As Integer
        If m_arr_optImgRD.Count > 0 And caricaImmagini Then
            For i = m_arr_optImgRD.Count To 1 Step -1
                m_arr_optImgRD.Remove i
            Next
        End If
    Else
        sWhere = ""
        If TxtMov(TXT_FINDCODCERMOV).Text <> "" Then
            If IsNumeric(txt(TXT_FINDCODCERMOV)) Then
                sWhere = " Codice like '" & TxtMov(TXT_FINDCODCERMOV).Text & "%'"
            Else
                sWhere = " Descrizione like '" & TxtMov(TXT_FINDCODCERMOV).Text & "%'"
            End If
        End If
    End If
    
    
    cmb.Clear
    If (rsCodiciCER Is Nothing) Then
        Set rsCodiciCER = CreateObject("ADODB.Recordset")
        rsCodiciCER.LockType = adLockBatchOptimistic
        rsCodiciCER.CursorLocation = adUseClient
        rsCodiciCER.CursorType = adOpenStatic
        
        rsCodiciCER.Source = "SELECT CodiciCer.RifiutoPericoloso, CodiciCer.codStatoFisico, CodiciCer.FlgAbilitatoFatturazione, CodiciCer.qtyMaxFatturabile, " & _
                              "CodiciCer.flgAmmessoPerUD, CodiciCer.qtyMaxGGUD, CodiciCer.qtyMaxAAUD, CodiciCer.flgAmmessoPerUA, CodiciCer.qtyMaxGGUA, " & _
                              "CodiciCer.qtyMaxAAUA, CodiciCer.costoUnitario, CodiciCer.ggMassimiGiacenza, CodiciCer.qtyMassimaGiacenza, " & _
                              "CodiciCer.ggMassimiGiacenzaNormativa, CodiciCer.qtyMassimaGiacenzaNormativa, " & _
                             "CodiciCerComunali.*, (CodiciCerComunali.CodCategoria+CodiciCerComunali.CodSottoCategoria+CodiciCerComunali.CodCer+CodiciCerComunali.CodSubCer) AS Codice, CategoriaCER.Descrizione AS DesCategoria, SottoCategoriaCER.Descrizione AS DesSottoCategoria " & _
                             "FROM ((CategoriaCER INNER JOIN SottoCategoriaCER ON CategoriaCER.CodCategoria = SottoCategoriaCER.CodCategoria) " & _
                             "INNER JOIN CodiciCerComunali ON (SottoCategoriaCER.CodCategoria = CodiciCerComunali.CodCategoria) AND (SottoCategoriaCER.CodSottocategoria = CodiciCerComunali.CodSottoCategoria)) " & _
                             "INNER JOIN CodiciCer ON (CodiciCerComunali.CodCer = CodiciCer.CodCer) AND (CodiciCerComunali.CodSottoCategoria = CodiciCer.CodSottoCategoria) AND (CodiciCerComunali.CodCategoria = CodiciCer.CodCategoria) " & _
                             "where (((CodiciCerComunali.FlgAbilitatoInRicerca) = True)) and CodiciCerComunali.codcomune='" & GetCurrentComune().Fields("CodIstat").value & "'" & _
                             "ORDER BY CodiciCerComunali.FlgBonus, CodiciCerComunali.flgIngombrante DESC , CategoriaCER.Descrizione, SottoCategoriaCER.Descrizione, CodiciCerComunali.Descrizione"
    
        
        Set rsCodiciCER.ActiveConnection = gDBC
        rsCodiciCER.Open
        Set rsCodiciCER.ActiveConnection = Nothing
        
    End If
    
    rsCodiciCER.Filter = vbNullString
    If imgCombo.Name = "imgCombo1" Then
        If sWhere <> vbNullString Then rsCodiciCER.Filter = sWhere
    End If
    
    primaVolta = False
    keyPrefisso = "_B"
    imgCombo.ComboItems.Clear
    imgCombo.ComboItems.Add , keyPrefisso, "Con Bonus", , , 0

    Do While Not rsCodiciCER.EOF
        
        If (rsCodiciCER!FlgBonus = False And Not primaVolta) Then
            keyPrefisso = "_N"
            imgCombo.ComboItems.Add , "ALTRI", "Altri", , , 0
            primaVolta = True
        
        End If
        
        
        
        If (keyCat <> keyPrefisso & rsCodiciCER!codCategoria & "00" & "00") Then
            key = keyPrefisso & rsCodiciCER!codCategoria & "00" & "00"
            'Combo1.ComboItems.Add , key, rsCodiciCer!DesCategoria & Space(5) & "[" & rsCodiciCer!codCategoria & "]", , , 3
            keyCat = key
            keySottCat = ""
        End If
        
        If (keySottCat <> keyPrefisso & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & "00") Then
            key = keyPrefisso & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & "00"
            'Combo1.ComboItems.Add , key, rsCodiciCer!DesSottoCategoria & Space(5) & "[" & rsCodiciCer!codSottoCategoria & "]", , , 4
            keySottCat = key
        End If
        
        key = keyPrefisso & rsCodiciCER!Codice
        'codice CER
        imgCombo.ComboItems.Add , key, rsCodiciCER!Descrizione & Space(5) & "[" & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & rsCodiciCER!codCer & rsCodiciCER!codSubCer & "]", IIf(rsCodiciCER!RifiutoPericoloso, 1, 2), , 9
        'corrispettivo bonus
        imgCombo.ComboItems(imgCombo.ComboItems.Count).Tag = rsCodiciCER!CorrispettivoBonus & "|" & _
            rsCodiciCER!UnitaDiMisura & "|" & rsCodiciCER!FlgAbilitatoFatturazione & "|" & _
            toZero(rsCodiciCER!qtyMaxFatturabile) & "|" & toZero(rsCodiciCER!qtyMaxGGUD) & "|" & _
            toZero(rsCodiciCER!qtyMaxAAUD) & "|" & toZero(rsCodiciCER!qtyMaxGGUA) & "|" & _
            toZero(rsCodiciCER!qtyMaxAAUA) & "|" & rsCodiciCER!flgAmmessoPerUD & "|" & rsCodiciCER!flgAmmessoPerUA & "|" & _
            rsCodiciCER!BonusPunti
        If nullToBlank(rsCodiciCER!nomeImmagine) <> vbNullString And caricaImmagini Then
            If Not m_arr_optImgRD.Contains(key) Then
                m_arr_optImgRD.Add rsCodiciCER!nomeImmagine.value, key, left$(rsCodiciCER!Descrizione, 30) & " [" & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & rsCodiciCER!codCer & rsCodiciCER!codSubCer & "]"
                
            End If
        End If
        
        
         
        cmb.AddItem rsCodiciCER!Descrizione & Space(5) & "[" & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & rsCodiciCER!codCer & rsCodiciCER!codSubCer & "]"
        cmb.ItemData(cmb.ListCount - 1) = imgCombo.ComboItems.Count
        
        
        rsCodiciCER.MoveNext
    Loop
    
    If caricaImmagini Then
        If m_arr_optImgRD.Count > 0 And Not m_arr_optImgRD.Contains("SEARCH") Then
            m_arr_optImgRD.Add "\IconsPGM\Search2.ico", "SEARCH", "Ricerca"
        End If
        
        m_lastCheckImgRD = 1
        generaIconeRD
        
        'If m_arr_optImgRD.Count > 0 Then
        '    If picSearchRD.Visible = True Then optImgRD(optImgRD.UBound).Value = 1
        'End If
    End If
    Exit Sub
caricaComboRifiutiERR:
    ScriviLog TPL_ERRORE, App.Title, "caricaComboRifiuti", Err.Description

End Sub
    


Private Sub Form_Resize()

    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelY
    
    resizeControl
End Sub

Private Sub resizeControl()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    
    resizeCommandButtons
    
    TabStrip.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight '- cmd(0).Height - 2 * offsetX
    
    posTop = TabStrip.ClientTop
    posLeft = TabStrip.ClientLeft
    
    For i = picContainer.LBound To picContainer.UBound
        picContainer(i).Move posLeft, posTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    Next i
    
    ProgressBar1.top = Me.ScaleHeight - ProgressBar1.Height - offsetY
    lblEsportazione.top = ProgressBar1.top - lblEsportazione.Height - offsetY
    
    PicTotali.Move offsetX * 2, cmd(0).top + cmd(0).Height / 2
    
    
    
    
End Sub

Private Sub resizeCommandButtons()
    Dim i As Integer
    Dim posTop As Long
    Dim posLeft As Long
    
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.UBound
    
        If i > 0 Then
            If Val(TabStrip.Tag) = 0 Then
                cmd(i).Visible = InStr(1, ButtonsRicercaMov, CStr(i) & ",") = 0
            ElseIf Val(TabStrip.Tag) = 1 Then
                cmd(i).Visible = InStr(1, ButtonsRicercaMov, CStr(i) & ",") > 0
            Else
                cmd(i).Visible = InStr(1, ButtonsRicercaacc, CStr(i) & ",") > 0
            End If
        End If
        
        If cmd(i).Visible Then
            If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
            cmd(i).Move posLeft, posTop
        End If
    Next
End Sub


Public Sub ShowME(key As String, Optional Esportazione As Boolean = False)
    

    
    Load Me
    
    TabStrip.Tabs(key).Selected = True
    
    Set m_cHdrIcons.ListView = lvwMovimenti
    lvwMovimenti.SortOrder = lvwAscending
    lvwMovimenti.SortKey = 1
    Call m_cHdrIcons.SetHeaderIcons(lvwMovimenti.SortKey, lvwMovimenti.SortOrder)
        
        
    frmScontrino.ShowME
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        
'        Me.width = 1024 * 15
'        Me.Height = 600 * 15

        
        Me.Show
    Else
        Me.Show
    End If
    
    m_esportazione = Esportazione
    
    If m_esportazione Then
        TabStrip.Tabs.Remove (1)
        ButtonsRicercaMov = CMD_ESCI & "," & CMD_ESPORTAMOV & "," & CMD_RICERCAMOV & ","
        DTPicker(0).value = Year(Now) & "-01-01"
        resizeCommandButtons
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lStyle As Long
    
    If gStartMode = 0 Or gStartMode = 1 Then
        lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
        lStyle = lStyle Or MY_WS_DOCK
        SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    
        If gStartMode = 1 Then
            SetParent Me.hWnd, hOldParent
        End If
    End If
    
    Esci = True
    
    If Not rsCodiciCER Is Nothing Then
        rsCodiciCER.Close
        Set rsCodiciCER = Nothing
    End If
    
    If Not rsCliente Is Nothing Then
        rsCliente.Close
        Set rsCliente = Nothing
    End If
    
    Set pfrmPesata = Nothing
    
    If gActiveAccess Then
        Set rsTotaleAccessi.ActiveConnection = Nothing
        If rsTotaleAccessi.state = adStateOpen Then rsTotaleAccessi.Close
        Set rsTotaleAccessi = Nothing
    End If
    
    Unload frmScontrino
End Sub

Private Sub ImgCombo2_Change()
Dim l As Long
Static Changing As Boolean
Static Found As Boolean
Static PartialText As String

On Error GoTo ImgCombo2_ChangeERR


If Changing Then Exit Sub Else Changing = True ' questo evita le chiamate nidificate della routine

If ImgCombo2.Text = "" Then
    Combo2(1).Text = ""
    GoTo Esci ' se non vi è testo esce dalla routine
End If

Combo2(1).Text = ImgCombo2.Text
PartialText = ImgCombo2.Text      'memorizza la parte di testo digitata

If Canceling Then GoTo Esci     ' se il tasto Cancel è premuto esce dalla routine

If PartialText = "" Then GoTo Esci

l = SendMessage(Combo2(1).hWnd, CB_FINDSTRING, -1, ByVal PartialText)    ' trova l'indice per l'elemento più simile al quello digitato


If l > -1 Then   ' se l'elemento viene trovato
    ' il testo parziale viene completato con quello trovato dall'API
    Combo2(1).Text = PartialText & Mid(Combo2(1).List(l), Len(PartialText) + 1)
    'ImageCombo1.Text = Combo2.Text
    ImgCombo2.ComboItems(Combo2(1).ItemData(l)).Selected = True
    If Not Deleting Then
        Combo2(1).SelStart = Len(PartialText)
    Else
        Combo2(1).SelStart = Len(PartialText) - IIf(Found, 1, 0)
    End If
    
    
    Combo2(1).SelLength = Len(Combo2(1).Text)
    
    ImgCombo2.SelStart = Combo2(1).SelStart
    ImgCombo2.SelLength = Combo2(1).SelLength
    
    DoEvents
    
    PartialText = Mid(Combo2(1).Text, 1, Combo2(1).SelStart)
    If PartialText = "" Then ImgCombo2.Text = "": Combo2(1).Text = ""
    Deleting = False
    Found = True
Else
    Found = False
    
  
End If

Esci:
    Changing = False
    Deleting = False
    Canceling = False
    Exit Sub
ImgCombo2_ChangeERR:
    ScriviLog TPL_ERRORE, App.Title, "ImgCombo2_Change", Err.Description

End Sub

Private Sub lblCambiaPesa_Click()
    'se sono in lettura continua annullo
    On Error GoTo lblCambiaPesa_ClickERR

    If cmd(CMD_LETTURAQUANTITA).Caption = "&Annulla" Then cmd_Click (CMD_LETTURAQUANTITA)
    mRichiestaPesoInCorso = False
    frmCambiaPesa.Show vbModal
    Exit Sub
lblCambiaPesa_ClickERR:
        ScriviLog TPL_ERRORE, App.Title, "lblCambiaPesa_Click", Err.Description

End Sub

Private Sub lvwImmobili_ItemCheck(ByVal item As MSComctlLib.ListItem)
    On Error GoTo lvwImmobili_ItemCheckERR
    
    
    If lvwImmobili.Tag <> vbNullString And lvwImmobili.ListItems.Count > 1 Then lvwImmobili.ListItems(lvwImmobili.Tag).Checked = False
        
    lvwImmobili.Tag = item.key
    mTipoUtenza = item.ListSubItems("TipoUtenza").Text
    
    
    txt(TXT_COD_TRASPORTATORE).enabled = True
    txt(TXT_NRFORMULARIO).enabled = True
    dtpickFormulario.enabled = True
    cmbTarga(1).enabled = True
    cmbRicerca(1).enabled = True
    
    item.Selected = True
    
    getTotaliPerCodiceCer Mid$(item.key, 2)
Exit Sub
lvwImmobili_ItemCheckERR:
     ScriviLog TPL_ERRORE, App.Title, "lvwImmobili_ItemCheck", Err.Description

End Sub

Private Sub lvwMovimenti_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    SortListview lvwMovimenti, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwMovimenti_ItemClick(ByVal item As MSComctlLib.ListItem)
    On Error Resume Next
    cmd(CMD_CANCELLAMOV).enabled = item.ListSubItems("Stato").Text <> "Contabilizzato" And item.ListSubItems("flgRegistrato").Text <> "Si"
End Sub


Private Sub optImgRD_Click(Index As Integer)
On Error GoTo optImgRD_ClickERR
    
    If m_noEvent Then Exit Sub
  
    'simulo options button
    
    If m_lastCheckImgRD <> Index And optImgRD(Index).value = 1 And m_lastCheckImgRD < optImgRD.UBound Then
        m_noEvent = True
        optImgRD(m_lastCheckImgRD).value = 0
        optImgRD(m_lastCheckImgRD).BackColor = &H8000000F
        m_noEvent = False
      
    End If
    
         
    Select Case Index
        Case optImgRD.UBound
        
            'ricerca
            If optImgRD(Index).value = 1 Then
                picSearchRD.Visible = True
                picSearchRD.ZOrder 0
                picImageRD.ZOrder 1
                
                If optImgRD(Index).left < picSearchRD.width Then
                    optImgRD(Index).left = picImageRD.ScaleWidth - optImgRD(Index).width
                End If
                m_lastCheckImgRD = Index
                
                optImgRD(Index).BackColor = &HFF&
                
            Else
                
                optImgRD(Index).left = optImgRD(Index - 1).left + optImgRD(Index).width + 5 * 15
                picSearchRD.ZOrder 1
                picImageRD.ZOrder 0
                picSearchRD.Visible = False
                'sync con bottoni
                Dim chkBox As CheckBox
                
                If imgCombo1.SelectedItem Is Nothing Then Exit Sub
                If m_arr_optImgRD.Contains(imgCombo1.SelectedItem.key) Then
                    Set chkBox = m_arr_optImgRD(imgCombo1.SelectedItem.key).LinkOBJ
                    chkBox.value = 1
                End If
                optImgRD(Index).BackColor = &H8000000F
             
                
            End If
        Case Else
            If optImgRD(Index).value = 1 Then
                optImgRD(Index).BackColor = &HFF&
                txt(TXT_FINDCODCER).Text = Mid$(m_arr_optImgRD.item(Index + 1).key, 3)
                m_lastCheckImgRD = Index
                ShowMessage ErrorRifiuto, vbNullString
            Else
                optImgRD(Index).BackColor = &H8000000F
                txt(TXT_FINDCODCER).Text = vbNullString
            End If
    End Select
Exit Sub
optImgRD_ClickERR:
     ScriviLog TPL_ERRORE, App.Title, "optImgRD_Click", Err.Description

    
End Sub

Private Sub picContainer_Resize(Index As Integer)
    Dim posY As Long
    Dim posX As Long
    Dim fatt As Double
    
    fatt = IIf(gActiveAccess = 1, 0.65, 1)
    On Error Resume Next
    Select Case Index
        Case 0
            If gStartMode = 0 Then
                posX = fmeTrasportatore.left
                posY = fmeTrasportatore.top + offsetY + fmeTrasportatore.Height
                
                'Immobili
                Label1(LBL_IMMOBILI).Move posX, posY
                lvwImmobili.Move posX, Label1(LBL_IMMOBILI).top + Label1(LBL_IMMOBILI).Height, picContainer(Index).ScaleWidth - 2 * offsetX, (picContainer(Index).ScaleHeight - (Label1(LBL_IMMOBILI).top + Label1(LBL_IMMOBILI).Height)) \ 2
                                
                'utente
                fmeUtente.width = lvwImmobili.width
                
                'trasportatore
                fmeTrasportatore.width = fmeUtente.width
                
                
                'totali
                fmeTotali.Move lvwImmobili.left, lvwImmobili.top + lvwImmobili.Height + offsetY, lvwImmobili.width, picContainer(Index).ScaleHeight - (lvwImmobili.top + lvwImmobili.Height + offsetY)
                lvwTotMovimenti.Move offsetX, offsetY, fmeTotali.width * fatt - (2 * offsetX), fmeTotali.Height - 2 * offsetY
                
                'accessi
                frmAccessi.Move lvwTotMovimenti.left + lvwTotMovimenti.width + (2 * offsetX), offsetY, fmeTotali.width * 0.35 - (2 * offsetX), fmeTotali.Height - cmd(0).Height - offsetY * 2
                lblTotaleAccessi.Move offsetX, frmAccessi.Height / 2 - lblTotaleAccessi.Height / 2, frmAccessi.width - 2 * offsetX
                imgTrafficLightGreen.Move offsetX, offsetY * 2
                imgTrafficLightRed.Move offsetX, offsetY * 2
                
                Set txtInfoUser.Container = fmeUtente
                txtInfoUser.Move cmbTarga(0).left + cmbTarga(0).width + offsetX, cmbTarga(0).top - offsetY, fmeUtente.width - cmbTarga(0).left - cmbTarga(0).width - 2 * offsetX, fmeUtente.Height - offsetY * 2 - fmeUtente.top
                
            Else 'mobile
                
                
                'trasportatore
                fmeTrasportatore.width = picContainer(Index).ScaleWidth - 2 * offsetX
                
                'totali
                fmeTotali.Move fmeTrasportatore.left, fmeTrasportatore.top + fmeTrasportatore.Height, fmeTrasportatore.width, picContainer(Index).ScaleHeight - (fmeTrasportatore.top + fmeTrasportatore.Height + offsetY)
            
                'se gli immobili non sono gestiti le informazioni utente le metto all'interno del frame utente
                'Immobili
                If gflgImmobili = 1 Then
                    posX = fmeUtente.width + 2 * offsetX
                    posY = fmeUtente.top - 15 * 5
                        
                    Label1(LBL_IMMOBILI).Move posX, posY
                    lvwImmobili.Move posX, fmeUtente.top + offsetY, picContainer(Index).ScaleWidth - offsetX - fmeUtente.width - fmeUtente.left, fmeUtente.Height - offsetY
                    
                    txtInfoUser.Move offsetX, offsetY + lblErrorMaxUtentiNelIsola.Height, fmeTotali.width - 2 * offsetX
                    posY = txtInfoUser.top + txtInfoUser.Height
                    lvwTotMovimenti.Move offsetX, posY, fmeTotali.width * fatt - (2 * offsetX), fmeTotali.Height - posY - offsetY
                    frmAccessi.Move lvwTotMovimenti.left + lvwTotMovimenti.width + (2 * offsetX), offsetY, fmeTotali.width * 0.35 - (2 * offsetX), fmeTotali.Height - cmd(0).Height - offsetY
                    lblTotaleAccessi.Move offsetX, frmAccessi.Height / 2 - lblTotaleAccessi.Height / 2, frmAccessi.width - 2 * offsetX
                    
                Else
                    'frame utenti
                    fmeUtente.width = picContainer(Index).ScaleWidth - 2 * offsetX
                    
                    Set txtInfoUser.Container = fmeUtente
                    txtInfoUser.Move cmbTarga(0).left + cmbTarga(0).width + offsetX, cmbTarga(0).top - offsetY, fmeUtente.width - cmbTarga(0).left - cmbTarga(0).width - 2 * offsetX
                    lvwTotMovimenti.Move offsetX, offsetY + lblErrorMaxUtentiNelIsola.Height, fmeTotali.width * fatt - (2 * offsetX), fmeTotali.Height - 2 * offsetY
                    frmAccessi.Move lvwTotMovimenti.left + lvwTotMovimenti.width + (2 * offsetX), offsetY, fmeTotali.width * 0.35 - (2 * offsetX), fmeTotali.Height - cmd(0).Height - offsetY
                    lblTotaleAccessi.Move offsetX, frmAccessi.Height / 2 - lblTotaleAccessi.Height / 2, frmAccessi.width - 2 * offsetX
                End If
                
                lblErrorMaxUtentiNelIsola.Move offsetX, offsetY, frmAccessi.left - 2 * offsetX
                lblErrorUtenzaDomestica.Move offsetX, frmAccessi.Height - offsetY - lblErrorUtenzaDomestica.Height, frmAccessi.width - 2 * offsetX
                
            End If
            
            lblInfoCodCer.width = picContainer(Index).ScaleWidth - 2 * offsetX
            lblErrorRifiuto.Move offsetX, lblInfoCodCer.top, lblInfoCodCer.width
            
            Label1(LBL_FATTURABILE).Move lblInfoCodCer.width - Label1(LBL_FATTURABILE).width - 2 * offsetX, lblInfoCodCer.top
            
            
            picImageRD.width = picContainer(0).ScaleWidth - picImageRD.left - offsetX
            picSearchRD.left = picImageRD.left
            picSearchRD.width = picImageRD.width - optImgRD(0).width - 5 * 15
            
            imgCombo1.width = picSearchRD.ScaleWidth - imgCombo1.left - 2 * offsetX
            
            tmrGeneraImgRD.enabled = True
            
            
            
        Case 1
            posY = cmbRicercaMov(2).top + offsetY + cmbRicercaMov(2).Height
            
            cmbRicercaMov(2).width = picContainer(1).ScaleWidth - cmbRicercaMov(2).left - 2 * offsetX
            
            ImgCombo2.width = picContainer(1).ScaleWidth - ImgCombo2.left - 2 * offsetX
            
            lvwMovimenti.Move offsetX, posY, picContainer(Index).ScaleWidth - 2 * offsetX, picContainer(Index).ScaleHeight - posY - cmd(0).Height - offsetY
        Case 2
            
            posY = DTPicker(2).top + offsetY + DTPicker(2).Height
            
            
            txtBadgeAccessi.Move offsetX, picContainer(Index).ScaleHeight - txtBadgeAccessi.Height - offsetY
            lblBadgeAccessi.Move offsetX, txtBadgeAccessi.top - lblBadgeAccessi.Height
            cmbRicerca(2).Move txtBadgeAccessi.width + 2 * offsetX, picContainer(Index).ScaleHeight - txtBadgeAccessi.Height - offsetY
            lvwAccessi.Move offsetX, posY, picContainer(Index).ScaleWidth - 2 * offsetX, picContainer(Index).ScaleHeight - posY - cmd(0).Height - offsetY
    
    End Select
    
End Sub

Private Function generaIconeRD() As Boolean
    Dim width As Long
    Dim imgWidth As Long
    Dim numImg As Long
    Dim i As Long
    
    On Error GoTo generaIconeRDERR
    
     optImgRD(0).Visible = True
    imgWidth = optImgRD(0).width + (5 * 15)
    width = picImageRD.ScaleWidth
    
    generaIconeRD = False
    
    numImg = optImgRD.UBound
    
    For i = numImg To 1 Step -1
        Unload optImgRD(i)
    Next
    
    numImg = width \ imgWidth
    
    If numImg > m_arr_optImgRD.Count Then numImg = m_arr_optImgRD.Count
    
    
    For i = 0 To numImg - 1
        If i > 0 Then
            Load optImgRD(i)
            optImgRD(i).Move optImgRD(i - 1).left + (5 * 15) + optImgRD(i - 1).width, optImgRD(i - 1).top
            optImgRD(i).Visible = True
            
        End If
        
        Set optImgRD(i).Picture = LoadPicture(App.Path & "\Icons\" & m_arr_optImgRD(i + 1).nameImage)
        Set m_arr_optImgRD(i + 1).LinkOBJ = optImgRD(i)
        optImgRD(i).ToolTipText = m_arr_optImgRD(i + 1).Description
    Next
    
    
    
    If optImgRD.Count = 1 Then optImgRD(0).Visible = False
    
    generaIconeRD = True
    
    picImageRD.Visible = True
    
    Exit Function
generaIconeRDERR:
    picImageRD.Visible = False
    Debug.Print Err.Description
End Function





Private Sub TabStrip_Click()
    
    If (TabStrip.Tag <> vbNullString) Then
        picContainer(TabStrip.Tag).Visible = False
    End If
    
    picContainer(TabStrip.SelectedItem.Index - 1).Visible = True
    TabStrip.Tag = TabStrip.SelectedItem.Index - 1
    
    cmd(CMD_RICERCAMOV).enabled = TabStrip.Tag = 1 Or TabStrip.Tag = 2
    cmd(CMD_RISTAMPAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_ESPORTAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_CANCELLAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_RISTAMPAMOV).Visible = TabStrip.Tag <> 2
    cmd(CMD_ESPORTAMOV).Visible = TabStrip.Tag <> 2
    cmd(CMD_CANCELLAMOV).Visible = TabStrip.Tag <> 2
    PicTotali.Visible = TabStrip.Tag = 0
    
    If TabStrip.Tag = 2 Then
        
        
        DTPicker(2).value = DateAdd("d", -1, DateTime.Date)
        DTPicker(3).value = DateAdd("s", 59, DateAdd("n", 59, DateAdd("h", 23, DateTime.Date)))
        CaricaAccessi
    End If
    
    If cmd(CMD_RICERCAMOV).enabled And TabStrip.Tag = 1 Then
        cmd(CMD_LETTURAQUANTITA).Tag = cmd(CMD_LETTURAQUANTITA).enabled
        cmd(CMD_SALVAESTAMPA).Tag = cmd(CMD_SALVAESTAMPA).enabled
        cmd(CMD_SALVAMOV).Tag = cmd(CMD_SALVAMOV).enabled
        cmd(CMD_LETTURAQUANTITA).enabled = False
        cmd(CMD_SALVAESTAMPA).enabled = False
        cmd(CMD_SALVAMOV).enabled = False
        
        If txt(TXT_COD_UTENTE).Text <> "" Then
            TxtMov(0).Text = txt(TXT_COD_UTENTE).Text
            cmdRicercaUtenteMov(0).Tag = "1"
            cmdRicercaUtenteMov_Click (0)
        End If
    ElseIf cmd(CMD_LETTURAQUANTITA).Tag <> "" Then
        cmd(CMD_LETTURAQUANTITA).enabled = cmd(CMD_LETTURAQUANTITA).Tag
        cmd(CMD_SALVAESTAMPA).enabled = cmd(CMD_SALVAESTAMPA).Tag
        cmd(CMD_SALVAMOV).enabled = cmd(CMD_SALVAMOV).Tag
    
    End If
    
    
    resizeCommandButtons
    On Error Resume Next
    If TabStrip.Tag = 0 Then
        txtBadge.SetFocus
    ElseIf TabStrip.Tag = 2 Then
        txtBadgeAccessi.SetFocus
    End If
End Sub

Private Sub letturaPeso(codPesa As String)
    Dim peso As Double
    Dim messaggio As String
    
        
    
    peso = 0
    Do While peso = 0 And Not Esci
        
        'richiesta peso alla bilancia
        If Not mRichiestaPesoInCorso Then
            If InvioComandoAllaBilancia(cmd_RichiestaPesa, True) Then mRichiestaPesoInCorso = True
           
        End If
        
        
        
        Select Case pesa(codPesa).ReadMessage(messaggio)
            Case ErroreRicevuto
                Debug.Print "errore in lettura"
                mRichiestaPesoInCorso = False
            Case InLettura
            Case LetturaTerminata
                'messaggio = "1402/07/200815:17:1702/07/200815:17:34   144   145    200101       CARTA CARTONE    200101       CARTA CARTONE   1170007230010                         3.25 kg                             0.00 kg        0 kg        0 kg     3,25 kg21"
                Select Case getCurrentPesa.ModelloPesa
                    Case "LAUMAS400"
                        msgPesa = getMessaggioPesataLAumas400(messaggio, m_badge)
                        If pesa(codPesa).COMControl.PortOpen Then
                             pesa(codPesa).COMControl.PortOpen = False
                             pesa(codPesa).COMControl.PortOpen = True
                        End If
                    Case "RETEL"
                        msgPesa = getMessaggioPesataRETEL(messaggio, vbNullString)
                        If pesa(codPesa).COMControl.PortOpen Then
                             pesa(codPesa).COMControl.PortOpen = False
                             pesa(codPesa).COMControl.PortOpen = True
                        End If
                    Case Else
                        msgPesa = getMessaggioPesata(messaggio)
                End Select
                
                messaggio = ""
                
                mRichiestaPesoInCorso = False
                
                ImpostaDatiDaPesata msgPesa
                
                'annulla pesata
                If Not getCurrentPesa.ModalitaLetturaPesoAutomatica Then
                    cmd_Click CMD_LETTURAQUANTITA
                Else
                    peso = 0 'ritorno in lettura peso
                End If
    
         End Select
         
         DoEvents
    Loop
    
    If Esci Then peso = 0
    
    Exit Sub
letturaPesoERR:
    'in caso di errore annullo la pesata
    cmd_Click CMD_LETTURAQUANTITA
End Sub

Private Sub ImpostaDatiUtenteDaBadge(codBadge As String)
    On Error GoTo ImpostaDatiUtenteDaBadgeERR

    If (mlastUtente <> Val(txt(TXT_COD_UTENTE).Text) Or (Trim$(codBadge) <> Trim$(mlastBadge))) And Val(codBadge) <> 0 Then
    
          'se la gestione immobili e disabilitata l'utente e l'immobile devo essere creati dalla procedura
        If gflgImmobili = 0 Then creaImmobileUtenteEBadge codBadge
                
        txt(TXT_COD_UTENTE).Text = getCodiceUtenteDaBadge(codBadge)
        If txt(TXT_COD_UTENTE) = "" Then
            MsgBox "Attenzione utente non trovato per codice:" & codBadge
            Exit Sub
        Else
            caricaDatiUtente txt(TXT_COD_UTENTE).Text
        End If
    End If
    Exit Sub
ImpostaDatiUtenteDaBadgeERR:
    ScriviLog TPL_ERRORE, App.Title, "ImpostaDatiUtenteDaBadge", Err.Description

End Sub



Private Sub ImpostaDatiDaPesata(msgPesa As messaggioPesa)
    If msgPesa.PesoNetto < 0 Then Exit Sub
            
    On Error GoTo ImpostaDatiDaPesataERR
            
            
    ImpostaDatiUtenteDaBadge msgPesa.codBadge
            
    'memorizzo la modalità da PESA
    mPesoLettoDaBilancia = True
            
    txt(TXT_QUANTITA).Text = Format$(msgPesa.PesoNetto, "#.00")
    cmd(CMD_SALVAESTAMPA).enabled = True
    cmd(CMD_SALVAMOV).enabled = True
    
    If Trim$(msgPesa.CodRifiuto) <> vbNullString Then
        'imposta il codice cer rilevato
        txt(TXT_FINDCODCER).Text = ""
        txt(TXT_FINDCODCER).Text = IIf(Trim$(msgPesa.CodRifiuto) = "999999", "", Trim$(msgPesa.CodRifiuto))
        If Trim$(msgPesa.CodRifiuto) = "999999" Then
            imgCombo1.Text = ""
            On Error Resume Next
            imgCombo1.SetFocus
        End If
    End If
    Exit Sub
ImpostaDatiDaPesataERR:
    ScriviLog TPL_ERRORE, App.Title, "ImpostaDatiUtenteDaBadge", Err.Description
    
End Sub



Private Sub imgCombo1_change()
Dim l As Long
Static Changing As Boolean
Static Found As Boolean
Static PartialText As String

On Error GoTo imgCombo1_changeERR


If Changing Then Exit Sub Else Changing = True ' questo evita le chiamate nidificate della routine

If imgCombo1.Text = "" Then
    Combo2(0).Text = ""
    GoTo Esci ' se non vi è testo esce dalla routine
End If

Combo2(0).Text = imgCombo1.Text
PartialText = imgCombo1.Text      'memorizza la parte di testo digitata

If Canceling Then GoTo Esci     ' se il tasto Cancel è premuto esce dalla routine

If PartialText = "" Then GoTo Esci

l = SendMessage(Combo2(0).hWnd, CB_FINDSTRING, -1, ByVal PartialText)    ' trova l'indice per l'elemento più simile al quello digitato


If l > -1 Then   ' se l'elemento viene trovato
    ' il testo parziale viene completato con quello trovato dall'API
    Combo2(0).Text = PartialText & Mid(Combo2(0).List(l), Len(PartialText) + 1)
    'imgcombo1.Text = combo2(0).Text
    imgCombo1.ComboItems(Combo2(0).ItemData(l)).Selected = True
    If Not Deleting Then
        Combo2(0).SelStart = Len(PartialText)
    Else
        Combo2(0).SelStart = Len(PartialText) - IIf(Found, 1, 0)
    End If
    
    
    Combo2(0).SelLength = Len(Combo2(0).Text)
    
    imgCombo1.SelStart = Combo2(0).SelStart
    imgCombo1.SelLength = Combo2(0).SelLength
    
    DoEvents
    
    imgCombo1_Click
    
    PartialText = Mid(Combo2(0).Text, 1, Combo2(0).SelStart)
    If PartialText = "" Then imgCombo1.Text = "": Combo2(0).Text = ""
    Deleting = False
    Found = True
Else
    Found = False
    
  
End If

Esci:
    Changing = False
    Deleting = False
    Canceling = False
Exit Sub
imgCombo1_changeERR:
    ScriviLog TPL_ERRORE, App.Title, "imgCombo1_change", Err.Description

End Sub



Private Sub imgCombo1_keydown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then Canceling = True
If KeyCode = vbKeyBack And imgCombo1.Text <> "" Then Deleting = True: KeyCode = 0
End Sub

Private Sub imgCombo2_keydown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDelete Then Canceling = True
    If KeyCode = vbKeyBack And imgCombo1.Text <> "" Then Deleting = True: KeyCode = 0
End Sub


Private Sub tmrAccessi_Timer()
    On Error GoTo tmrAccessi_TimerERR

    'ogni n sec cancello gli utenti che teoricamente dovrebbero essere usciti dall'isola
    'ragionando su un tempo medio di accesso e rilascio rifiuti
    GestioneAccessi -1, False, False, False, Pulizia
    
    Set rsTotaleAccessi.ActiveConnection = gDBC
    rsTotaleAccessi.Requery
    
    lblTotaleAccessi.Caption = rsTotaleAccessi.Fields(0).value
    
    
      
    ShowMessage ErrorMaxUtentiNelIsola, IIf(rsTotaleAccessi.Fields(0).value >= gMaxUsersIsland, "SUPERATO LIMITE MASSIMO UTENZE PIATTAFORMA", vbNullString)
    
        
            
    Set rsTotaleDaPesare.ActiveConnection = gDBC
    rsTotaleDaPesare.Requery
    If rsTotaleDaPesare.EOF Or rsTotaleDaPesare(0).value = 0 Then
        ShowMessage WarningUtenzaNonDomestica, vbNullString
    Else
        ShowMessage WarningUtenzaNonDomestica, "UTENZE NON DOMESTICHE:" & CStr(rsTotaleDaPesare(0).value)
    End If
    
    DoEvents
    Exit Sub
tmrAccessi_TimerERR:
    ScriviLog TPL_ERRORE, App.Title, "tmrAccessi_Timer", Err.Description
    
End Sub

Private Sub ShowMessage(typeMessage As typeAlertMessage, message As String)
On Error GoTo ShowMessageERR

    Select Case typeMessage
        Case typeAlertMessage.ErrorMaxUtentiNelIsola
            lblErrorMaxUtentiNelIsola.Visible = message <> vbNullString
            lblErrorMaxUtentiNelIsola.Caption = message
        Case typeAlertMessage.ErrorRifiuto
            lblErrorRifiuto.Visible = message <> vbNullString
            lblErrorRifiuto.Caption = message
        Case typeAlertMessage.WarningUtenzaNonDomestica
            lblErrorUtenzaDomestica.Visible = message <> vbNullString
            lblErrorUtenzaDomestica.Caption = message
        Case typeAlertMessage.InfoCodCer
            lblInfoCodCer.Caption = message
    End Select
    Exit Sub
ShowMessageERR:
    ScriviLog TPL_ERRORE, App.Title, "ShowMessage", Err.Description
    
End Sub

Private Sub tmrGeneraImgRD_Timer()
    On Error Resume Next
    tmrGeneraImgRD.enabled = False
    generaIconeRD
End Sub


Private Sub txt_Change(Index As Integer)
    Dim cmbItem As ComboItem
    Dim sWhere As String
    On Error GoTo txt_ChangeERR
    
    If Index = TXT_FINDCODCER Then
    
        caricaComboRifiuti imgCombo1, Combo2(0), False
        'mi posiziono sul primo rifiuto
        For Each cmbItem In imgCombo1.ComboItems
            If (Mid$(cmbItem.key, 5, 2) <> "00") And Len(cmbItem.key) >= 6 Then
                cmbItem.Selected = True
                Exit Sub
            End If
        Next
    ElseIf Index = TXT_COD_UTENTE Then
    End If
    
Exit Sub
txt_ChangeERR:
    ScriviLog TPL_ERRORE, App.Title, "txt_Change", Err.Description

    
End Sub


Public Sub applicaRestrizioniUtente()
    'txt(TXT_COD_UTENTE).Enabled = frmLogin.Livello <> LVL_USER
    'txt(TXT_FINDCODCER).Enabled = frmLogin.Livello <> LVL_USER
    'txt(TXT_QUANTITA).Enabled = frmLogin.Livello <> LVL_USER
    
    'imgCombo1.Enabled = frmLogin.Livello <> LVL_USER
    'cmbRicerca(0).Enabled = frmLogin.Livello <> LVL_USER
    'cmd(CMD_SALVA).Enabled = frmLogin.Livello <> LVL_USER
End Sub

Function getPropietaCER(sp As String, proprieta As ProprietaCER) As String
    'CorrispettivoBonus|UnitaDiMisura|FlgAbilitatoFatturazione|qtyMaxFatturabile|qtyMaxGGUD|qtyMaxAAUD|qtyMaxGGUA|qtyMaxAAUA
    Dim s() As String
    
    s = Split(sp, "|")
    
    Select Case proprieta
        Case ProprietaCER.CorrispettivoBonus
            getPropietaCER = s(0)
        Case ProprietaCER.UnitaDiMisura
            getPropietaCER = s(1)
        Case ProprietaCER.FlgAbilitatoFatturazione
            getPropietaCER = s(2)
        Case ProprietaCER.qtyMaxYYUA
            getPropietaCER = s(7)
        Case ProprietaCER.qtyMaxYYUD
            getPropietaCER = s(5)
        Case ProprietaCER.qtyMaxGGUA
            getPropietaCER = s(6)
        Case ProprietaCER.qtyMaxGGUD
            getPropietaCER = s(4)
        Case ProprietaCER.qtyMaxFatturabile
            getPropietaCER = s(3)
        Case ProprietaCER.flgAmmessoPerUD
            getPropietaCER = s(8)
        Case ProprietaCER.flgAmmessoPerUA
            getPropietaCER = s(9)
        Case ProprietaCER.CorrispettivoBonusPunti
            getPropietaCER = s(10)
    
    End Select
End Function


Function getTotaleMovimentiAnnoPerImmobile(codImmobile As Long) As Double
    Dim rs As ADODB.Recordset
    
    On Error GoTo getTotaleMovimentiAnnoPerImmobileERR


    Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                "From Movimenti " & _
                "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                " and  Movimenti.DataMovimento Between dateadd('d', datediff('d',  date()  ,year(date()) &'-01-01' ) , date()) +'00.00.00' And Now() AND  Movimenti.CodImmobile=" & codImmobile)


    
    getTotaleMovimentiAnnoPerImmobile = 0
    If Not rs.EOF Then
        getTotaleMovimentiAnnoPerImmobile = 0
        If Not IsNull(rs(0)) Then
            getTotaleMovimentiAnnoPerImmobile = rs(0)
        End If
        rs.Close
    End If
Exit Function
getTotaleMovimentiAnnoPerImmobileERR:
    ScriviLog TPL_ERRORE, App.Title, "getTotaleMovimentiAnnoPerImmobileERR", Err.Description

End Function

Function getTotaleMovimentiAnnoPerCER(codiceCer As String, codImmobile As Long) As Double
    Dim rs As ADODB.Recordset
    Dim codCategoria As String, codSottoCategoria As String, codCer As String
    Dim codSubCer As String
    
    On Error GoTo getTotaleMovimentiAnnoPerCERERR
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codSubCer = Right$(codCer, 2)
    codCer = Mid$(codCer, 5, 2)

    Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                "From Movimenti " & _
                "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                " and  Movimenti.DataMovimento Between dateadd('d', datediff('d',  date()  ,year(date()) &'-01-01' ) , date()) +'00.00.00' And Now() AND  Movimenti.CodImmobile=" & codImmobile & _
                " AND Movimenti.CodCategoria='" & codCategoria & "' AND Movimenti.CodSottoCategoria='" & codSottoCategoria & "'  AND Movimenti.CodCer='" & codCer & "'")


    
    getTotaleMovimentiAnnoPerCER = 0
    If Not rs.EOF Then
        getTotaleMovimentiAnnoPerCER = 0
        If Not IsNull(rs(0)) Then
            getTotaleMovimentiAnnoPerCER = rs(0)
        End If
        rs.Close
    End If
    Exit Function
getTotaleMovimentiAnnoPerCERERR:
    ScriviLog TPL_ERRORE, App.Title, "getTotaleMovimentiAnnoPerCER", Err.Description
    
End Function

Function getTotaleMovimentiNelPeriodoPerCER(codiceCer As String, codImmobile As Long) As Double
    Dim rs As ADODB.Recordset
    Dim codCategoria As String, codSottoCategoria As String, codCer As String
    Dim codSubCer  As String
    
    On Error GoTo getTotaleMovimentiNelPeriodoPerCERERR
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codSubCer = Right$(codCer, 2)
    codCer = Mid$(codCer, 5, 2)
    
    
    If gQtaMaxGiornaliera Then
    
        Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                    "From Movimenti   " & _
                    "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                    " and Movimenti.DataMovimento Between Date()+'00.00.00' And Now() AND  Movimenti.CodImmobile=" & codImmobile & _
                    " AND Movimenti.CodCategoria='" & codCategoria & "' AND Movimenti.CodSottoCategoria='" & codSottoCategoria & "'  AND Movimenti.CodCer='" & codCer & "'")
    Else
        Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                "From Movimenti   " & _
                "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                " and Movimenti.DataMovimento Between  (date() - weekday(date(),0) +1 ) And Now() AND  Movimenti.CodImmobile=" & codImmobile)
                '" AND Movimenti.CodCategoria='" & codCategoria & "' AND Movimenti.CodSottoCategoria='" & codSottoCategoria & "'  AND Movimenti.CodCer='" & codCer & "'")

    End If

    
    getTotaleMovimentiNelPeriodoPerCER = 0
    If Not rs.EOF Then
        getTotaleMovimentiNelPeriodoPerCER = 0
        If Not IsNull(rs(0)) Then
            getTotaleMovimentiNelPeriodoPerCER = rs(0)
        End If
        rs.Close
    End If
    Exit Function
getTotaleMovimentiNelPeriodoPerCERERR:
        ScriviLog TPL_ERRORE, App.Title, "getTotaleMovimentiNelPeriodoPerCER", Err.Description
    
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case Index
        Case TXT_COD_TRASPORTATORE, TXT_COD_UTENTE
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
        
        Case TXT_QUANTITA
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete _
                And KeyAscii <> Asc(",") Then
                KeyAscii = 0
            Else
                mPesoLettoDaBilancia = False
                msgPesa.PesoStabile = True
            End If
    End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
    On Error GoTo txt_LostFocusERR

    If (Index = TXT_COD_UTENTE) Then caricaDatiUtente txt(TXT_COD_UTENTE)
    If (Index = TXT_COD_TRASPORTATORE) Then caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE)
    Exit Sub
txt_LostFocusERR:
    ScriviLog TPL_ERRORE, App.Title, "txt_LostFocus", Err.Description

End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    Select Case Index
        Case TXT_QUANTITA, TXT_COD_UTENTE
            If txt(Index) <> "" And Not IsNumeric(txt(Index)) Then
                MsgBox "Il valore inserito deve essere numerico", vbExclamation + vbOKOnly, App.Title
                Cancel = True
            End If
    End Select
End Sub

Private Sub txtBadge_KeyPress(KeyAscii As Integer)
On Error GoTo txtBadge_KeyPressERR

    'se ricevo badge da lettore alla ricezione del VBCRLF
    If Not getCurrentPesa.BadgeInStringaPesata Then
        'avvio la lettura peso in automatico
        If KeyAscii = Asc(vbCr) And Trim$(txtBadge.Text) <> vbNullString Then
            If Len(Trim$(txtBadge.Text)) = gLunghezzaBadge Then
                ImpostaDatiUtenteDaBadge txtBadge.Text
                If cmd(CMD_LETTURAQUANTITA).Caption <> "&Annulla" Then cmd_Click (CMD_LETTURAQUANTITA)
            Else
                txtBadge.Text = vbNullString
            End If
        End If
    End If
    Exit Sub
txtBadge_KeyPressERR:
    ScriviLog TPL_ERRORE, App.Title, "txtBadge_KeyPress", Err.Description
    
End Sub

Private Sub txtBadgeAccessi_KeyPress(KeyAscii As Integer)
On Error GoTo txtBadgeAccessi_KeyPressERR

'se ricevo badge da lettore alla ricezione del VBCRLF
    If KeyAscii = Asc(vbCr) And Trim$(txtBadgeAccessi.Text) <> vbNullString Then
       PlexaNetKX1_OnCodeEnter 2, 100, 1, txtBadgeAccessi.Text
       txtBadgeAccessi.Text = vbNullString
    End If
    Exit Sub
txtBadgeAccessi_KeyPressERR:
    ScriviLog TPL_ERRORE, App.Title, "txtBadgeAccessi_KeyPress", Err.Description
    
 End Sub

Private Sub TxtMov_Change(Index As Integer)
    Dim cmbItem As ComboItem
    Dim sWhere As String
    
    On Error GoTo TxtMov_ChangeERR
    
    If Index = TXT_FINDCODCERMOV Then
    
        caricaComboRifiuti ImgCombo2, Combo2(1), False
        'mi posiziono sul primo rifiuto
        For Each cmbItem In ImgCombo2.ComboItems
            If (Mid$(cmbItem.key, 5, 2) <> "00") And Len(cmbItem.key) >= 6 Then
                cmbItem.Selected = True
                Exit Sub
            End If
        Next
    End If
    Exit Sub
TxtMov_ChangeERR:
    ScriviLog TPL_ERRORE, App.Title, "TxtMov_Change", Err.Description
    
End Sub

Private Sub TxtMov_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case Index
        Case 0, 2
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub TxtMov_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    Select Case Index
        Case 0
            If TxtMov(Index) <> "" And Not IsNumeric(TxtMov(Index)) Then
                MsgBox "Il valore inserito deve essere numrico", vbExclamation + vbOKOnly, App.Title
                Cancel = True
            End If
    End Select

End Sub


' ----------------- START Sestione sbarra automatica
Private Sub PlexaNetKX1_OnConnect()
    On Error Resume Next
    MDIMain.SbarraCollegata True, "Sbarra"
End Sub

Private Sub PlexaNetKX1_OnDisconnect()
    On Error Resume Next
    MDIMain.SbarraCollegata False, "Sbarra"
End Sub


Private Function ImpostaDatiUtenteDaCodFiscalePIVa(Barcode As String, ReperisciUtente As Boolean, scaduto As Boolean) As Boolean
    Dim rs As ADODB.Recordset
    
    On Error GoTo ImpostaDatiUtenteDaCodFiscalePIVaERR
    
    Set rs = gDBC.Execute("select Badge.CodBadge, Badge.DataBadgeFine from Badge inner join " & _
      "LegameBarcodeBadge as leg on  Badge.codBadge = leg.CodBadge " & _
      "where leg.Barcode ='" & Barcode & "' or leg.codBadge ='" & left$(Barcode, gLunghezzaBadge) & "'")
    
    
    ImpostaDatiUtenteDaCodFiscalePIVa = False
    scaduto = False
    
    If Not rs.EOF Then
        If ReperisciUtente Then ImpostaDatiUtenteDaBadge rs!codBadge
        scaduto = rs!DataBadgeFine < Now()
        'se il badge risulta scaduto, tratto il cliente come se fosse un ospite
        If (Not scaduto) Then
            mlastBadgeFromDeviceAccess = rs!codBadge
            mIsOspiteFromDeviceAccess = False
            ImpostaDatiUtenteDaCodFiscalePIVa = True
        End If
    End If
    
   
    
    If Not rs.EOF And Not ReperisciUtente Then
        'verfico se l'utenza è di tipo non domestica
        Set rs = gDBC.Execute("SELECT " & _
                    "exists(SELECT 1 " & _
                    "FROM Badge INNER JOIN Immobili ON Badge.CodComune = Immobili.CodComune AND Badge.CodUtente = Immobili.CodUtente " & _
                    "where Immobili.TipoUtenza = 'N' And Badge.codBadge = legame1.codBadge And Badge.codComune = legame1.codComune) as UND " & _
                    "from  LegameBarcodeBadge as legame1 where Barcode ='" & Barcode & "' or codBadge ='" & left$(Barcode, gLunghezzaBadge) & "'")
        
        

        mIsUtenteDomestico = CBool(rs!UND.value) = False
        
        If Not mIsUtenteDomestico Then
            PlexaNetKX1.UpdatePulser 1, 3, 1, 30, 1
        End If
        
        'ShowMessage WarningUtenzaNonDomestica, IIf(CBool(rs!UND.Value) = True, "UTENZA NON DOMESTICA", vbNullString)
    End If
    
    If Not rs Is Nothing Then
        If rs.state = adStateOpen Then rs.Close
    End If
    
    Exit Function
    
ImpostaDatiUtenteDaCodFiscalePIVaERR:
    mlastBadgeFromDeviceAccess = ""
    mIsOspiteFromDeviceAccess = False
    ScriviLog TPL_ERRORE, App.Title, "ImpostaDatiUtenteDaCodFiscalePIVa", Err.Description

End Function




Private Sub PlexaNetKX1_OnCodeEnter(ByVal Address As Long, ByVal Resource As Long, ByVal Technology As Long, ByVal value As String)
    Dim invioComandoApertura As Boolean
    Dim scaduto As Boolean
    
    On Error GoTo PlexaNetKX1_OnCodeEnterERR
    
    invioComandoApertura = False
    'ricezione barcode
    If Address = 2 And m_noEvent = False Then
        If mlastBadgeFromDeviceAccess = vbNullString And lblErrorMaxUtentiNelIsola.Visible = False Then
            If Not ImpostaDatiUtenteDaCodFiscalePIVa(value, False, scaduto) Then
                'verifico se già in isola
                Dim ret As typeForzatura
                ret = ForzaEntrataUtenteGiaInIsola(value)
                If ret = NonPresente Then
                    
                    m_noEvent = True
                    Beep
                    If MsgBox("Attenzione !" & vbCrLf & vbCrLf & "L'utente " & value & " non è riconosciuto." & IIf(scaduto, vbCrLf & "Il badge risulta scaduto" & vbCrLf, "") & vbCrLf & vbCrLf & "Si vuole Aprire la sbarra?", vbYesNo + vbQuestion) = vbYes Then
                        'Invio Comando Di Apertura
                        invioComandoApertura = True
                        mlastBadgeFromDeviceAccess = value
                        mIsOspiteFromDeviceAccess = True
                        mIsUtenteDomestico = True
                    End If
                    m_noEvent = False
                ElseIf ret = Forzato Then
                    invioComandoApertura = True
                End If
            ElseIf ForzaEntrataUtenteGiaInIsola(mlastBadgeFromDeviceAccess) <> NonForzatoPresente Then
                 invioComandoApertura = True
            End If
            
            
            If invioComandoApertura Then
                'invio comando apertura
                PlexaNetKX1.UpdatePulser 1, 1, 1, 30, 1
                TrafficLight Verde
            End If
        End If
        
    End If
    Exit Sub
PlexaNetKX1_OnCodeEnterERR:
    invioComandoApertura = False
    mlastBadgeFromDeviceAccess = ""
    mIsOspiteFromDeviceAccess = False
    mIsUtenteDomestico = False
    TrafficLight Rosso
    
    ScriviLog TPL_ERRORE, App.Title, "PlexaNetKX1_OnCodeEnter", Err.Description
    
End Sub

Private Function ForzaEntrataUtenteGiaInIsola(badge As String) As typeForzatura
On Error GoTo ForzaEntrataUtenteGiaInIsolaERR

    ForzaEntrataUtenteGiaInIsola = NonPresente
    If GestioneAccessi(badge, False, False, False, UtenteGiaRilevato) Then
        If MsgBox("Attenzione l'utente " & badge & " risulta già presente in piattaforma." & vbCrLf & "Si vuole Aprire la sbarra?", vbYesNo + vbQuestion) = vbYes Then
            ForzaEntrataUtenteGiaInIsola = Forzato
            mlastBadgeFromDeviceAccess = badge
        Else
            mlastBadgeFromDeviceAccess = vbNullString
            mIsOspiteFromDeviceAccess = False
            ForzaEntrataUtenteGiaInIsola = NonForzatoPresente
        End If
    ElseIf GestioneAccessi(badge, False, False, False, UtenteGiaRilevatoInGiornata) Then
        If MsgBox("Attenzione l'utente " & badge & " risulta già essere stato nella giornata odierna, all'interno della piattaforma." & vbCrLf & "Si vuole Aprire la sbarra?", vbYesNo + vbQuestion) = vbYes Then
            ForzaEntrataUtenteGiaInIsola = Forzato
            mlastBadgeFromDeviceAccess = badge
        Else
            mlastBadgeFromDeviceAccess = vbNullString
            mIsOspiteFromDeviceAccess = False
            ForzaEntrataUtenteGiaInIsola = NonForzatoPresente
        End If
    End If
    Exit Function
ForzaEntrataUtenteGiaInIsolaERR:
    ScriviLog TPL_ERRORE, App.Title, "ForzaEntrataUtenteGiaInIsola", Err.Description

End Function

'Gestione Ingressi Isola
Private Function GestioneAccessi(CodBarcode As String, ospite As Boolean, utenteDomestico As Boolean, richiestaPeso As Boolean, richiesta As typeAccessiRichiesta)
    Dim cmd As ADODB.Command
    
    On Error GoTo GestioneAccessiERR

    
    Set cmd = CreateObject("ADODB.command")
    Set cmd.ActiveConnection = gDBC
    
    Select Case richiesta
        Case typeAccessiRichiesta.Ingresso
            cmd.CommandText = "insert into RilevazioneAccessi (CodBadge,CodComune,DataOraIngresso,DataOraUscita,Stato,Ospite,UtenzaDomestica,DaPesare) " & _
                          "values ('" & CodBarcode & "','" & GetCurrentComune!codistat.value & "',now,null,'2'," & FormatFLD(ospite, adBoolean) & "," & FormatFLD(utenteDomestico, adBoolean) & "," & FormatFLD(richiestaPeso, adBoolean) & ")"
            cmd.Execute
        Case typeAccessiRichiesta.Uscita
            cmd.CommandText = "Update  RilevazioneAccessi Set stato='0',DataOraUscita=now " & _
                          "where  DataOraIngresso in (SELECT   MIN(DataOraIngresso) FROM RilevazioneAccessi Where DataOraUscita is null and DaPesare=false)"
            cmd.Execute
        Case typeAccessiRichiesta.Pulizia
            cmd.CommandText = "update RilevazioneAccessi  set stato='0', DataOraUscita=now where DataOraIngresso<=#" & Format$(DateAdd("n", -gRecycleTimeUsersIsland, Now), "yyyy-mm-dd hh:mm:ss") & "# and DataOraUscita is null"
            cmd.Execute
        Case typeAccessiRichiesta.UtenteGiaRilevato
            cmd.CommandText = "select 1 as IsInside from RilevazioneAccessi where (CodBadge ='" & CodBarcode & _
                        "' and CodComune='" & GetCurrentComune!codistat.value & "' and DataOraUscita is null)"
            Dim rs As ADODB.Recordset
            Set rs = cmd.Execute
            If Not rs.EOF Then
                GestioneAccessi = CBool(rs!IsInside.value)
            Else
                GestioneAccessi = False
            End If
        Case typeAccessiRichiesta.UtenteGiaRilevatoInGiornata
         cmd.CommandText = "select 1 as IsInside from RilevazioneAccessi where (CodBadge ='" & CodBarcode & _
                     "' and CodComune='" & GetCurrentComune!codistat.value & _
                     "' and DataOraIngresso>#" & Format$(Date, "yyyy-mm-dd hh:mm:ss") & "# and DataOraUscita<now )"
         Set rs = cmd.Execute
         If Not rs.EOF Then
             GestioneAccessi = CBool(rs!IsInside.value)
         Else
             GestioneAccessi = False
         End If
        Case typeAccessiRichiesta.DisponibileDopoPesata
            cmd.CommandText = "update RilevazioneAccessi  set  DaPesare=False where CodBadge='" & CodBarcode & "'"
            cmd.Execute
    End Select
    
    
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Function
GestioneAccessiERR:
        ScriviLog TPL_ERRORE, App.Title, "GestioneAccessi", Err.Description
    
End Function


Private Sub PlexaNetKX1_OnDigitalLevelRead(ByVal Address As Long, ByVal Resource As Long, ByVal value As Long)
    On Error GoTo PlexaNetKX1_OnDigitalLevelReadERR
    If Address = 1 Then
        Select Case Resource
            Case 1 'fotocellula
                If value = 0 And mlastBadgeFromDeviceAccess <> vbNullString Then
                    GestioneAccessi mlastBadgeFromDeviceAccess, mIsOspiteFromDeviceAccess, mIsUtenteDomestico, mIsUtenteDomestico = False, Ingresso
                    mlastBadgeFromDeviceAccess = vbNullString
                    mIsOspiteFromDeviceAccess = False
                    mIsUtenteDomestico = False
                    TrafficLight Rosso
                End If
             Case 2   'Uscita
                If value = 0 Then GestioneAccessi -1, False, False, False, Uscita
         End Select
     End If
     
     Exit Sub
PlexaNetKX1_OnDigitalLevelReadERR:
    mlastBadgeFromDeviceAccess = vbNullString
    mIsOspiteFromDeviceAccess = False
    mIsUtenteDomestico = False
    TrafficLight Rosso
        ScriviLog TPL_ERRORE, App.Title, "PlexaNetKX1_OnDigitalLevelRead", Err.Description
    
End Sub


Private Sub PlexaNetKX1_OnError(ByVal errorCode As Long, ByVal ErrorDesc As String)
Select Case errorCode
    Case 201
        MsgBox "Attenzione, rilevatore accessi non rilevato, verificare i collegamenti", vbCritical + vbOKOnly, "Rilevatore Accessi"
End Select
End Sub

Private Sub TrafficLight(state As typeTrafficlightState)
    
    imgTrafficLightGreen.Visible = False
    imgTrafficLightRed.Visible = False
    On Error Resume Next
    If state = Rosso Then
        'invio segnale di rosso al semaforo
        imgTrafficLightRed.Visible = True
        PlexaNetKX1.UpdatePulser 1, 4, 0, 0, 1
    Else
        'invio segnale di verde al semaforo
        imgTrafficLightGreen.Visible = True
        PlexaNetKX1.UpdatePulser 1, 4, 1, 0, 1
    End If
    
End Sub

' ----------------- END Qestione Sbarra automatica
