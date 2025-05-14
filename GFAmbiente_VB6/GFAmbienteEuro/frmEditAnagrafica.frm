VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditAnagrafica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anagrafica del comune"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   360
      Left            =   6000
      Top             =   7080
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "GFAmbiente"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmEditAnagrafica.frx":0000
      DataSource      =   "ADO1"
      Height          =   315
      Left            =   5985
      TabIndex        =   44
      Top             =   5865
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "DESCRIZION"
      BoundColumn     =   "PROVINCIA"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   8880
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   8880
      TabIndex        =   13
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   8880
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   8880
      TabIndex        =   11
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   3240
      TabIndex        =   9
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   3240
      TabIndex        =   8
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CheckBox chkProvinciaOttimale 
      Alignment       =   1  'Right Justify
      Caption         =   "Appartenenza ambito territoriale ottimale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10800
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   25
         Left            =   1320
         TabIndex        =   43
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regione:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   24
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   23
         Left            =   8640
         TabIndex        =   41
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   22
         Left            =   2400
         TabIndex        =   40
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodIstat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   21
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodIstat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   24
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   23
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   3
         Left            =   3000
         TabIndex        =   22
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblCampi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "01/01/1998"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   4
         Left            =   9240
         TabIndex        =   21
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia Comunale Ottimale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   6000
      TabIndex        =   37
      Top             =   5520
      Width           =   2580
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superficie Terr. Comunale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   6000
      TabIndex        =   36
      Top             =   4200
      Width           =   2325
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num.Medio Turisti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   6000
      TabIndex        =   35
      Top             =   3720
      Width           =   1620
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Domiciliati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   6000
      TabIndex        =   34
      Top             =   3240
      Width           =   1440
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num.Medio Residenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   6000
      TabIndex        =   33
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emigrati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   360
      TabIndex        =   32
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Immigrati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   360
      TabIndex        =   31
      Top             =   6600
      Width           =   810
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 6 o più componenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   360
      TabIndex        =   30
      Top             =   6120
      Width           =   2880
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 5 componenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   360
      TabIndex        =   29
      Top             =   5640
      Width           =   2400
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 4 componenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   360
      TabIndex        =   28
      Top             =   5160
      Width           =   2400
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 3 componenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   360
      TabIndex        =   27
      Top             =   4680
      Width           =   2400
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 2 componenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   360
      TabIndex        =   26
      Top             =   4200
      Width           =   2400
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia con 1 componente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   2475
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nucleo Familiare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num.Abitanti Residenti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   2025
   End
End
Attribute VB_Name = "frmEditAnagrafica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Lvw As ListView
Dim OldTxtValue As String
Dim PercScartoPop As Long
Dim Popolazione As Long

Public Sub Edita(plvw As ListView)
    Dim i As Integer
    
    ADO1.ConnectionString = gDBC
    

    Set Lvw = plvw
    lblCampi(0) = Lvw.SelectedItem.SubItems(1)
    lblCampi(1) = Lvw.SelectedItem.Text
    lblCampi(4) = Lvw.SelectedItem.SubItems(2)
    lblCampi(2) = Lvw.SelectedItem.SubItems(4)
    lblCampi(3) = Lvw.SelectedItem.SubItems(5)
    lblCampi(25) = Lvw.SelectedItem.SubItems(3)
    
    On Error Resume Next
    ADO1.Recordset.Close
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "Select * from Provincie inner Join Regioni ON Provincie.CodRegione=Regioni.CodRegione Where Regione='" & lblCampi(25) & "'"
    ADO1.Refresh
    
    DataCombo1.BoundText = ADO1.Recordset.fields("Provincia")
    
    For i = 0 To txt.Count - 1
        txt(i) = Lvw.SelectedItem.SubItems(i + 6)
    Next i
    
    If Lvw.SelectedItem.SubItems(20) = "True" Or Lvw.SelectedItem.SubItems(20) = "Vero" Then
        chkProvinciaOttimale.Value = 1
        DataCombo1.BoundText = Lvw.SelectedItem.SubItems(21)
    Else
        chkProvinciaOttimale.Value = 0
    End If
    
    chkProvinciaOttimale_Click
    
    txt_LostFocus (0)
    
    PercScartoPop = GetParamValue(Lvw.SelectedItem.SubItems(2), Lvw.SelectedItem.Text, "Popolazione")
    If PercScartoPop = -1 Then
        LetParamValue Lvw.SelectedItem.SubItems(2), Lvw.SelectedItem.Text, "Popolazione", True, 10
        PercScartoPop = 10
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    Set rs.ActiveConnection = gDBC
    rs.Source = "Select * from Comuni where codIstat='" & Lvw.SelectedItem.Text & "'"
    rs.Open
    Set rs.ActiveConnection = Nothing
    PercScartoPop = rs("Popolazione") * (PercScartoPop / 100)
    Popolazione = rs("Popolazione")
    Set rs = Nothing
    frmEditAnagrafica.Show vbModal
    
End Sub



Private Sub chkProvinciaOttimale_Click()
    If chkProvinciaOttimale.Value = 1 Then
        DataCombo1.Enabled = True
    Else
        DataCombo1.Enabled = False
    End If
End Sub

Private Sub chkProvinciaOttimale_GotFocus()
    Info "Riferimento ART 23 - Gestione dei rifiuti urbani in ambito territoriale ottimale D.Lgs 22/97 e suss."
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 1
            'conferma
            Lvw.SelectedItem.SubItems(1) = lblCampi(0)
            Lvw.SelectedItem.Text = lblCampi(1)
            Lvw.SelectedItem.SubItems(2) = lblCampi(4)
            Lvw.SelectedItem.SubItems(3) = lblCampi(25)
            Lvw.SelectedItem.SubItems(4) = lblCampi(3)
    
            For i = 0 To 13
                 Lvw.SelectedItem.SubItems(i + 6) = CDbl(txt(i).Text)
            Next i
            
            If chkProvinciaOttimale = 1 Then
                Lvw.SelectedItem.SubItems(20) = "True"
                Lvw.SelectedItem.SubItems(21) = DataCombo1.BoundText
            Else
                Lvw.SelectedItem.SubItems(20) = "False"
                Lvw.SelectedItem.SubItems(21) = vbNullString

            End If
            
            Unload Me
        Case 0
            'annullA
            Unload Me
    End Select
        
    frmDatiAnagrafici.Visible = True
    
End Sub

Private Sub Command1_GotFocus(Index As Integer)
    Select Case Index
        Case 1
            Info "Conferma modifiche dati"
        Case 0
            Info "Annulla modifiche dati"

    End Select
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    Info "Selezionare la provincia di apparteneza del comune"

End Sub

Private Sub Form_Activate()
    txt(0).SetFocus
End Sub

Private Sub Form_Load()
    CentraME Me
    frmDatiAnagrafici.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Info vbNullString
End Sub

Private Sub txt_GotFocus(Index As Integer)

    If txt(Index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub

    txt(Index).MaxLength = 12
    txt(Index).Text = CDbl(txt(Index).Text)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index))
    OldTxtValue = txt(Index)
    Select Case Index
        Case 0
            Info "Numero dei residenti attuali"
        Case 1
            Info "Numero nucleo familiare"
        Case 2
            Info "Numero Famiglie con 1 componente"
        Case 3
            Info "Numero Famiglie con 2 componenti"
        Case 4
            Info "Numero Famiglie con 3 componenti"
        Case 5
            Info "Numero Famiglie con 4 componenti"
        Case 6
            Info "Numero Famiglie con 5 componenti"
        Case 7
            Info "Numero Famiglie con 6 o più componenti"
        Case 8
            Info "Numero Immigrati anno di riferimento"
        Case 9
            Info "Numero Emmigrati anno di riferimento"
        Case 10
            Info "Numero medio residenti"
        Case 11
            Info "Numero dei nuclei domiciliati non residenti"
        Case 12
            Info "Numero medio presenze turistiche all'anno"
        Case 13
            Info "Superficie territorio comunale (mq)"
    End Select
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim CancelERR As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Index + 1 > txt.UBound Then
            Index = 0
        Else
            txt_Validate Index, CancelERR
            If Not CancelERR Then txt(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Dim i As Integer
    
    txt(Index).MaxLength = 15
    
    For i = 0 To txt.Count - 1
        If i >= 1 And i <= 7 Then
            txt(i).Text = Format$(CDbl(txt(i).Text), FORMATVALUTA)
        Else
            txt(i).Text = Format$(CDbl(txt(i).Text), FORMATSTANDARD)
        End If
    Next
    

    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    'Il valore deve essere numerico
    
    If Not IsNumeric(txt(Index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txt(Index) = OldTxtValue: Exit Sub

    If Index = 0 Then
        'controllo che la popolazione non sia < o > di una certa % rispetto a quella censita
        If (CLng(txt(Index).Text) > Popolazione + PercScartoPop Or CLng(txt(Index).Text) < Popolazione - PercScartoPop) And txt(Index).Text <> 0 Then
            MsgBox "Attenzione la popolazione può differenziarsi in + o - rispetto a quella censita (" & Popolazione & ") solo di (" & PercScartoPop & ") unita'", vbExclamation, App.Title
            txt(Index).Text = Popolazione
            Cancel = True
        End If
    End If
    
    'componenti familiari non possono essere superiori del nucleo familiare!
    If Index >= 1 And Index < 7 Then
        txt(7) = CDbl(txt(1)) - (CDbl(txt(2)) + CDbl(txt(3)) + CDbl(txt(4)) + CDbl(txt(5)) + CDbl(txt(6)))
        If txt(7) < 0 Then
            Cancel = True
            MsgBox "Attenzione i componenti familiari non possono essere superiori del nucleo familiare", vbInformation + vbOKOnly, App.Title
        End If
    End If
End Sub
