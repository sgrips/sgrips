VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditDatiTecnici 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dati Tecnici del comune"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmEditDatiTecnici.frx":0000
      DataSource      =   "ADO1"
      Height          =   315
      Left            =   3435
      TabIndex        =   56
      Top             =   7020
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descrizione"
      BoundColumn     =   "NumeroComponenti"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   2235
      Top             =   7920
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
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
      Caption         =   "ADO1"
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
      Left            =   3435
      TabIndex        =   53
      Top             =   6570
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
      Left            =   3435
      TabIndex        =   50
      Top             =   4170
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
      Index           =   13
      Left            =   3435
      TabIndex        =   49
      Top             =   4650
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
      Left            =   8835
      TabIndex        =   47
      Top             =   4170
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
      Left            =   8820
      TabIndex        =   44
      Top             =   5100
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
      Index           =   7
      Left            =   8820
      TabIndex        =   43
      Top             =   4650
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Importazione Dati"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   42
      Top             =   7950
      Width           =   1455
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
      Index           =   19
      Left            =   8820
      TabIndex        =   40
      Top             =   7020
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
      Index           =   18
      Left            =   8820
      TabIndex        =   38
      Top             =   6540
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
      Index           =   16
      Left            =   8820
      TabIndex        =   35
      Top             =   5580
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
      Index           =   17
      Left            =   8820
      TabIndex        =   34
      Top             =   6060
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
      Index           =   15
      Left            =   3435
      TabIndex        =   32
      Top             =   7470
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
      Index           =   14
      Left            =   3435
      TabIndex        =   30
      Top             =   5130
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
      Left            =   3435
      TabIndex        =   28
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   8
      Top             =   7950
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   7950
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
      Index           =   6
      Left            =   8835
      TabIndex        =   6
      Top             =   3705
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
      Left            =   8820
      TabIndex        =   5
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
      Index           =   4
      Left            =   8805
      TabIndex        =   4
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
      Index           =   3
      Left            =   3435
      TabIndex        =   3
      Top             =   6075
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
      Left            =   3435
      TabIndex        =   2
      Top             =   5595
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
      Left            =   3435
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
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
         Index           =   102
         Left            =   1320
         TabIndex        =   27
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
         Index           =   204
         Left            =   240
         TabIndex        =   26
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
         Index           =   200
         Left            =   8640
         TabIndex        =   25
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
         Index           =   201
         Left            =   2400
         TabIndex        =   24
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
         Index           =   205
         Left            =   240
         TabIndex        =   23
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
         Index           =   203
         Left            =   240
         TabIndex        =   22
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
         Index           =   100
         Left            =   240
         TabIndex        =   17
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
         Index           =   101
         Left            =   1320
         TabIndex        =   16
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
         Index           =   103
         Left            =   1320
         TabIndex        =   15
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
         Index           =   104
         Left            =   3000
         TabIndex        =   14
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
         Index           =   105
         Left            =   9240
         TabIndex        =   13
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
      Left            =   3435
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nr. Componenti Seconda Casa"
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
      Left            =   240
      TabIndex        =   55
      Top             =   7020
      Width           =   2775
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nr. Ruoli Seconda Casa"
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
      Left            =   240
      TabIndex        =   54
      Top             =   6570
      Width           =   2160
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Rifiuti Utenze Attività:"
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
      Left            =   240
      TabIndex        =   52
      Top             =   4170
      Width           =   2610
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Rifiuti Utenze Domestiche:"
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
      Left            =   240
      TabIndex        =   51
      Top             =   4650
      Width           =   3105
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gettito Totale"
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
      Left            =   5610
      TabIndex        =   48
      Top             =   4185
      Width           =   1200
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot.Gettito da Utenze Domestiche"
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
      Left            =   5595
      TabIndex        =   45
      Top             =   4665
      Width           =   3000
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot.Gettito da Utenze Attività"
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
      Left            =   5595
      TabIndex        =   46
      Top             =   5100
      Width           =   2505
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo Kg  medio Quota Variabile"
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
      Index           =   19
      Left            =   5595
      TabIndex        =   41
      Top             =   7020
      Width           =   2910
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo Kg medio Quota Fissa"
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
      Left            =   5595
      TabIndex        =   39
      Top             =   6540
      Width           =   2550
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo Totale Kg rifiuti gestiti"
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
      Left            =   5595
      TabIndex        =   37
      Top             =   5580
      Width           =   2445
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo Kg rifiuti da Raccolta Diff."
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
      Left            =   5595
      TabIndex        =   36
      Top             =   6060
      Width           =   2760
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superf..Tot. Iscritti a Ruolo (mq)"
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
      Left            =   240
      TabIndex        =   33
      Top             =   7470
      Width           =   2760
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nr. Ruoli Totali"
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
      Left            =   240
      TabIndex        =   31
      Top             =   5130
      Width           =   1320
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Rifiuti a Smaltimento"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   2580
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superf.Tot. non Iscritti"
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
      Left            =   5625
      TabIndex        =   21
      Top             =   3735
      Width           =   1890
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superf.Tot. Utenza Attività"
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
      Left            =   5625
      TabIndex        =   20
      Top             =   3240
      Width           =   2280
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superf.Tot. Utenza Domestica"
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
      Index           =   4
      Left            =   5610
      TabIndex        =   19
      Top             =   2760
      Width           =   2670
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nr. Ruoli Utenza Attività"
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
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   6075
      Width           =   2070
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nr. Ruoli Residenti Domestici"
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
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   5595
      Width           =   2610
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Rifiuti da Raccolta Diff."
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
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantità Rifiuti Totali"
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditDatiTecnici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lvw As ListView
Private OldTxtValue As String

'ID text box
Private Const Q_RIFIUTI_TOT_ID = 0
Private Const Q_RIFIUTI_DA_RACC_DIFF_ID = 1
Private Const NR_RUOLI_UD_ID = 2
Private Const NR_RUOLI_UA_ID = 3
Private Const SUP_TOT_UD_ID = 4
Private Const SUP_TOT_UA_ID = 5
Private Const SUP_TOT_NON_ISCR_RUOLO_ID = 6
Private Const TOT_GETTITODAUTENZADOMESTICA = 7
Private Const TOT_GETTITODAUTENZAATTIVA = 8
Private Const Q_RIFIUTI_UA_ID = 9
Private Const NR_RUOLI_2CASA_ID = 10

Private Const GETTITO_ID = 11
Private Const Q_RIFIUTI_SMALTITI_ID = 12
Private Const Q_RIFIUTI_UD_ID = 13
Private Const NR_RUOLI_TOT_ID = 14
Private Const SUP_TOT_ISCR_RUOLO_ID = 15
Private Const COSTO_TOT_KG_RIFIUTI_GESTITI_ID = 16
Private Const COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID = 17
Private Const COSTO_KG_RIFIUTI_QUOTA_FISSA_ID = 18
Private Const COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID = 19



'Command Button
Private Const CMD_IMPORTAZIONE = 2
Private Const CMD_CONFERMA = 1
Private Const CMD_ANNULLA = 0
'
Private mImportazione As Long

Dim UltimoTxtValido As Long

Public Sub Edita(plvw As ListView)
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    
    
    UltimoTxtValido = NR_RUOLI_2CASA_ID
    
    Set lvw = plvw
    
    lblCampi(100) = lvw.SelectedItem.SubItems(1)
    lblCampi(101) = lvw.SelectedItem.Text
    lblCampi(102) = lvw.SelectedItem.SubItems(3)
    lblCampi(103) = lvw.SelectedItem.SubItems(4)
    lblCampi(104) = lvw.SelectedItem.SubItems(5)
    lblCampi(105) = lvw.SelectedItem.SubItems(2)
    
    Set rs = DatiCostiComune(lblCampi(101), lblCampi(105))
    If rs.EOF Then
        MsgBox "Attenzione non sono stati ancora inseriti i costi comunali per questo comune", vbCritical, App.Title
        
        Unload Me
        
        frmDatiTecnici.Visible = True
    Else
        For i = 0 To txt.Count - 1
            txt(i) = 0
            If i < txt.Count + 1 - UltimoTxtValido Then
                txt(i) = CDbl(lvw.SelectedItem.SubItems(i + 6))
            End If
            
            txt(i).TabIndex = i
            txt(i).Locked = IIf(lblCampi(i).ForeColor = FORECOLOR_CAMPI_CALC, True, False)
            txt(i).BackColor = IIf(lblCampi(i).ForeColor = FORECOLOR_CAMPI_CALC, BACKCOLOR_CAMPI_CALC, vbWhite)
            txt(i).TabStop = Not txt(i).Locked
            txt(i).MaxLength = 15
            txt(i).Tag = CDbl(txt(i))
        Next i
        
    
        If CLng(lvw.SelectedItem.SubItems(NR_RUOLI_2CASA_ID + 1 + 6)) > 0 Then DataCombo1.BoundText = CLng(lvw.SelectedItem.SubItems(NR_RUOLI_2CASA_ID + 1 + 6))
        txt_LostFocus 0
    
    
        Me.Show vbModal
    End If
End Sub



Private Sub Command1_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case CMD_CONFERMA
            'conferma
            lvw.SelectedItem.SubItems(1) = lblCampi(100)
            lvw.SelectedItem.Text = lblCampi(101)
            lvw.SelectedItem.SubItems(3) = lblCampi(102)
            lvw.SelectedItem.SubItems(4) = lblCampi(103)
            lvw.SelectedItem.SubItems(5) = lblCampi(104)
            lvw.SelectedItem.SubItems(2) = lblCampi(105)
            
            For i = 0 To txt.Count - UltimoTxtValido
                lvw.SelectedItem.SubItems(i + 6) = CDbl(txt(i).Text)
            Next i
            
            If DataCombo1.Enabled Then
                 lvw.SelectedItem.SubItems(UltimoTxtValido + 1 + 6) = DataCombo1.BoundText
            Else
                lvw.SelectedItem.SubItems(UltimoTxtValido + 1 + 6) = 0
            End If
                        
            Unload Me
        Case CMD_ANNULLA
            'annullA
            Unload Me
    End Select
        
    frmDatiTecnici.Visible = True
End Sub

Private Sub DataCombo1_GotFocus()
    Info "Selezionare il numero dei componenti famiglia attribuiti dal regolamento comunale per le seconde case"
End Sub

Private Sub Form_Load()
    CentraME Me
    frmDatiTecnici.Visible = False


    '
    lblCampi(Q_RIFIUTI_UA_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(Q_RIFIUTI_UD_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(Q_RIFIUTI_SMALTITI_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(NR_RUOLI_UA_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(NR_RUOLI_TOT_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(SUP_TOT_ISCR_RUOLO_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(COSTO_TOT_KG_RIFIUTI_GESTITI_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(COSTO_KG_RIFIUTI_QUOTA_FISSA_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(SUP_TOT_UA_ID).ForeColor = FORECOLOR_CAMPI_CALC
    lblCampi(GETTITO_ID).ForeColor = FORECOLOR_CAMPI_CALC
    
    
    On Error Resume Next
    ADO1.Recordset.Close
    ADO1.CommandType = adCmdTable
    ADO1.RecordSource = "NumeroComponentiFamiglia"
    
    ADO1.ConnectionString = gDBC
    ADO1.Refresh

    DataCombo1.BoundText = 1
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Info vbNullString
End Sub

Private Sub txt_GotFocus(Index As Integer)
    

    'Attiva il bottone Importazione dati da MDB del comune
    mImportazione = Index
    
'    If Index <> NR_RUOLI_UD_ID And Index <> NR_RUOLI_UA_ID Then
        Command1(CMD_IMPORTAZIONE).Enabled = False
'    Else
'        Command1(CMD_IMPORTAZIONE).Enabled = True
'    End If
    
    
    Select Case Index
        Case Q_RIFIUTI_TOT_ID
            Info "Quantità rifiuti totali in kg/Anno soggetti a privativa"
        Case Q_RIFIUTI_DA_RACC_DIFF_ID, Q_RIFIUTI_SMALTITI_ID
            Info lblCampi(Index) & " in kg/Anno"
        Case Q_RIFIUTI_UD_ID
            Info "Quantità dei rifiuti in kg/Anno da utenze domestiche, determinata come differenza tra la quantità dei rifiuti totali e quella utenze non domestiche"
        Case Q_RIFIUTI_UA_ID
            Info "Quantità dei rifiuti in kg/Anno da utenze non domestiche, determinata utilizzando i coefficienti Kd (kg/mq Anno) e le superfici delle utenze non domestiche"
 
        Case NR_RUOLI_UD_ID
            Info "Numero ruoli utenze domestiche iscritti per l'anno di riferimento"
        Case NR_RUOLI_UA_ID
            Info "Numero ruoli utenze non domestiche iscritti per l'anno di riferimento"
        Case NR_RUOLI_TOT_ID
            Info "Numero ruoli iscritti per l'anno di riferimento"
        Case NR_RUOLI_2CASA_ID
            Info "Vengono determinati automaticamente quando il numero dei ruoli delle utenze domestiche è superiore ai nuclei familiari"
        Case SUP_TOT_UD_ID, SUP_TOT_UA_ID, SUP_TOT_NON_ISCR_RUOLO_ID, SUP_TOT_ISCR_RUOLO_ID
            Info lblCampi(Index) & " in (mq)"
        Case TOT_GETTITODAUTENZADOMESTICA, TOT_GETTITODAUTENZAATTIVA
            Info lblCampi(Index) & " in €/Anno. Dato utilizzato nel calcolo della tariffa con il metodo TRANSITORIO"
        Case GETTITO_ID
            Info lblCampi(Index) & " in €/Anno"
        Case COSTO_TOT_KG_RIFIUTI_GESTITI_ID, COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID, _
             COSTO_KG_RIFIUTI_QUOTA_FISSA_ID, COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID
            Info lblCampi(Index) & " in €/kg"
    End Select
    
    
    If txt(Index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub
    
    'Imposto il numero MAX di caratteri in EDIT MODE
    txt(Index) = CDbl(txt(Index))
    txt(Index).MaxLength = 12
    
    
    DoEvents

    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index))
    OldTxtValue = txt(Index)

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
    
    txt(Index).MaxLength = 15
    
    RicalcolaCampi (Index)
    
    If txt(NR_RUOLI_2CASA_ID) > 0 Then
        DataCombo1.Enabled = True
    Else
        DataCombo1.Enabled = False
    End If
    
End Sub


Sub RicalcolaCampi(Index As Integer)
    Dim rs As Recordset
    Dim i As Long
    Dim sParAbitanti As String
    
    'Inizializzo valori
    txt(Q_RIFIUTI_SMALTITI_ID) = 0
    txt(NR_RUOLI_TOT_ID) = 0
    txt(SUP_TOT_ISCR_RUOLO_ID) = 0
    txt(COSTO_TOT_KG_RIFIUTI_GESTITI_ID) = 0
    txt(COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID) = 0
    txt(COSTO_KG_RIFIUTI_QUOTA_FISSA_ID) = 0
    txt(COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID) = 0

    'Superficie totale delle utenze attività e numero ruoli attività
    Set rs = gDBC.Execute("SELECT sum(NumRuoli) as TotRuoli,sum(Superficie) as SupTot FROM CategorieAttivitaComunali where CodComune = '" & lblCampi(101) & "' and AnnoRiferimento = '" & lblCampi(105) & "'")
    If Not rs.EOF Then
        txt(SUP_TOT_UA_ID) = Format$(IIf(IsNull(rs("SupTot")), 0, rs("SupTot")), FORMATVALUTA)
        txt(NR_RUOLI_UA_ID) = Format$(IIf(IsNull(rs("TotRuoli")), 0, rs("TotRuoli")), FORMATVALUTA)
    End If
   
    
    'Costi del Comune
    Set rs = DatiCostiComune(lblCampi(101), lblCampi(105))
    txt(GETTITO_ID) = Format$(CDbl(rs("GETTITO")), FORMATVALUTA)
    
    'Calcolo Gettito Utenze Domestiche e Attivita
    Select Case Index
        Case TOT_GETTITODAUTENZAATTIVA
            txt(TOT_GETTITODAUTENZADOMESTICA) = Format$(CDbl(txt(GETTITO_ID)) - CDbl(txt(TOT_GETTITODAUTENZAATTIVA)), FORMATVALUTA)
        Case TOT_GETTITODAUTENZADOMESTICA
            txt(TOT_GETTITODAUTENZAATTIVA) = Format$(CDbl(txt(GETTITO_ID)) - CDbl(txt(TOT_GETTITODAUTENZADOMESTICA)), FORMATVALUTA)
        Case Else
            If CDbl(txt(TOT_GETTITODAUTENZADOMESTICA)) + txt(TOT_GETTITODAUTENZAATTIVA) <> CDbl(txt(GETTITO_ID)) Then
                txt(TOT_GETTITODAUTENZADOMESTICA) = 0
                txt(TOT_GETTITODAUTENZAATTIVA) = 0
            End If
    End Select

    
    txt(Q_RIFIUTI_UA_ID) = Format(QuantitaRifiutiDaUtenzeAttivita(lblCampi(101), lblCampi(105)), FORMATVALUTA)
    txt(Q_RIFIUTI_UD_ID) = Format(CDbl(txt(Q_RIFIUTI_TOT_ID)) - CDbl(txt(Q_RIFIUTI_UA_ID)), FORMATVALUTA)
    
    If rs.EOF Then
        MsgBox "Occorre inserire i Costi del comune prima di procedere con quelli Tecnici", vbExclamation + vbOKOnly, App.Title
    Else
        If Not rs.EOF Then
            'Dati Tecnici
            txt(Q_RIFIUTI_SMALTITI_ID) = Format$(CDbl(txt(Q_RIFIUTI_TOT_ID)) - CDbl(txt(Q_RIFIUTI_DA_RACC_DIFF_ID)), FORMATVALUTA)
            txt(NR_RUOLI_TOT_ID) = Format$(CDbl(txt(NR_RUOLI_UD_ID)) + CDbl(txt(NR_RUOLI_UA_ID)) + CDbl(txt(NR_RUOLI_2CASA_ID)), FORMATSTANDARD)
            txt(SUP_TOT_ISCR_RUOLO_ID) = Format$(CDbl(txt(SUP_TOT_UD_ID)) + CDbl(txt(SUP_TOT_UA_ID)), FORMATVALUTA)
            
            'Calcolo Costo Rifiuti
            If CDbl(txt(Q_RIFIUTI_TOT_ID)) <> 0 Then txt(COSTO_TOT_KG_RIFIUTI_GESTITI_ID) = Format$(rs("Tot_servizio") / CDbl(txt(Q_RIFIUTI_TOT_ID)), FORMATVALUTA)
            If CDbl(txt(Q_RIFIUTI_DA_RACC_DIFF_ID)) <> 0 Then txt(COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID) = Format$(rs("CRD") / CDbl(txt(Q_RIFIUTI_DA_RACC_DIFF_ID)), FORMATVALUTA)
            If CDbl(txt(Q_RIFIUTI_TOT_ID)) <> 0 Then txt(COSTO_KG_RIFIUTI_QUOTA_FISSA_ID) = Format$(rs("ParteFissa") / CDbl(txt(Q_RIFIUTI_TOT_ID)), FORMATVALUTA)
            If CDbl(txt(Q_RIFIUTI_TOT_ID)) <> 0 Then txt(COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID) = Format$(rs("ParteVariabile") / CDbl(txt(Q_RIFIUTI_TOT_ID)), FORMATVALUTA)
        End If
    End If
    
    'Seconde case
    Set rs = DatiAnagraficiComunali(lblCampi(101), lblCampi(105), sParAbitanti)
    If rs.EOF Then
        MsgBox "Occorre inserire i dati Anagrafici del comune prima di procedere con quelli Tecnici", vbExclamation + vbOKOnly, App.Title
    Else
        If CDbl(txt(NR_RUOLI_UD_ID)) > rs("NumNucleoFamiliare") And CDbl(txt(NR_RUOLI_2CASA_ID)) = 0 Then
            txt(NR_RUOLI_2CASA_ID) = CDbl(txt(NR_RUOLI_UD_ID)) - rs("NumNucleoFamiliare")
            txt(NR_RUOLI_UD_ID) = CDbl(txt(NR_RUOLI_UD_ID)) - CDbl(txt(NR_RUOLI_2CASA_ID))
        End If
    End If
    Set rs = Nothing



    For i = 0 To txt.UBound
'        If i <> TOT_GETTITODAUTENZADOMESTICA And _
 '          i <> TOT_GETTITODAUTENZAATTIVA And _
  '         i <> GETTITO_ID And _
   '        i <> COSTO_TOT_KG_RIFIUTI_GESTITI_ID And _
    '       i <> COSTO_KG_RIFIUTI_DA_RACC_DIFF_ID And _
     '      i <> COSTO_KG_RIFIUTI_QUOTA_FISSA_ID And _
      '      i <> Q_RIFIUTI_UA_ID And _
       '     i <> Q_RIFIUTI_DA_RACC_DIFF_ID And _
        '    i <> Q_RIFIUTI_SMALTITI_ID And _
         '   i <> Q_RIFIUTI_UD_ID And _
          '  i <> Q_RIFIUTI_TOT_ID And _
           'i <> COSTO_KG_RIFIUTI_QUOTA_VARIABILE_ID Then
            
           ' txt(i) = Format$(CDbl(txt(i)), FORMATSTANDARD)
        'Else
            txt(i) = Format$(CDbl(txt(i)), FORMATVALUTA)
        'End If
    Next


End Sub


Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
        'Il valore deve essere numerico
        If Not IsNumeric(txt(Index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txt(Index) = OldTxtValue: Exit Sub

End Sub
