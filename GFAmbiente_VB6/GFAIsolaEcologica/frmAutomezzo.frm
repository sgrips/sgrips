VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAutomezzi 
   Caption         =   "Elenco Automezzi"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   14250
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4140
      Index           =   1
      Left            =   1440
      ScaleHeight     =   4110
      ScaleWidth      =   5505
      TabIndex        =   20
      Top             =   360
      Width           =   5535
      Begin VB.CheckBox chkFlgDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Principale"
         DataField       =   "flgDefault"
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCodSoggetto 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codSoggetto"
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   810
         Width           =   660
      End
      Begin VB.TextBox txtidSoggetto 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codUtente"
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   315
         Width           =   660
      End
      Begin VB.TextBox txtTaraTotale 
         DataField       =   "taraTotale"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   4
         Top             =   3345
         Width           =   1515
      End
      Begin VB.TextBox txtTargaRimorchio 
         DataField       =   "targaRimorchio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         MaxLength       =   11
         TabIndex        =   2
         Top             =   2310
         Width           =   1500
      End
      Begin VB.TextBox txtTarga 
         DataField       =   "targa"
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1800
         Width           =   1560
      End
      Begin VB.TextBox txtDescrizione 
         DataField       =   "descrizione"
         DataMember      =   "cmdAutomezzi"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   0
         Top             =   1305
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Soggetto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   420
         TabIndex        =   28
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ID Soggetto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   585
         TabIndex        =   26
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tara Totale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   645
         TabIndex        =   24
         Top             =   3390
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Targa Rimorchio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   195
         TabIndex        =   23
         Top             =   2355
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Targa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1125
         TabIndex        =   22
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrizione:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   585
         TabIndex        =   21
         Top             =   1320
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   12615
      TabIndex        =   8
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAutomezzo.frx":0000
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   3
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "cmdAutomezzi"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "codUtente"
            Caption         =   "codUtente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "codSoggetto"
            Caption         =   "codSoggetto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "descrizione"
            Caption         =   "Descrizione"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "targa"
            Caption         =   "Targa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "targaRimorchio"
            Caption         =   "Targa Rimorchio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "taraTotale"
            Caption         =   "Tara Totale"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "flgDefault"
            Caption         =   "Principale"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3420,284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1814,74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1814,74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   945,071
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox cmbGrid 
      Height          =   315
      Left            =   3480
      TabIndex        =   39
      Text            =   "Combo1"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3180
      Index           =   2
      Left            =   5520
      ScaleHeight     =   3150
      ScaleWidth      =   7785
      TabIndex        =   29
      Top             =   4440
      Width           =   7815
      Begin VB.ListBox List1 
         Height          =   1815
         Index           =   1
         Left            =   3720
         TabIndex        =   36
         ToolTipText     =   "Tipo operazione Recupero associata"
         Top             =   600
         Width           =   2655
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   0
         Left            =   2640
         ScaleHeight     =   1935
         ScaleWidth      =   735
         TabIndex        =   31
         Top             =   240
         Width           =   735
         Begin VB.CommandButton Command1 
            Caption         =   ">"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Tag             =   "Associa l'elemento selezionato"
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "<<"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Disassocia tutti gli elementi"
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   ">>"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Associa tutti gli elementi"
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "<"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Disassocia l'elemento selezionato"
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Index           =   0
         Left            =   240
         TabIndex        =   30
         ToolTipText     =   "Tipo operazione Recupero da associata"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rifiuti da Associare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rifiuti Associati"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3840
         TabIndex        =   37
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6360
      TabIndex        =   15
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Nuovo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   7680
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   9000
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Chiudi"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   11640
      TabIndex        =   11
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ricerca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10320
      TabIndex        =   12
      Top             =   8040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   3
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   8865
      TabIndex        =   10
      Top             =   4560
      Width           =   8895
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Targa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Targa Rimorchio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrizione"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3360
         TabIndex        =   17
         Top             =   120
         Width           =   1020
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Key             =   "LISTA"
            Object.ToolTipText     =   "Elenco dei trasportatori"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dettaglio"
            Key             =   "DETTAGLIO"
            Object.ToolTipText     =   "Modifica dettaglio trasportatore"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rifiuti Trasportabili"
            Key             =   "TRASPORTABILI"
            Object.ToolTipText     =   "Elenco delle operazioni di Smaltimento/Recupero"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro"
            Key             =   "FILTRO"
            Object.ToolTipText     =   "Imposta un filtro per la ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAutomezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private sourceElenco As String
Private pidSoggetto As Long
Private pcodSoggetto As String


Const TXT_TARGA = 0
Const TXT_TARGARIMORCHIO = 1
Const TXT_DESCRIZIONE = 2

Const LIST_RIFIUTI_DA_ASS = 0
Const LIST_RIFIUTI_ASS = 1

Const CMD_LST_ASSOCIA_ALL = 0
Const CMD_LST_ASSOCIA_ONE = 1
Const CMD_LST_DISASSOCIA_ALL = 2
Const CMD_LST_DISASSOCIA_ONE = 3

Private colPicture As New Collection
Private pErrore As Boolean
Private pNoEvent As Boolean

Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    
    On Error GoTo cmd_ClickERR
    Select Case Index
        Case 0
            cancelCurrentOperation
            Unload Me
        Case 2
            ricaricaLista
        Case 3
            pErrore = DevEnvMDB.rscmdAutomezzi.EditMode <> adEditNone
            If (validaDati) Then
                If DevEnvMDB.rscmdAutomezzi.EditMode <> adEditNone Then
                    DataGrid1.Refresh
                    DevEnvMDB.rscmdAutomezzi.UpdateBatch adAffectAllChapters
                End If
                
                If Not (TabStrip1.SelectedItem Is Nothing) Then
                    If (TabStrip1.SelectedItem.key = "TRASPORTABILI") Then
                        salvaAssociazioni
                    End If
                    
                    If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
                    
                    pErrore = False
                End If
            End If
            
            
        Case 4
            TabStrip1.Tabs("DETTAGLIO").Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdAutomezzi.AddNew
            
            If pidSoggetto > 0 Then
                txtidSoggetto.Text = pidSoggetto
                txtCodSoggetto.Text = pcodSoggetto
                
                disabilitaAssociazione False
                
                
            End If
            
        Case 5
            If (MsgBox("Attenzione, se sicuro di voler cancellare l'impianto: " & DevEnvMDB.rscmdAutomezzi("targa"), vbQuestion + vbYesNo) = vbYes) Then
                DevEnvMDB.rscmdAutomezzi.Delete
                DevEnvMDB.rscmdAutomezzi.UpdateBatch adAffectAllChapters
            End If
    End Select
    
    Exit Sub
cmd_ClickERR:
    MsgBox Err.Description
End Sub


Private Function formatValue(Item As Column) As Variant
    On Error Resume Next
    Dim v As Variant
    
    v = Item.Value
    If Err.Number > 0 Then v = Item.Text
    
    If (Item.NumberFormat <> "") Then
        formatValue = toZero(v)
    Else
        formatValue = "=" & Chr$(34) & nullToBlank(v) & Chr$(34)
    End If
End Function


Private Sub ricaricaLista()
    Dim search As String
    
    Screen.MousePointer = vbHourglass
    
    search = vbNullString
    
    If txt(TXT_TARGA) <> vbNullString Then
        search = " Targa like '" & txt(TXT_TARGA) & "%'"
    End If
    

    If txt(TXT_TARGARIMORCHIO) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " targarimorchio like '" & txt(TXT_TARGARIMORCHIO) & "%'"
    End If
    
    If txt(TXT_DESCRIZIONE) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " Descrizione like '" & txt(TXT_DESCRIZIONE) & "%'"
    End If
    
    
    'If pidSoggetto > 0 Then
    '    If search <> vbNullString Then search = search & " AND "
    '    search = search & " codUtente=" & pidSoggetto & " and codSoggetto='" & pcodSoggetto & "'"
   '
    '
    'End If
    
    DevEnvMDB.rscmdAutomezzi.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdAutomezzi.Filter = search & " Descrizione <> '123'"
    End If
    
    TabStrip1.Tabs("LISTA").Selected = True

    Screen.MousePointer = vbDefault
End Sub


Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs("DETTAGLIO").Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'caricamento
    
    If LastRow > -1 And LastCol > -1 Then
        cmbGrid.Visible = False
        cmbGrid.Move DataGrid1.left + DataGrid1.Columns(DataGrid1.col).left + 7, _
                    DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row) - 15, _
                    DataGrid1.Columns(DataGrid1.col).width
        
        
        Select Case UCase$(DataGrid1.Columns(DataGrid1.col).DataField)
            Case UCase$("FlgDefault")
                caricaComboSINO cmbGrid, DataGrid1.col
        
        End Select
    
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo LoadERR
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    
    cmbGrid.font = DataGrid1.font
    cmbGrid.font.size = DataGrid1.font.size
    cmbGrid.Visible = False

    Exit Sub
LoadERR:
    MsgBox "Load frmAutomezzi:" & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    resizeControl

    DoEvents
End Sub



Private Sub resizeControl()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    TabStrip1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmd(0).Height - 2 * offsetX
    
    
    For i = Picture1.LBound To Picture1.ubound
        Picture1(i).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    DevEnvMDB.rscmdAutomezzi.Filter = vbNullString
    
    Set colPicture = Nothing
End Sub

Private Sub Picture1_Resize(Index As Integer)
    Dim widthList As Long

    Select Case Index
        Case 0
            DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
        Case 1
        Case 2
            widthList = (Picture1(2).ScaleWidth - Picture2(0).ScaleWidth - 3 * offsetX) / 2
            
            'smaltimento
            Label2(0).Move offsetX, offsetY
            List1(LIST_RIFIUTI_DA_ASS).Move offsetX, Label2(0).top + Label2(0).Height, widthList, Picture1(2).ScaleHeight - List1(LIST_RIFIUTI_DA_ASS).top
            Picture2(0).Move offsetX + widthList, (List1(0).Height - Picture2(0).ScaleHeight) / 2
            Label2(1).Move widthList + Picture2(0).width + offsetX, offsetY
            List1(LIST_RIFIUTI_ASS).Move Label2(1).left, List1(LIST_RIFIUTI_DA_ASS).top, widthList, Picture1(2).ScaleHeight - List1(LIST_RIFIUTI_DA_ASS).top
    End Select
End Sub

Private Sub cancelCurrentOperation()
    On Error GoTo cancelCurrentOperationERR
 
    txtidSoggetto.Text = ""
    txtCodSoggetto.Text = ""

    If (DevEnvMDB.rscmdAutomezzi.EOF Or DevEnvMDB.rscmdAutomezzi.BOF) Then Exit Sub
    
    If DevEnvMDB.rscmdAutomezzi.EditMode = adEditAdd Then
        DevEnvMDB.rscmdAutomezzi.CancelBatch adAffectAllChapters
        If (Not DevEnvMDB.rscmdAutomezzi.EOF And Not DevEnvMDB.rscmdAutomezzi.BOF) Then
            DevEnvMDB.rscmdAutomezzi.MoveFirst
        End If
    ElseIf DevEnvMDB.rscmdAutomezzi.EditMode = adEditInProgress Then
        DevEnvMDB.rscmdAutomezzi.CancelUpdate
    End If
cancelCurrentOperationERR:
    
End Sub



Private Sub TabStrip1_Click()
    
   
    If pNoEvent Then Exit Sub
    

    If (Not DevEnvMDB.rscmdAutomezzi.EOF And Not DevEnvMDB.rscmdAutomezzi.BOF) Then
        If DevEnvMDB.rscmdAutomezzi.EditMode <> adEditNone Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click 3
                
                If pErrore Then
                    pNoEvent = True
                    TabStrip1.Tabs("DETTAGLIO").Selected = True
                    pNoEvent = False
                End If
            Else
                cancelCurrentOperation
            End If
        End If
    End If
    
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    cmd(2).enabled = TabStrip1.SelectedItem.key = "FILTRO"
    cmd(2).default = TabStrip1.SelectedItem.key = "FILTRO"
    cmd(3).enabled = TabStrip1.SelectedItem.key = "DETTAGLIO" Or TabStrip1.SelectedItem.key = "TRASPORTABILI"
    cmd(3).default = TabStrip1.SelectedItem.key = "DETTAGLIO" Or TabStrip1.SelectedItem.key = "TRASPORTABILI"
    cmd(0).default = TabStrip1.SelectedItem.key = "LISTA"
    
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    colPicture(TabStrip1.SelectedItem.key).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1



    If (TabStrip1.SelectedItem.key = "DETTAGLIO") Then
        txtDescrizione.SetFocus
         cmbGrid.Visible = False
    End If

    If (TabStrip1.SelectedItem.key = "TRASPORTABILI") Then
        cmbGrid.Visible = False
        disabilitaAssociazione txtTarga <> ""
        List1(LIST_RIFIUTI_DA_ASS).SetFocus
        
        If List1(LIST_RIFIUTI_DA_ASS).Tag <> txtTarga.Text Then
            caricaListe
            List1(LIST_RIFIUTI_DA_ASS).Tag = txtTarga.Text
        End If
        
    End If


    If (TabStrip1.SelectedItem.key = "FILTRO") Then
        txt(0).SetFocus
        cmbGrid.Visible = False
    End If



End Sub

Public Sub showMe(Optional idSoggetto As Long = 0, Optional Des As String = "", Optional codSoggetto As String = "")
    
    pidSoggetto = idSoggetto
    pcodSoggetto = codSoggetto
    
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
        
    
    DevEnvMDB.cmdAutomezzi pidSoggetto, pcodSoggetto
    
    colPicture.Add Picture1(0), TabStrip1.Tabs(1).key
    colPicture.Add Picture1(1), TabStrip1.Tabs(2).key
    colPicture.Add Picture1(2), TabStrip1.Tabs(3).key
    colPicture.Add Picture1(3), TabStrip1.Tabs(4).key
    
    If idSoggetto > 0 Then Me.Caption = "Elenco Automezzi di " & Des
    
    If (codSoggetto = "U") Then
        TabStrip1.Tabs.Remove ("TRASPORTABILI")
    End If
    
    TabStrip1.Tabs("LISTA").Selected = True

    Screen.MousePointer = vbDefault
    
    sourceElenco = DevEnvMDB.rscmdAutomezzi.Source
    


    
    Me.Show vbModal
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (5)
End Sub


Private Sub caricaListe()
    Dim i As Integer
    
    For i = List1.LBound To List1.ubound
        List1(i).Clear
    Next
    
    If DevEnvMDB.rscmdCodCerAutomezziDaAssociare.state = adStateOpen Then
        DevEnvMDB.rscmdCodCerAutomezziDaAssociare.Close
    End If
    
    
        
    DevEnvMDB.cmdCodCerAutomezziDaAssociare txtTarga, pidSoggetto, pcodSoggetto
        
    With DevEnvMDB.rscmdCodCerAutomezziDaAssociare
        Do While Not .EOF
            If IsNull(.Fields("targa")) Then
                List1(LIST_RIFIUTI_DA_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
            Else
                List1(LIST_RIFIUTI_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
            End If
            .MoveNext
        Loop
    End With

End Sub

Private Sub disabilitaAssociazione(abilita As Boolean)
    Dim i As Integer
    
    For i = Command1.LBound To Command1.ubound
        Command1(i).enabled = abilita
    Next

End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case CMD_LST_ASSOCIA_ALL, CMD_LST_ASSOCIA_ONE
            spostaElemento List1(LIST_RIFIUTI_DA_ASS), List1(LIST_RIFIUTI_ASS), Index = CMD_LST_ASSOCIA_ALL
        Case CMD_LST_DISASSOCIA_ONE, CMD_LST_DISASSOCIA_ALL
            spostaElemento List1(LIST_RIFIUTI_ASS), List1(LIST_RIFIUTI_DA_ASS), Index = CMD_LST_DISASSOCIA_ALL
    End Select
End Sub


Private Sub spostaElemento(DaList As ListBox, AList As ListBox, all As Boolean)
    Dim i As Integer
    
    If all = False Then
        
        If DaList.ListIndex = -1 Then Exit Sub
        
        AList.AddItem DaList.List(DaList.ListIndex)
        DaList.RemoveItem DaList.ListIndex
    Else
        For i = 0 To DaList.ListCount - 1
            DaList.ListIndex = 0
            spostaElemento DaList, AList, False
        Next
    End If
    
End Sub

Private Sub List1_DblClick(Index As Integer)
    Select Case Index
        Case LIST_RIFIUTI_DA_ASS
            Command1_Click (CMD_LST_ASSOCIA_ONE)
        Case LIST_RIFIUTI_ASS
            Command1_Click (CMD_LST_DISASSOCIA_ONE)
    
    End Select
End Sub

Private Sub salvaAssociazioni()
    Dim i As Integer
    Dim qry As ADODB.Command
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codCer As String
    
    gDBC.Execute "delete * from AssAutomezzoCodiciCer where targa='" & txtTarga & "' and  idSoggetto=" & pidSoggetto & " and  codSoggetto='" & pcodSoggetto & "'"
    
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "insert into AssAutomezzoCodiciCer (codCategoria,codSottoCategoria,codCer " _
        & ",targa,idSoggetto,codSoggetto) values (?,?,?,?,?,?)"
    qry.CommandType = adCmdText
    qry.Prepared = True
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 15)
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 1)

    
    Set qry.ActiveConnection = gDBC
    
    For i = 0 To List1(LIST_RIFIUTI_ASS).ListCount - 1
        
        codCer = left$(List1(LIST_RIFIUTI_ASS).List(i), 6)
        codCategoria = left$(codCer, 2)
        codSottoCategoria = Mid$(codCer, 3, 2)
        codCer = Right$(codCer, 2)
        
        qry.Parameters(0).Value = codCategoria
        qry.Parameters(1).Value = codSottoCategoria
        qry.Parameters(2).Value = codCer
        qry.Parameters(3).Value = txtTarga
        qry.Parameters(4).Value = pidSoggetto
        qry.Parameters(5).Value = pcodSoggetto
        qry.Execute
    Next
    
    
End Sub




Private Sub cmbGrid_Change()
    If cmbGrid.Tag <> "" And cmbGrid.Locked = False Then DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
End Sub

Private Sub cmbGrid_Click()
    cmbGrid_Change
End Sub



Private Sub caricaComboSINO(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    cmb.AddItem "Si"
    cmb.AddItem "No"
    On Error Resume Next
    cmb.Text = IIf(DataGrid1.Columns(DataGrid1.col).Value = -1, "Si", "No")
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub




Private Sub txtTaraTotale_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub

Private Function validaDati() As Boolean
    Dim campoInErrore As Boolean
    
    validaDati = False
    campoInErrore = False
    
    txtDescrizione.ToolTipText = ""
    txtTarga.ToolTipText = ""

    
    txtDescrizione.BackColor = -2147483643
    txtTarga.BackColor = -2147483643
    
    
    If txtDescrizione.Text = "" Then
        txtDescrizione.ToolTipText = "Attenzione la descrizione è obbligatoria"
        txtDescrizione.BackColor = vbRed
        txtDescrizione.SetFocus
        
        campoInErrore = True
    End If
    
    If txtTarga.Text = "" Then
        txtTarga.ToolTipText = "Attenzione la Targa è obbligatoria"
        txtTarga.BackColor = vbRed
        txtTarga.SetFocus
        
        campoInErrore = True
    End If
    
      
    validaDati = Not campoInErrore

End Function

