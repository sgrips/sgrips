VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCoefficienteKAKB 
   Caption         =   "Coefficiente"
   ClientHeight    =   7740
   ClientLeft      =   2130
   ClientTop       =   1755
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   375
      Index           =   0
      Left            =   7140
      TabIndex        =   13
      Top             =   6900
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   180
      TabIndex        =   5
      Top             =   6420
      Width           =   6435
      Begin VB.CommandButton Cmd 
         Caption         =   "&IPS UA"
         Height          =   375
         Index           =   1
         Left            =   4140
         TabIndex        =   11
         Top             =   480
         Width           =   885
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   75
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1590
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Cmd 
         Caption         =   " &IPS UD"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   885
      End
      Begin MSComCtl2.DTPicker DTPAnno 
         Height          =   420
         Left            =   5100
         TabIndex        =   12
         Top             =   420
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   741
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
         CustomFormat    =   "yyyy"
         Format          =   19726339
         UpDown          =   -1  'True
         CurrentDate     =   36967
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IPS UD"
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
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IQS UD"
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
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   675
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteKAKB.frx":0000
      Height          =   1035
      Index           =   3
      Left            =   660
      TabIndex        =   4
      Top             =   2760
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1826
      _Version        =   393216
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
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
      BeginProperty Column01 
         DataField       =   "CodUtenza"
         Caption         =   "CodUtenza"
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
         DataField       =   "AnnoRiferimento"
         Caption         =   "Anno"
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
         DataField       =   "Descrizione"
         Caption         =   "Numero Componenti"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   6510,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteKAKB.frx":0013
      Height          =   1275
      Index           =   2
      Left            =   660
      TabIndex        =   3
      Top             =   4020
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   2249
      _Version        =   393216
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ForeColor       =   -2147483642
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
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
      BeginProperty Column01 
         DataField       =   "CodAttivita"
         Caption         =   "CodAttivita"
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
         DataField       =   "AnnoRiferimento"
         Caption         =   "Anno"
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
         DataField       =   "Descrizione"
         Caption         =   "Attivita"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "IQS"
         Caption         =   "IQS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "IPS"
         Caption         =   "IPS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   6510,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   989,858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteKAKB.frx":0026
      Height          =   825
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   1455
      _Version        =   393216
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "IdCoefficiente"
         Caption         =   "IdCoefficiente"
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
      BeginProperty Column01 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
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
         DataField       =   "AnnoRiferimento"
         Caption         =   "Anno"
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
         DataField       =   "Descrizione"
         Caption         =   "Nr.Componenti "
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Minimo"
         Caption         =   "Minimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Massimo"
         Caption         =   "Massimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Medio"
         Caption         =   "Medio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteKAKB.frx":0039
      Height          =   1200
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   2117
      _Version        =   393216
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
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
      BeginProperty Column01 
         DataField       =   "IDCoefficiente"
         Caption         =   "IDCoefficiente"
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
         DataField       =   "AnnoRiferimento"
         Caption         =   "Anno"
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
         DataField       =   "Descrizione"
         Caption         =   "Attivita"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "MinK"
         Caption         =   "Minimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "MaxK"
         Caption         =   "Massimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   6809,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   9234
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado1 
      Height          =   330
      Left            =   2700
      Top             =   7890
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
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
      DataSourceName  =   ""
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
End
Attribute VB_Name = "FrmCoefficiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mCodComune As String
Private mParteFissa As Boolean
Private mArea  As String
Private mAbitanti As Double
Private sParteF As String
Private sParGeo As String
Private sAbitanti As String
Private mAnnoRiferimento As String
Private mTipoCoeff As tipo_coefficiente

Private Enum tipo_coefficiente
    tpo_coeff_KAKB = 0
    tpo_coeff_KCKD = 1
    tpo_coeff_507UA = 2
    tpo_coeff_507UD = 3
End Enum

Private MaxCoeff As Double
Private MinCoeff As Double

Private Const COEFF_IQSUD_ID = 2
Private Const COEFF_IPSUD_ID = 3

Public Sub Carica(Optional sCodComune As String = vbNullString)
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    'mDataIns = DateAdd("yyyy", -1, Now)
    mAnnoRiferimento = Year(Now)
    mCodComune = sCodComune
    DTPAnno.Value = Format$(Now, "dd/mm/yyyy")
    
    Me.Show vbModal
End Sub




Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 2, 1
            CalcolaIPS
    End Select
End Sub


Private Sub DataGrid1_BeforeUpdate(Index As Integer, Cancel As Integer)
    'Valida il valore inserito
    With DataGrid1(mTipoCoeff)
    
        Select Case mTipoCoeff
        
        
            Case tpo_coeff_KAKB, tpo_coeff_KCKD
                If mParteFissa And mTipoCoeff = tpo_coeff_KAKB Then
                    If CDbl(.Columns("Valore")) > CDbl(.Columns("Massimo")) Then
                        If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                            Cancel = False
                        Else
                            Cancel = True
                        End If
                    End If
                ElseIf CDbl(.Columns("Valore")) > CDbl(.Columns("Massimo")) Or CDbl(.Columns("Valore")) < CDbl(.Columns("Minimo")) Then
                        If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                             Cancel = False
                        Else
                            Cancel = True
                        End If
                End If
            Case tpo_coeff_507UA
                If (CDbl(.Columns("IQS").Text) > MaxCoeff Or CDbl(.Columns("IQS").Text) < MinCoeff) And .Columns("IQS").Text <> 0 Then
                    If MsgBox("Valore IQS non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                        Cancel = False
                    Else
                        Cancel = True
                    End If
                End If
        End Select
        
        
        
        
    End With

End Sub

Private Sub DataGrid1_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
        If DataError = 7011 Then Response = 0
End Sub

Private Sub DataGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{ENTER}", True
    End If
End Sub



Private Sub DataGrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.Value)
    If mTipoCoeff = tpo_coeff_507UA Or mTipoCoeff = tpo_coeff_507UD Then
        CaricaDati507
    Else
        CaricaDati
    End If
End Sub

Private Sub Form_Activate()
    'Mi posiziono sulla griglia
    
    DataGrid1(mTipoCoeff).SetFocus
    
End Sub

Private Sub Form_Load()
    
    
    TabStrip.Tabs.Clear
    TabStrip.Tabs.Add , "KA", "PF Utenza Domestica KA"
    TabStrip.Tabs.Add , "KB", "PV Utenza Domestica KB"
    TabStrip.Tabs.Add , "KC", "PF Utenza Attività KC"
    TabStrip.Tabs.Add , "KD", "PV Utenza Attività KD"
    TabStrip.Tabs.Add , "507UA", "507 Utenza Attività"
    TabStrip.Tabs.Add , "507UD", "507 Utenza Domestica"
    
    MDIMain.mnuComprimi.Enabled = False

   
    ado1.ConnectionString = gDBC
    ado1.CommandType = adCmdText
    ado1.RecordSource = vbNullString
    
    GetParamValue mAnnoRiferimento, mCodComune, "COEFFIQSUD", MaxCoeff, MinCoeff

    TabStrip.Tabs(1).Selected = True
End Sub


Private Function CaricaDati() As Boolean
    Dim Font As New StdFont
    Static SistemaColonne As Boolean
    Dim ParAbitanti As String
    Dim rs As Recordset
    Dim qry As ADODB.Command
    

    CaricaDati = False
    
    On Error Resume Next
    ado1.Recordset.Close
    On Error GoTo 0
    
    'reperisco i dati del comune
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, ParAbitanti)
    
    If rs.EOF Then
        'MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
        mArea = "No Dati"
        mAbitanti = 0
        sParGeo = "No Dati"
    Else
        'Area di appartenenza del comune
        mArea = rs("Descrizione")
        mAbitanti = rs("NumAbitantiResidenti")
        sParGeo = rs("AreaGeografica")
        sAbitanti = ParAbitanti
    
        
        'controllo se esistono già i coefficienti
        Set rs = Nothing
        If mTipoCoeff = tpo_coeff_KAKB Then
            sParGeo = IIf(mParteFissa, sParGeo, "0")
            sAbitanti = IIf(mParteFissa, sAbitanti, "0")
            Set qry = gADOX.Procedures("InsNrComponeniFamigliaInCoeffUtenzeDomesticheComunali").Command
        Else
            Set qry = gADOX.Procedures("InsCategorieAttivitaInCoeffCategorieAttivitaComunali").Command
        End If
        qry.Parameters.Refresh
        qry.Parameters("Comune") = mCodComune
        qry.Parameters("CodParGeo") = sParGeo
        qry.Parameters("CodR") = sAbitanti
        qry.Parameters("ParteF") = mParteFissa
        qry.Parameters("[Anno]") = mAnnoRiferimento
        qry.Execute
        
        If mTipoCoeff = tpo_coeff_KAKB Then
            Set qry = gADOX.Procedures("InserisciUtenzeDomesticheComunali").Command
        
        Else
            Set qry = gADOX.Procedures("InserisciCategorieAttivitaComunali").Command
        End If
        qry.Parameters.Refresh
        qry.Parameters(0).Value = mCodComune
        qry.Parameters(1) = mAnnoRiferimento
        qry.Execute
        
    End If
    

    If mTipoCoeff = tpo_coeff_KAKB Then
        ado1.RecordSource = "SELECT ValoriCoefficientiUDComunali.AnnoRiferimento,ValoriCoefficientiUDComunali.CodComune,ValoriCoefficientiUDComunali.IdCoefficiente , NumeroComponentiFamiglia.NumeroComponenti,  CoefficientiUtenzeDomestiche.Minimo, CoefficientiUtenzeDomestiche.Massimo,CoefficientiUtenzeDomestiche.Medio, ValoriCoefficientiUDComunali.Valore, NumeroComponentiFamiglia.Descrizione " & _
                             "FROM (CoefficientiUtenzeDomestiche INNER JOIN ValoriCoefficientiUDComunali ON CoefficientiUtenzeDomestiche.IdCoefficiente = ValoriCoefficientiUDComunali.IdCoefficiente) INNER JOIN NumeroComponentiFamiglia ON CoefficientiUtenzeDomestiche.IdNumeroComponenti = NumeroComponentiFamiglia.NumeroComponenti " & _
                             "WHERE ValoriCoefficientiUDComunali.CodComune='" & mCodComune & "' and CoefficientiUtenzeDomestiche.CodParGeografica='" & sParGeo & "' AND CoefficientiUtenzeDomestiche.CodRange='" & sAbitanti & "' AND CoefficientiUtenzeDomestiche.ParteFissa=" & IIf(mParteFissa, "True", "False") & " and  ValoriCoefficientiUDComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by NumeroComponentiFamiglia.NumeroComponenti"
        
        
        With DataGrid1(mTipoCoeff)
            If mParteFissa And Not SistemaColonne Then
                Font.Name = "MS Sans Serif"
                Font.size = 12
                .HeadFont = Font
                
                Font.Name = "MS Sans Serif"
                Font.size = 11
                .Font = Font
                
                .Columns("Minimo").Caption = mArea
                .Columns(mArea).Width = 2000
                .Columns("Valore").Width = 2000
                .Columns("Massimo").Visible = False
                .Columns("Medio").Visible = False
                SistemaColonne = True
            ElseIf Not SistemaColonne Then
                Font.Name = "MS Sans Serif"
                Font.size = 12
                .HeadFont = Font
                
                Font.Name = "MS Sans Serif"
                Font.size = 10
                .Font = Font
                
                .Columns("Minimo").Width = 1305.071
                .Columns("Valore").Width = 975.1182
                
                .Columns("Minimo").Caption = "Minimo"
                .Columns("Massimo").Visible = True
                .Columns("Medio").Visible = True
                SistemaColonne = True
            End If
        End With
    Else
        ado1.RecordSource = "SELECT ValoriCoefficientiCAComunali.AnnoRiferimento,ValoriCoefficientiCAComunali.IdCoefficiente,ValoriCoefficientiCAComunali.CodComune,CoefficientiCategorieAttivita.MinK, CoefficientiCategorieAttivita.MaxK, ValoriCoefficientiCAComunali.Valore, CategorieAttivitaComunali.Descrizione " & _
                             "FROM (CoefficientiCategorieAttivita INNER JOIN ValoriCoefficientiCAComunali ON CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente) " _
                             & "INNER JOIN CategorieAttivitaComunali ON CoefficientiCategorieAttivita.CodTipologiaAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and ValoriCoefficientiCAComunali.CodComune = CategorieAttivitaComunali.CodComune and ValoriCoefficientiCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " _
                             & "WHERE ValoriCoefficientiCAComunali.CodComune='" & mCodComune & "' and CoefficientiCategorieAttivita.CodParGeografica='" & sParGeo & "' AND CoefficientiCategorieAttivita.CodRange='" & sAbitanti & "' AND CoefficientiCategorieAttivita.ParteFissa=" & IIf(mParteFissa, "True", "False") & " and " _
                             & "ValoriCoefficientiCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by CoefficientiCategorieAttivita.CodTipologiaAttivita"
    End If
    ado1.Refresh

    If DataGrid1(mTipoCoeff).HeadLines > 1 Then DataGrid1(mTipoCoeff).Row = DataGrid1(mTipoCoeff).HeadLines - 1
    
    CaricaDati = Not ado1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1(mTipoCoeff).RowHeight = 30
    Else
        DataGrid1(mTipoCoeff).RowHeight = 270
    End If
    
End Function

Private Function CaricaDati507() As Boolean
    Dim sAbitanti As String
    Dim rs As Recordset
    Dim qry As New ADODB.Command
    
    CaricaDati507 = False
    
    
    'Reperisco i dati comunali e anagrafici
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, sAbitanti)
    If rs.EOF Then
        MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
    Else
        'Inserisco le attività
        Set rs = Nothing
        If mTipoCoeff = tpo_coeff_507UA Then
            Set qry = gADOX.Procedures("InsCategorieAttivitaInCoeffIQSCAComunali").Command
        Else
            Set qry = gADOX.Procedures("InsUtenzeDomesticheInCoeff507UDComunali").Command
        End If
        qry.Parameters.Refresh
        qry.Parameters("pCodComune") = mCodComune
        qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
        qry.Execute
        Set qry = Nothing
        
    End If
    
    On Error Resume Next
    ado1.Recordset.Close
    On Error GoTo 0
    If mTipoCoeff = tpo_coeff_507UA Then
    
        ado1.RecordSource = "SELECT ValoriCoefficientiIQSCAComunali.AnnoRiferimento,ValoriCoefficientiIQSCAComunali.CodAttivita,ValoriCoefficientiIQSCAComunali.CodComune, ValoriCoefficientiIQSCAComunali.IQS,ValoriCoefficientiIQSCAComunali.IPS, CategorieAttivitaComunali.Descrizione " & _
                             "FROM  CategorieAttivitaComunali INNER JOIN ValoriCoefficientiIQSCAComunali ON ValoriCoefficientiIQSCAComunali.CodAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and " & _
                             " ValoriCoefficientiIQSCAComunali.CodComune = CategorieAttivitaComunali.CodComune and " & _
                             " ValoriCoefficientiIQSCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " & _
                             "WHERE ValoriCoefficientiIQSCAComunali.CodComune='" & mCodComune & "'and ValoriCoefficientiIQSCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by ValoriCoefficientiIQSCAComunali.CodAttivita"
    Else
        ado1.RecordSource = "SELECT ValoriCoefficienti507UDComunali.CodComune,NumeroComponentiFamiglia.Descrizione, ValoriCoefficienti507UDComunali.AnnoRiferimento, ValoriCoefficienti507UDComunali.CodUtenza, ValoriCoefficienti507UDComunali.Valore " & _
                            "FROM NumeroComponentiFamiglia INNER JOIN (UtenzeDomesticheComunali INNER JOIN ValoriCoefficienti507UDComunali ON (UtenzeDomesticheComunali.IdNumeriComponenti = ValoriCoefficienti507UDComunali.CodUtenza) AND (UtenzeDomesticheComunali.AnnoRiferimento = ValoriCoefficienti507UDComunali.AnnoRiferimento) AND (UtenzeDomesticheComunali.CodComune = ValoriCoefficienti507UDComunali.CodComune)) ON NumeroComponentiFamiglia.NumeroComponenti = UtenzeDomesticheComunali.IdNumeriComponenti " & _
                            "WHERE (((ValoriCoefficienti507UDComunali.CodComune)='" & mCodComune & "') AND ((ValoriCoefficienti507UDComunali.AnnoRiferimento)='" & mAnnoRiferimento & "')) order by ValoriCoefficienti507UDComunali.CodUtenza"
    End If
    ado1.Refresh

    CaricaDati507 = Not ado1.Recordset.EOF
    
    If Not CaricaDati507 Then
        DataGrid1(mTipoCoeff).RowHeight = 30
    Else
        DataGrid1(mTipoCoeff).RowHeight = 270
        DataGrid1(mTipoCoeff).Refresh
    End If
    
   txt(COEFF_IQSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), "0.00")
   txt(COEFF_IPSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), "0.00")

End Function


Private Sub Form_Resize()
    Dim i As Integer
    
    TabStrip.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
    
    For i = DataGrid1.LBound To DataGrid1.UBound
        DataGrid1(i).Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    Next
    Cmd(0).Move Me.ScaleWidth - Cmd(0).Width - 100, Me.ScaleHeight - Cmd(0).Height - 100
    Frame1.Move 0, Me.ScaleHeight - Frame1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Calcola Quantità Rifiuti su UA
    QuantitaRifiutiDaUtenzeAttivita mCodComune, mAnnoRiferimento
    

    MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub TabStrip_Click()
    Dim i As Integer
    
    
    
    For i = 0 To DataGrid1.Count - 1
        DataGrid1(i).Visible = False
    Next
    mParteFissa = False
    Select Case TabStrip.SelectedItem.Key
        Case "KA"
            mParteFissa = True
            mTipoCoeff = tpo_coeff_KAKB
        Case "KB"
            mTipoCoeff = tpo_coeff_KAKB
        Case "KC"
            mParteFissa = True
            mTipoCoeff = tpo_coeff_KCKD
        Case "KD"
            mTipoCoeff = tpo_coeff_KCKD
        Case "507UA"
            mTipoCoeff = tpo_coeff_507UA
        Case "507UD"
            mTipoCoeff = tpo_coeff_507UD
    End Select
    
    DataGrid1(mTipoCoeff).Visible = True
    
    If mTipoCoeff = tpo_coeff_507UA Or mTipoCoeff = tpo_coeff_507UD Then
        CaricaDati507
    Else
        CaricaDati
    End If

    
    If mTipoCoeff = tpo_coeff_KAKB Then
        TabStrip.SelectedItem.Caption = IIf(mParteFissa, "PF Utenza Attività KA", "PV Utenza Domestica KB") & " - " & mArea
    ElseIf mTipoCoeff = tpo_coeff_KCKD Then
        TabStrip.SelectedItem.Caption = IIf(mParteFissa, "PF Utenza Attività KC", "PV Utenza Attività KD")
    ElseIf mTipoCoeff = tpo_coeff_507UA Then
        TabStrip.SelectedItem.Caption = "507 Utenza Attività"
    ElseIf mTipoCoeff = tpo_coeff_507UD Then
        TabStrip.SelectedItem.Caption = "507 Utenza Domestica"
    End If
    
End Sub
   
Sub CalcolaIPS()
    Dim rsKD As ADODB.Recordset
    Dim KM As Double
    Dim rsDatiTecniciComune As ADODB.Recordset
    Dim rsANA As ADODB.Recordset
    Dim mParAbitanti As String
    Dim i  As Integer
    Dim iOld As Integer
    Dim IPSUD As Double

    
    
    Set rsANA = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
    Set rsDatiTecniciComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    
    If rsDatiTecniciComune.EOF Then
        MsgBox "Attenzione dati tecnici mancanti", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    ElseIf rsANA.EOF Then
        MsgBox "Attenzione dati anagrafici", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    End If
    'KM
    KM = rsDatiTecniciComune("QRifiutiTotali") / rsDatiTecniciComune("TotSuperficieIscrRuolo")
    
    
    If mTipoCoeff = tpo_coeff_507UD Then
        
        IPSUD = (rsDatiTecniciComune("QRifiutiUtenzaDomestica") / rsDatiTecniciComune("TotSuperficieUtenzaDomestica")) / KM
        txt(COEFF_IPSUD_ID).Text = Format$(IPSUD, "0.00")
    Else
        'KD
        Set rsKD = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    
        i = 0
        iOld = DataGrid1(mTipoCoeff).Row
        DataGrid1(mTipoCoeff).Row = i
        Do While Not rsKD.EOF
            If DataGrid1(mTipoCoeff).Columns("CodAttivita").Text = rsKD("CodTipologiaAttivita") Then
                DataGrid1(mTipoCoeff).Columns("IPS") = rsKD("Valore") / KM
                i = i + 1
                If i < DataGrid1(mTipoCoeff).VisibleRows Then DataGrid1(mTipoCoeff).Row = i
            End If
            rsKD.MoveNext
        Loop
        DataGrid1(mTipoCoeff).Row = iOld
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index = COEFF_IQSUD_ID And ((CDbl(txt(COEFF_IQSUD_ID)) > MaxCoeff Or CDbl(txt(COEFF_IQSUD_ID)) < MinCoeff) And CDbl(txt(COEFF_IQSUD_ID)) <> 0) Then
        If MsgBox("Valore IQS non ammesso nel range " & MinCoeff & " - " & MaxCoeff & vbCrLf & " Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
    
End Sub
