VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmCoefficiente 
   Caption         =   "Coefficienti"
   ClientHeight    =   10440
   ClientLeft      =   2130
   ClientTop       =   1755
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficiente.frx":0000
      Height          =   1275
      Index           =   2
      Left            =   660
      TabIndex        =   3
      Top             =   4020
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   2249
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
      Bindings        =   "FrmCoefficiente.frx":0013
      Height          =   1035
      Index           =   3
      Left            =   660
      TabIndex        =   4
      Top             =   2760
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1826
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
            Format          =   "0,00"
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
      Bindings        =   "FrmCoefficiente.frx":0026
      Height          =   825
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   1455
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
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
      Bindings        =   "FrmCoefficiente.frx":0039
      Height          =   1200
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   2117
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
            Format          =   "0,00"
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
         Left            =   5040
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
         Format          =   58392579
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
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   9234
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
Private pTipoCalcolo  As tpoTariffa
Private lastColumn As Long

Public Enum tipo_coefficiente
    tpo_coeff_KAKB = 0
    tpo_coeff_KCKD = 1
    tpo_coeff_507UA = 2
    tpo_coeff_507UD = 3
End Enum

Private MaxCoeff As Double
Private MinCoeff As Double

Private Const COEFF_IQSUD_ID = 2
Private Const COEFF_IPSUD_ID = 3

Private gDisableError As Boolean


Public Sub Carica(Optional sCodComune As String = vbNullString, Optional AnnoRiferimento As String = vbNullString, Optional tipoCalcolo As tpoTariffa = TariffaNonImpostata, Optional modal As Boolean = True)
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    pTipoCalcolo = tipoCalcolo
    'mDataIns = DateAdd("yyyy", -1, Now)
    mAnnoRiferimento = IIf(AnnoRiferimento <> vbNullString, AnnoRiferimento, Year(Now))
    mCodComune = sCodComune
    If AnnoRiferimento = vbNullString Then
        DTPAnno.value = Format$(Now, "dd/mm/yyyy")
    Else
        DTPAnno.Enabled = False
        
    End If
    
    Cmd(0).Enabled = DTPAnno.Enabled
    Me.Caption = "Coefficienti:" & getDesComune(sCodComune)


    tabStrip.Tabs.Clear
    
    If pTipoCalcolo = TariffaNonImpostata Or _
       pTipoCalcolo = TariffaNormalizzata Or _
       pTipoCalcolo = TariffaCostiIndivisibili Then
        tabStrip.Tabs.Add , "KA", "651 PF Utenza Domestica KA"
        tabStrip.Tabs.Add , "KB", "651 PV Utenza Domestica KB"
        tabStrip.Tabs.Add , "KC", "651 PF Utenza Attività KC"
        tabStrip.Tabs.Add , "KD", "651 PV Utenza Attività KD"
    End If
    
    If pTipoCalcolo = TariffaNonImpostata Or pTipoCalcolo = Tariffa507 Then
        tabStrip.Tabs.Add , "507UA", "652 Utenza Attività"
        tabStrip.Tabs.Add , "507UD", "652 Utenza Domestica"
    End If

    tabStrip.Tabs(1).Selected = True
    
    If modal Then
        Me.Show vbModal
    Else
         Me.Show vbModeless
    End If
End Sub




Private Sub Cmd_Click(index As Integer)
    Select Case index
        Case 0
        
            LetParamValue mAnnoRiferimento, mCodComune, "COEFFIQSUD", True, Txt(COEFF_IQSUD_ID).Text
            LetParamValue mAnnoRiferimento, mCodComune, "COEFFIPSUD", True, Txt(COEFF_IPSUD_ID).Text
            Unload Me
        Case 2, 1
            CalcolaIPS
    End Select
End Sub





Private Sub DataGrid1_AfterColUpdate(index As Integer, ByVal ColIndex As Integer)
    If (Round(DataGrid1(index).Columns(ColIndex).value, 2)) <> DataGrid1(index).Columns(ColIndex).value Then
       DataGrid1(index).Columns(ColIndex).value = Round(DataGrid1(index).Columns(ColIndex).value, 2)
    End If
End Sub

Private Sub DataGrid1_BeforeColUpdate(index As Integer, ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim Valore As Double
    
    'Valida il valore inserito
    If gDisableError Then Exit Sub

    On Error GoTo DataGrid1_BeforeColUpdateERR
    
    Valore = CDbl(DataGrid1(index).Tag)
    With DataGrid1(index)
    
        Select Case mTipoCoeff
        
         
            Case tpo_coeff_KAKB, tpo_coeff_KCKD
                If mParteFissa And mTipoCoeff = tpo_coeff_KAKB Then
                    If Valore > CDbl(.Columns("Massimo")) Then
                        
                        If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                            Cancel = False
                        Else
                            Cancel = True
                        End If
                    End If
                ElseIf Valore > CDbl(.Columns("Massimo")) Or Valore < CDbl(.Columns("Minimo")) Then
                        If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                             Cancel = False
                        Else
                            Cancel = True
                        End If
                End If
            Case tpo_coeff_507UA
                If .Columns(ColIndex).Caption = "IQS" And (Valore > MaxCoeff Or Valore < MinCoeff) And Valore <> 0 Then
                    If MsgBox("Valore IQS non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
                        Cancel = False
                    Else
                        Cancel = True
                    End If
                End If
        End Select
        Exit Sub
DataGrid1_BeforeColUpdateERR:
        Cancel = True
        
        
        
    End With
End Sub

Private Sub DataGrid1_Change(index As Integer)
      ' Change is the easiest event to tap into.
      ' It is fired every time the cell text changes.

      ' Because you can only save the text when the cell is being edited,
      ' you have to store the value in another variable if you want
      ' to use that value in other events.

      ' Place the contents of the cell into the TAG property as a buffer.
      ' You could also use a variable to store the information.

      DataGrid1(index).Tag = GetCellText(DataGrid1(index))
      
End Sub

Private Sub DataGrid1_Error(index As Integer, ByVal DataError As Integer, Response As Integer)
        If DataError = 7011 Or DataError = 6153 Then Response = 0
End Sub

Private Sub DataGrid1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        lastColumn = DataGrid1(index).col
        Sendkeys "{ENTER}", True
    End If
End Sub

Private Sub DataGrid1_RowColChange(index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If lastColumn > 0 And index = mTipoCoeff Then DataGrid1(index).col = lastColumn
End Sub

Private Sub DataGrid1_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.value)
    
    
    If mTipoCoeff = tpo_coeff_507UA Or mTipoCoeff = tpo_coeff_507UD Then
        CaricaDati507
    Else
        CaricaDati
    End If
    SistemaCodiceAreaSuTAB
    
End Sub

Private Sub Form_Activate()
    'Mi posiziono sulla griglia
    
    DataGrid1(mTipoCoeff).SetFocus
    
End Sub

Private Sub Form_Load()
    
    

    
    MDIMain.mnuComprimi.Enabled = False

   
    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = vbNullString
    
    GetParamValue mAnnoRiferimento, mCodComune, "COEFFIQSUD", MaxCoeff, MinCoeff

End Sub

Public Sub CreaCoefficenti(tipoCoeff As tipo_coefficiente, CodComune As String, AnnoRiferimento As String, ParteFissa As Boolean)
    Dim ParAbitanti As String
    Dim rs As Recordset
    Dim qry As ADODB.Command
  
     
    'reperisco i dati del comune
    Set rs = DatiAnagraficiComunali(CodComune, AnnoRiferimento, ParAbitanti)
    
    'cancello i dati di coefficenti DCAMAG o DCAMIN se il comune cambia il numero di abitanti
    Dim qryDelete As String
    

    
    If rs.EOF Then
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
        If tipoCoeff = tpo_coeff_KAKB Then
        
            qryDelete = "DELETE ValoriCoefficientiUDComunali.* " & _
               "FROM CoefficientiUtenzeDomestiche INNER JOIN ValoriCoefficientiUDComunali ON " & _
               "CoefficientiUtenzeDomestiche.IdCoefficiente = ValoriCoefficientiUDComunali.IdCoefficiente " & _
               "WHERE ValoriCoefficientiUDComunali.CodComune=" & Chr$(34) & mCodComune & Chr$(34) & " and " & _
               " ValoriCoefficientiUDComunali.AnnoRiferimento like " & Chr$(34) & mAnnoRiferimento & "%" & Chr$(34) & _
               " AND CoefficientiUtenzeDomestiche.CodRange=" & Chr$(34) & IIf(ParAbitanti = "DCAMAG", "DCAMIN", "DCAMAG") & Chr$(34)

            gDBC.Execute (qryDelete)
        
            sParGeo = IIf(ParteFissa, sParGeo, "0")
            sAbitanti = IIf(ParteFissa, sAbitanti, "0")
            Set qry = gADOX.Procedures("InsNrComponeniFamigliaInCoeffUtenzeDomesticheComunali").Command
        Else

            qryDelete = "DELETE ValoriCoefficientiCAComunali.* " & _
               "FROM CoefficientiCategorieAttivita INNER JOIN ValoriCoefficientiCAComunali ON " & _
               "CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente " & _
               "WHERE ValoriCoefficientiCAComunali.CodComune=" & Chr$(34) & mCodComune & Chr$(34) & " and " & _
               " ValoriCoefficientiCAComunali.AnnoRiferimento like " & Chr$(34) & mAnnoRiferimento & "%" & Chr$(34) & _
               " AND CoefficientiCategorieAttivita.CodRange=" & Chr$(34) & IIf(ParAbitanti = "DCAMAG", "DCAMIN", "DCAMAG") & Chr$(34)

            gDBC.Execute (qryDelete)
        
            Set qry = gADOX.Procedures("InsCategorieAttivitaInCoeffCategorieAttivitaComunali").Command
        End If
        qry.Parameters.Refresh
        qry.Parameters("Comune") = CodComune
        qry.Parameters("CodParGeo") = sParGeo
        qry.Parameters("CodR") = sAbitanti
        qry.Parameters("ParteF") = ParteFissa
        qry.Parameters("[Anno]") = AnnoRiferimento
        qry.Execute
        
        If tipoCoeff = tpo_coeff_KAKB Then
            Set qry = gADOX.Procedures("InserisciUtenzeDomesticheComunali").Command
        
        Else
            Set qry = gADOX.Procedures("InserisciCategorieAttivitaComunali").Command
        End If
        qry.Parameters.Refresh
        qry.Parameters(0).value = CodComune
        qry.Parameters(1) = AnnoRiferimento
        qry.Execute
        
    End If
End Sub



Private Function CaricaDati() As Boolean
    Dim Font As New StdFont
    CaricaDati = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    CreaCoefficenti mTipoCoeff, mCodComune, mAnnoRiferimento, mParteFissa
    

    If mTipoCoeff = tpo_coeff_KAKB Then
        ADO1.RecordSource = "SELECT ValoriCoefficientiUDComunali.AnnoRiferimento,ValoriCoefficientiUDComunali.CodComune,ValoriCoefficientiUDComunali.IdCoefficiente , NumeroComponentiFamiglia.NumeroComponenti,  CoefficientiUtenzeDomestiche.Minimo, CoefficientiUtenzeDomestiche.Massimo,CoefficientiUtenzeDomestiche.Medio, ValoriCoefficientiUDComunali.Valore, NumeroComponentiFamiglia.Descrizione " & _
                             "FROM (CoefficientiUtenzeDomestiche INNER JOIN ValoriCoefficientiUDComunali ON CoefficientiUtenzeDomestiche.IdCoefficiente = ValoriCoefficientiUDComunali.IdCoefficiente) INNER JOIN NumeroComponentiFamiglia ON CoefficientiUtenzeDomestiche.IdNumeroComponenti = NumeroComponentiFamiglia.NumeroComponenti " & _
                             "WHERE ValoriCoefficientiUDComunali.CodComune='" & mCodComune & "' and CoefficientiUtenzeDomestiche.CodParGeografica='" & sParGeo & "' AND CoefficientiUtenzeDomestiche.CodRange='" & sAbitanti & "' AND CoefficientiUtenzeDomestiche.ParteFissa=" & IIf(mParteFissa, "True", "False") & " and  ValoriCoefficientiUDComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by NumeroComponentiFamiglia.NumeroComponenti"
        
        
        With DataGrid1(mTipoCoeff)
            Font.Name = "MS Sans Serif"
            Font.size = 12
            .HeadFont = Font
            
            Font.Name = "MS Sans Serif"
            Font.size = 11
            .Font = Font
            
            
            On Error Resume Next
            Dim sCaption As String
            sCaption = ""
            sCaption = .Columns("Minimo").Caption
            sCaption = .Columns("No Dati").Caption
            sCaption = .Columns(mArea).Caption
            
            If mParteFissa Then
                If mArea <> sCaption And sCaption <> "" Then
                    .Columns(sCaption).Caption = mArea
                    .Columns(mArea).Width = 2000
                    .Columns("Valore").Width = 2000
                    .Columns("Massimo").Visible = False
                    .Columns("Medio").Visible = False
                End If
            Else
                If "Minimo" <> Caption And sCaption <> "" Then
                    .Columns("Minimo").Width = 1305.071
                    .Columns("Valore").Width = 975.1182
                    
                    .Columns(sCaption).Caption = "Minimo"
                    .Columns("Massimo").Visible = True
                    .Columns("Medio").Visible = True

                End If
            End If
            
        End With
    Else
        ADO1.RecordSource = "SELECT ValoriCoefficientiCAComunali.AnnoRiferimento,ValoriCoefficientiCAComunali.IdCoefficiente,ValoriCoefficientiCAComunali.CodComune,CoefficientiCategorieAttivita.MinK, CoefficientiCategorieAttivita.MaxK, ValoriCoefficientiCAComunali.Valore, CategorieAttivitaComunali.Descrizione " & _
                             "FROM (CoefficientiCategorieAttivita INNER JOIN ValoriCoefficientiCAComunali ON CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente) " _
                             & "INNER JOIN CategorieAttivitaComunali ON CoefficientiCategorieAttivita.CodTipologiaAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and ValoriCoefficientiCAComunali.CodComune = CategorieAttivitaComunali.CodComune and ValoriCoefficientiCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " _
                             & "WHERE ValoriCoefficientiCAComunali.CodComune='" & mCodComune & "' and CoefficientiCategorieAttivita.CodParGeografica='" & sParGeo & "' AND CoefficientiCategorieAttivita.CodRange='" & sAbitanti & "' AND CoefficientiCategorieAttivita.ParteFissa=" & IIf(mParteFissa, "True", "False") & " and " _
                             & "ValoriCoefficientiCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by CoefficientiCategorieAttivita.CodTipologiaAttivita"
    End If
    ADO1.Refresh

    If DataGrid1(mTipoCoeff).HeadLines > 1 Then DataGrid1(mTipoCoeff).row = DataGrid1(mTipoCoeff).HeadLines - 1
    
    CaricaDati = Not ADO1.Recordset.EOF
    
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
'        MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
    Else
        'Inserisco le attività
        gDBC.BeginTrans
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
        gDBC.CommitTrans
    End If
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    If mTipoCoeff = tpo_coeff_507UA Then
    
        ADO1.RecordSource = "SELECT ValoriCoefficientiIQSCAComunali.AnnoRiferimento,ValoriCoefficientiIQSCAComunali.CodAttivita,ValoriCoefficientiIQSCAComunali.CodComune, ValoriCoefficientiIQSCAComunali.IQS,ValoriCoefficientiIQSCAComunali.IPS, CategorieAttivitaComunali.Descrizione " & _
                             "FROM  CategorieAttivitaComunali INNER JOIN ValoriCoefficientiIQSCAComunali ON ValoriCoefficientiIQSCAComunali.CodAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and " & _
                             " ValoriCoefficientiIQSCAComunali.CodComune = CategorieAttivitaComunali.CodComune and " & _
                             " ValoriCoefficientiIQSCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " & _
                             "WHERE ValoriCoefficientiIQSCAComunali.CodComune='" & mCodComune & "'and ValoriCoefficientiIQSCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by ValoriCoefficientiIQSCAComunali.CodAttivita"
    Else
        ADO1.RecordSource = "SELECT ValoriCoefficienti507UDComunali.CodComune,NumeroComponentiFamiglia.Descrizione, ValoriCoefficienti507UDComunali.AnnoRiferimento, ValoriCoefficienti507UDComunali.CodUtenza, ValoriCoefficienti507UDComunali.Valore " & _
                            "FROM NumeroComponentiFamiglia INNER JOIN (UtenzeDomesticheComunali INNER JOIN ValoriCoefficienti507UDComunali ON (UtenzeDomesticheComunali.IdNumeriComponenti = ValoriCoefficienti507UDComunali.CodUtenza) AND (UtenzeDomesticheComunali.AnnoRiferimento = ValoriCoefficienti507UDComunali.AnnoRiferimento) AND (UtenzeDomesticheComunali.CodComune = ValoriCoefficienti507UDComunali.CodComune)) ON NumeroComponentiFamiglia.NumeroComponenti = UtenzeDomesticheComunali.IdNumeriComponenti " & _
                            "WHERE (((ValoriCoefficienti507UDComunali.CodComune)='" & mCodComune & "') AND ((ValoriCoefficienti507UDComunali.AnnoRiferimento)='" & mAnnoRiferimento & "')) order by ValoriCoefficienti507UDComunali.CodUtenza"
    End If
    ADO1.Refresh

    CaricaDati507 = Not ADO1.Recordset.EOF
    
    If Not CaricaDati507 Then
        DataGrid1(mTipoCoeff).RowHeight = 30
    Else
        DataGrid1(mTipoCoeff).RowHeight = 270
        DataGrid1(mTipoCoeff).Refresh
    End If
    
   Txt(COEFF_IQSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), "0.00")
   Txt(COEFF_IPSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), "0.00")

End Function


Private Sub Form_Resize()
    Dim i As Integer
    
    If WindowState = vbMinimized Then Exit Sub
    
    If pTipoCalcolo = TariffaNormalizzata Or pTipoCalcolo = TariffaCostiIndivisibili Then Frame1.Height = 0
    
    tabStrip.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
    
    For i = DataGrid1.LBound To DataGrid1.UBound
        DataGrid1(i).Move tabStrip.ClientLeft + 50, tabStrip.ClientTop + 50, tabStrip.ClientWidth - 50, tabStrip.ClientHeight - 50
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
    
    lastColumn = 0
    Txt(COEFF_IQSUD_ID).Enabled = False
    Txt(COEFF_IPSUD_ID).Enabled = False
    Cmd(2).Enabled = False
    
    Cmd(1).Enabled = False
    
    For i = 0 To DataGrid1.Count - 1
        DataGrid1(i).Visible = False
    Next
    mParteFissa = False
    Select Case tabStrip.SelectedItem.Key
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
            Cmd(1).Enabled = True
        Case "507UD"
            mTipoCoeff = tpo_coeff_507UD
            Txt(COEFF_IQSUD_ID).Enabled = True
            Txt(COEFF_IPSUD_ID).Enabled = True
            Cmd(2).Enabled = True
    
    End Select
    
    DataGrid1(mTipoCoeff).Visible = True
    
    If mTipoCoeff = tpo_coeff_507UA Or mTipoCoeff = tpo_coeff_507UD Then
        CaricaDati507
    Else
        CaricaDati
    End If


    SistemaCodiceAreaSuTAB
    
    
    
End Sub
   
Sub SistemaCodiceAreaSuTAB()

    If mTipoCoeff = tpo_coeff_KAKB Then
        tabStrip.SelectedItem.Caption = IIf(mParteFissa, "651 PF Utenza Domestica KA", "651 PV Utenza Domestica KB") & " - " & mArea
    ElseIf mTipoCoeff = tpo_coeff_KCKD Then
        tabStrip.SelectedItem.Caption = IIf(mParteFissa, "651 PF Utenza Attività KC", "651 PV Utenza Attività KD")
    ElseIf mTipoCoeff = tpo_coeff_507UA Then
        tabStrip.SelectedItem.Caption = "652 Utenza Attività"
    ElseIf mTipoCoeff = tpo_coeff_507UD Then
        tabStrip.SelectedItem.Caption = "652 Utenza Domestica"
    End If
End Sub

Sub CalcolaIPS()
    Dim rsKD As ADODB.Recordset
    Dim KM As Double
    Dim rsDatiTecniciComune As ADODB.Recordset
    Dim rsANA As ADODB.Recordset
    Dim mParAbitanti As String
    Dim i  As Integer
    Dim iBookmark As Variant
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
        Txt(COEFF_IPSUD_ID).Text = Format$(IPSUD, "0.00")
    Else
        'KD
        Set rsKD = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    
        i = 0
        gDisableError = True
        iBookmark = ADO1.Recordset.Bookmark
        ADO1.Recordset.MoveFirst
        On Error Resume Next
        Do While Not rsKD.EOF
            If ADO1.Recordset.fields("codattivita") = rsKD("CodTipologiaAttivita") Then
                ADO1.Recordset.fields("IPS") = rsKD("Valore") / KM
                ADO1.Recordset.MoveNext
            End If
            rsKD.MoveNext
        Loop
        gDisableError = False
        ADO1.Recordset.Bookmark = iBookmark
    End If
End Sub

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_Validate(index As Integer, Cancel As Boolean)
    If index = COEFF_IQSUD_ID And ((CDbl(Txt(COEFF_IQSUD_ID)) > MaxCoeff Or CDbl(Txt(COEFF_IQSUD_ID)) < MinCoeff) And CDbl(Txt(COEFF_IQSUD_ID)) <> 0) Then
        If MsgBox("Valore IQS non ammesso nel range " & MinCoeff & " - " & MaxCoeff & vbCrLf & " Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
    
End Sub

