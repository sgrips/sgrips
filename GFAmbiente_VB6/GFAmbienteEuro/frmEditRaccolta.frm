VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmEditRaccolta 
   Caption         =   "Form1"
   ClientHeight    =   11850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   11850
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   0
      Left            =   360
      ScaleHeight     =   6495
      ScaleWidth      =   6495
      TabIndex        =   1
      Top             =   2160
      Width           =   6495
      Begin VB.Frame fmeGenerale 
         Height          =   975
         Left            =   0
         TabIndex        =   49
         Top             =   5640
         Width           =   5535
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   2760
            TabIndex        =   50
            Text            =   "0"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Indice Medio Svuotamento"
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
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   2385
         End
      End
      Begin VB.Frame fmeParteFissa 
         Caption         =   "Determinazione Parte Fissa"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   0
         TabIndex        =   42
         Top             =   4200
         Width           =   5535
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo per Contenitore"
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
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   2775
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Costo al Metro Quadro (MQ)"
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
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Value           =   1  'Checked
            Width           =   2775
         End
      End
      Begin VB.Frame fmeParteVariabile 
         Caption         =   "Determinazione Parte Variabile"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   5535
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Calcolo raccolta RD"
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
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Calcolo raccolta RSU"
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
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid dtgPercRD 
            Height          =   2535
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4471
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   8
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
               DataField       =   "AnnoRiferimento"
               Caption         =   "AnnoRiferimento"
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
               DataField       =   "CodCategoria"
               Caption         =   "CodCategoria"
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
               DataField       =   "CodSottoCategoria"
               Caption         =   "CodSottoCategoria"
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
               DataField       =   "CodCer"
               Caption         =   "CodCer"
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
               DataField       =   "Descrizione"
               Caption         =   "Tipologia RD"
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
            BeginProperty Column06 
               DataField       =   "Percentuale"
               Caption         =   "Percentuale"
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
            BeginProperty Column07 
               DataField       =   "flgSimulazione"
               Caption         =   "flgSimulazione"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   2954,835
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcmbUnitaMisura 
            DataField       =   "ID"
            Height          =   360
            Left            =   2640
            TabIndex        =   40
            Top             =   690
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "Descrizione"
            BoundColumn     =   "ID"
            Text            =   "DataCombo1"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unità di misura calcolo PV"
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
            Left            =   2640
            TabIndex        =   41
            Top             =   435
            Width           =   2310
         End
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   4
      Left            =   3360
      ScaleHeight     =   3135
      ScaleWidth      =   6375
      TabIndex        =   47
      Top             =   2880
      Width           =   6375
      Begin MSDataGridLib.DataGrid dtgDensita 
         Height          =   3015
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5318
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "CodiceCer"
            Caption         =   "Cod Cer"
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
            DataField       =   "Descrizione"
            Caption         =   "Rifiuto"
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
            DataField       =   "UnitaDiMisuraDA"
            Caption         =   "U.M. Da"
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
            DataField       =   "UnitaDiMisuraA"
            Caption         =   "U.M. A"
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
            DataField       =   "Coefficiente"
            Caption         =   "Coefficiente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#0.######"
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
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4110,236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   2
      Left            =   720
      ScaleHeight     =   3015
      ScaleWidth      =   6015
      TabIndex        =   17
      Top             =   720
      Width           =   6015
      Begin VB.Frame freContenitori 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2535
         Index           =   1
         Left            =   2640
         TabIndex        =   32
         Top             =   0
         Width           =   3450
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Ok"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   2640
            TabIndex        =   37
            Top             =   1920
            Width           =   570
         End
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Aggiorna"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   105
            TabIndex        =   36
            Top             =   1920
            Width           =   930
         End
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Cancella/Ripristina"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   35
            Top             =   1920
            Width           =   1485
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Text            =   "0"
            Top             =   1320
            Width           =   3135
         End
         Begin MSDataListLib.DataCombo dcmbContenitori 
            Bindings        =   "frmEditRaccolta.frx":0000
            DataField       =   "Descrizione"
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Descrizione"
            BoundColumn     =   "ID"
            Text            =   "DataCombo1"
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Valore "
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
            Left            =   120
            TabIndex        =   39
            Top             =   1065
            Width           =   645
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Contenitori"
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
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   945
         End
      End
      Begin MSComctlLib.TreeView tvw 
         Height          =   2835
         Index           =   1
         Left            =   -120
         TabIndex        =   31
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5001
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   531
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   3
      Left            =   3480
      ScaleHeight     =   3135
      ScaleWidth      =   6375
      TabIndex        =   45
      Top             =   4080
      Width           =   6375
      Begin MSDataGridLib.DataGrid dtgContenitori 
         Height          =   3015
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5318
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
            DataField       =   "AnnoRiferimento"
            Caption         =   "AnnoRiferimento"
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
            DataField       =   "ID"
            Caption         =   "ID"
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
            DataField       =   "TipoUtenza"
            Caption         =   "Utenza"
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
            DataField       =   "Descrizione"
            Caption         =   "Contenitore"
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
            DataField       =   "CostoUnitario"
            Caption         =   "Costo Unitario"
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
         BeginProperty Column06 
            DataField       =   "flgSimulazione"
            Caption         =   "flgSimulazione"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2385,071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   2160
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdodcUnitaMisura 
      Height          =   330
      Left            =   0
      Top             =   8880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc AdodcContenitori 
      Height          =   330
      Left            =   0
      Top             =   9240
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   4
      CommandType     =   1
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
      Connect         =   "select * from contenitori"
      OLEDBString     =   "select * from contenitori"
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
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   1
      Left            =   4440
      ScaleHeight     =   3135
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.Frame freContenitori 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2535
         Index           =   0
         Left            =   2760
         TabIndex        =   23
         Top             =   0
         Width           =   3450
         Begin MSDataListLib.DataCombo dcmbContenitori 
            Bindings        =   "frmEditRaccolta.frx":0019
            DataField       =   "Descrizione"
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Descrizione"
            BoundColumn     =   "ID"
            Text            =   "DataCombo1"
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
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Text            =   "0"
            Top             =   1320
            Width           =   3135
         End
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Cancella/Ripristina"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   26
            Top             =   1920
            Width           =   1485
         End
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Aggiorna"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   105
            TabIndex        =   25
            Top             =   1920
            Width           =   930
         End
         Begin VB.CommandButton cmdContenitori 
            Caption         =   "&Ok"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   2640
            TabIndex        =   24
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Contenitori"
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
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Valore "
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
            Left            =   120
            TabIndex        =   28
            Top             =   1065
            Width           =   645
         End
      End
      Begin MSComctlLib.TreeView tvw 
         Height          =   2835
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5001
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   531
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSAdodcLib.Adodc ADODC 
      Height          =   330
      Left            =   0
      Top             =   9600
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   1455
      Left            =   6960
      TabIndex        =   0
      Top             =   7320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Generale"
            Key             =   "GEN"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte Variabile"
            Key             =   "PV"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte Fissa"
            Key             =   "PF"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contenitori"
            Key             =   "CONT"
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
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10800
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
         TabIndex        =   16
         Top             =   240
         Width           =   1245
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
         TabIndex        =   15
         Top             =   1800
         Width           =   435
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
         TabIndex        =   14
         Top             =   1800
         Width           =   945
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
         TabIndex        =   13
         Top             =   840
         Width           =   915
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         TabIndex        =   11
         Top             =   1800
         Width           =   1005
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
         TabIndex        =   10
         Top             =   840
         Width           =   975
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
         TabIndex        =   9
         Top             =   1800
         Width           =   540
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
         TabIndex        =   8
         Top             =   240
         Width           =   585
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
         TabIndex        =   7
         Top             =   1320
         Width           =   960
      End
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
         TabIndex        =   6
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   3
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salva"
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   2
      Top             =   9480
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcContenitoriUpdate 
      Height          =   330
      Left            =   7200
      Top             =   8280
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   4
      CommandType     =   1
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
      Connect         =   "select * from contenitori"
      OLEDBString     =   "select * from contenitori"
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
   Begin MSAdodcLib.Adodc AdodcDensita 
      Height          =   330
      Left            =   1800
      Top             =   9600
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   4
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
      RecordSource    =   $"frmEditRaccolta.frx":0032
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
Attribute VB_Name = "frmEditRaccolta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT = "Root"
Private Const UA = "UA"
Private Const UD = "UD"
Private Const SEPARATOREVALORE = " - "

Private oNode As Node
Private Livello As Long
Private lvw As ListView
Private mCodComune As String
Private mAnnoRiferimento As String
Private InitForm As Boolean
Private noEvent As Boolean

Private Sub chk_Click(index As Integer)
    Select Case index
        Case 0, 3
            chk(index).value = 1
        Case 1
            dtgPercRD.Enabled = chk(index).value = 1
            If dtgPercRD.Enabled Then InsertCodiciCerRD
        Case 2
            tvw(1).Enabled = chk(index).value = 1
            MostraTabParteFissa tvw(1).Enabled
    End Select
End Sub


Private Sub MostraTabParteFissa(Visible As Boolean)
    Dim tabobj  As MSComctlLib.Tab
    
    On Error Resume Next
    pic(2).Visible = False
    If Not Visible Then
        tabStrip.Tabs.Remove (3)
    Else
        Set tabobj = tabStrip.Tabs.Add(3, "PF", "Dati Parte Fissa")
    End If
End Sub


Private Sub cmdContenitori_Click(index As Integer)
    Dim oNodeP As Node
    Dim oNodeP1 As Node
    Dim i As Integer
    
    Select Case index
        Case 0, 5
            'Agg
            GestioneNodo tvw(tabStrip.Tag - 1), False
            
        Case 1, 3
            'Cancella
            If Not oNode Is Nothing Then
            
                Select Case Livello
                    Case 3
                       For i = 1 To 2
                            If i = 1 Then
                                Set oNodeP1 = oNode.Child
                            Else
                                Set oNodeP1 = oNodeP1.Next
                            End If
                            
                            Set oNodeP = oNodeP1.Child
                            Do While Not oNodeP Is Nothing
                                oNodeP.ForeColor = IIf(oNodeP.ForeColor = vbRed, vbBlack, vbRed)
                                Set oNodeP = oNodeP.Next
                            Loop
                        Next
                    
                    Case 4
                        oNode.ForeColor = IIf(oNode.ForeColor = vbRed, vbBlack, vbRed)
                End Select
                
                                
                
            End If
        Case 2, 4
            'Modifica
            GestioneNodo tvw(tabStrip.Tag - 1), True
    End Select
    
    

End Sub

Private Sub Command1_Click(index As Integer)
    Dim value As Integer
    Dim value1 As Integer
    
    Select Case index
        Case 0
            'annulla
            
        Case 1
        

        
            If chk(1).value = 0 Then
                gDBC.Execute ("Delete from PercentualePesiRDComunali where CodComune='" & lblCampi(1) & "' and annoRiferimento='" & lblCampi(4) & "' and flgsimulazione=false")
            Else
                If Not ValidateRD Then
                    On Error Resume Next
                    dtgPercRD.SetFocus
                    Exit Sub
                End If
                On Error GoTo Command1_ClickERR
            End If
            
            '
            If dcmbUnitaMisura.BoundText = "SVU" Then
                gDBC.Execute ("Delete from ContenitoriSvuotamentiCAComunali  where idContenitore<>0")
                gDBC.Execute ("Delete from ContenitoriSvuotamentiUDComunali  where idContenitore<>0")
            Else
                gDBC.Execute ("Delete from ContenitoriSvuotamentiCAComunali  where idContenitore=0")
                gDBC.Execute ("Delete from ContenitoriSvuotamentiUDComunali  where idContenitore=0")
            End If
            
            'Conferma
            SalvaSuDB ContenitoriSingoliSvuotamentiUD, tvw(0)
            SalvaSuDB ContenitoriInDotazioneUD, tvw(1)
            
            If chk(0).value = 1 And chk(1).value = 0 Then
                value = 0
            Else
                value = 1
            End If
            
            If chk(2).value = 1 And chk(3).value = 0 Then
                value1 = 0
            Else
                value1 = 1
            End If
            
            lvw.SelectedItem.ListSubItems("ModalitaCalcoloPV") = value
            lvw.SelectedItem.ListSubItems("ModalitaCalcoloPF") = value1
            lvw.SelectedItem.ListSubItems("UnitaDiMisuraPV") = dcmbUnitaMisura.BoundText
            
            LetParamValue mAnnoRiferimento, mCodComune, "COEFFMEDSVULT", True, CDbl(Txt(2).Text)
            
    End Select
    
    Unload Me
    

    Exit Sub
Command1_ClickERR:
     MsgBox Err.description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub dcmbUnitaMisura_Click(Area As Integer)
    Dim rs As ADODB.Recordset
    
 
    If noEvent Then Exit Sub
    
    If Area = dbcAreaList Then
        
        If Not InitForm And _
            ((dcmbUnitaMisura.BoundText = "SVU" And dcmbUnitaMisura.Tag <> "SVU") Or _
            (dcmbUnitaMisura.BoundText <> "SVU" And dcmbUnitaMisura.Tag = "SVU")) Then
            If MsgBox("Attenzione se viene modificata l'unità di misura, " + vbCrLf + "tutti i dati relativi agli svuotamenti, precedentemente inseriti andranno persi!!!", vbOKCancel) = vbCancel Then
                noEvent = True
                dcmbUnitaMisura.BoundText = dcmbUnitaMisura.Tag
                noEvent = False
                Exit Sub
            End If
            SvuotoAlbero tvw(0)
        End If
            
            
        dcmbUnitaMisura.Tag = dcmbUnitaMisura.BoundText
        Set rs = CreateObject("adodb.recordset")
        Set rs.ActiveConnection = gDBC
        
        
        If dcmbUnitaMisura.BoundText = "SVU" Then
            rs.Source = "select * from contenitori where id<>1000"
            dcmbContenitori(0).Enabled = True
        Else
            rs.Source = "select * from contenitori where id=1000"
            dcmbContenitori(0).BoundText = "SVU"
            dcmbContenitori(0).Enabled = False
        End If
        
        rs.CursorLocation = adUseClient
        rs.LockType = adLockBatchOptimistic
        rs.CursorType = adOpenStatic
        rs.Open
        Set rs.ActiveConnection = Nothing

        
        Set dcmbContenitori(0).RowSource = Nothing
        Set dcmbContenitori(0).DataSource = Nothing

        Set dcmbContenitori(0).RowSource = rs.Clone
        Set dcmbContenitori(0).DataSource = rs.Clone
        dcmbContenitori(0).Refresh
        rs.Close
        Set rs = Nothing
        
    End If
    
    
End Sub

Private Function ValidateRD() As Boolean
    Dim i As Long
    Dim Perc As Long
    
    Perc = 0
    ValidateRD = True
    For i = 0 To dtgPercRD.VisibleRows - 1
        dtgPercRD.row = i
        If dtgPercRD.Columns("Percentuale") <= 0 Or dtgPercRD.Columns("Percentuale") > 100 Then Exit For
        Perc = Perc + dtgPercRD.Columns("Percentuale")
    Next
    
    If Perc <> 100 Then
        MsgBox "Attenzione la percentuale totale deve essere 100", vbExclamation + vbOKOnly
        ValidateRD = False
    End If

    
End Function

Private Sub Form_Load()
    Dim i As Integer
    
    CentraME Me
  
    InitForm = True
    On Error Resume Next
    ADODC.Recordset.Close
    AdodcContenitori.Recordset.Close
    AdodcContenitoriUpdate.Recordset.Close
    
    On Error GoTo 0
    ADODC.CommandType = adCmdText
    ADODC.RecordSource = "select PerPesiRDCom.*,CodiciCer.Descrizione from PercentualePesiRDComunali as PerPesiRDCom  " & _
                " inner join CodiciCer on CodiciCer.CodCategoria = PerPesiRDCom.CodCategoria" & _
                " and CodiciCer.CodSottoCategoria = PerPesiRDCom.CodSottoCategoria" & _
                " and CodiciCer.CodCer = PerPesiRDCom.CodCer"
    
    ADODC.ConnectionString = gDBC.ConnectionString
    ADODC.Refresh

    
    AdodcContenitori.CommandType = adCmdText
    AdodcContenitori.RecordSource = "select * from contenitori where id<>0"
    AdodcContenitori.ConnectionString = gDBC.ConnectionString
    AdodcContenitori.Refresh

    AdodcUnitaMisura.ConnectionString = gDBC.ConnectionString
    AdodcUnitaMisura.CommandType = adCmdTable
    AdodcUnitaMisura.RecordSource = "UnitaDiMisura"
    AdodcUnitaMisura.Refresh
    
    
    AdodcDensita.ConnectionString = gDBC.ConnectionString
    AdodcDensita.CommandType = adCmdText
    AdodcDensita.RecordSource = "SELECT UnitaDiMisuraConvesioni.CodCategoria + UnitaDiMisuraConvesioni.CodSottoCategoria + " & _
                                    "UnitaDiMisuraConvesioni.CodCer as CodiceCer, CodiciCer.Descrizione, UnitaDiMisuraConvesioni.UnitaDiMisuraDA, " & _
                                    "UnitaDiMisuraConvesioni.UnitaDiMisuraA, UnitaDiMisuraConvesioni.Coefficiente FROM CodiciCer " & _
                                    "INNER JOIN UnitaDiMisuraConvesioni ON (CodiciCer.CodCer = UnitaDiMisuraConvesioni.CodCer) AND " & _
                                    "(CodiciCer.CodSottoCategoria = UnitaDiMisuraConvesioni.CodSottoCategoria) AND " & _
                                    "(CodiciCer.CodCategoria = UnitaDiMisuraConvesioni.CodCategoria)"
    AdodcDensita.Refresh
    
    
    
    Set dcmbUnitaMisura.RowSource = AdodcUnitaMisura
    Set dcmbUnitaMisura.DataSource = AdodcUnitaMisura
    Set dcmbContenitori(1).RowSource = AdodcContenitori.Recordset
    Set dcmbContenitori(1).DataSource = AdodcContenitori.Recordset

    Set dtgDensita.DataSource = AdodcDensita
    Set dtgPercRD.DataSource = ADODC
  
    frmDatiRaccolta.Visible = False
    dtgPercRD.Enabled = False
    
    
    Txt(2).Text = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFMEDSVULT")
    
    For i = 0 To pic.Count - 1
        pic(i).Visible = False
    Next
    InitForm = False
End Sub

Private Sub InsertCodiciCerRD()
    Dim Cmd As ADODB.Command
    
    If ADODC.Recordset.RecordCount = 0 Then



        Set Cmd = gADOX.Procedures("InsPercentualePesiRDComunali").Command
        Cmd.Parameters.Refresh
        Cmd.Parameters("pCodComune") = mCodComune
        Cmd.Parameters("pAnnoRiferimento") = mAnnoRiferimento
        Cmd.Execute
        
        Cmd.ActiveConnection.BeginTrans
        Cmd.Execute
        Cmd.ActiveConnection.CommitTrans
        Set Cmd.ActiveConnection = Nothing
        Set Cmd = Nothing
        
        DoEvents
        
        ADODC.Recordset.Close
        ADODC.Refresh
        

        
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Cancello quelli con percentuale a ZERO
    gDBC.Execute ("Delete from PercentualePesiRDComunali  where Percentuale=0")
    gDBC.Execute ("Delete from ContenitoriInDotazioneCAComunali  where valore=0")
    gDBC.Execute ("Delete from ContenitoriInDotazioneUDComunali  where valore=0")
    gDBC.Execute ("Delete from ContenitoriSvuotamentiCAComunali  where valore=0")
    gDBC.Execute ("Delete from ContenitoriSvuotamentiUDComunali  where valore=0")
    
    
    If UnloadMode = 0 Then Command1_Click 0
    
    
End Sub

Private Sub Form_Resize()
    ResizeControl
End Sub

Private Sub ResizeControl()
    Dim i As Integer
    Dim Left As Long
    Dim Top As Long
    
    Top = Frame1.Height
    Left = 60
    
    Frame1.Move Left, 0, Me.ScaleWidth - 60
    lblCampi(4).Left = Frame1.Width - lblCampi(4).Width - 60
    lblCampi(23).Left = lblCampi(4).Left - lblCampi(23).Width - 60
    
    
    tabStrip.Move Left, Top, Me.Width - 50, Me.ScaleHeight - Top - Command1(0).Height - 60 * 2
    
    Command1(0).Move Me.ScaleWidth - Command1(0).Width - 60, Me.ScaleHeight - Command1(0).Height - 60
    For i = 1 To Command1.Count - 1
        Command1(i).Move Command1(i - 1).Left - Command1(i).Width - 60, Command1(i - 1).Top
    Next
    
    For i = 0 To pic.Count - 1
        pic(i).Move tabStrip.ClientLeft, tabStrip.ClientTop, tabStrip.ClientWidth, tabStrip.ClientHeight
    Next
    
    fmeParteVariabile.Width = pic(0).ScaleWidth - 120
    dtgPercRD.Width = fmeParteVariabile.Width - 60 * 2 - dtgPercRD.Left
    
    dtgPercRD.Columns("Tipologia RD").Width = dtgPercRD.Width * 0.75
    dtgPercRD.Columns("Percentuale").Width = dtgPercRD.Width * 0.2
     
    
    freContenitori(0).Left = tabStrip.ClientWidth - 60 - freContenitori(0).Width
    freContenitori(1).Left = tabStrip.ClientWidth - 60 - freContenitori(0).Width
    
    
    tvw(0).Move 0, 0, freContenitori(0).Left - 60, tabStrip.ClientHeight
    tvw(1).Move 0, 0, freContenitori(1).Left - 60, tabStrip.ClientHeight
    
    fmeParteFissa.Move fmeParteVariabile.Left, fmeParteVariabile.Height + fmeParteVariabile.Top + 60, fmeParteVariabile.Width
    fmeGenerale.Move fmeParteFissa.Left, fmeParteFissa.Height + fmeParteFissa.Top + 60, fmeParteFissa.Width
    
    
    dtgContenitori.Width = pic(3).ScaleWidth - 120 - dtgContenitori.Left
    dtgContenitori.Height = pic(3).ScaleHeight - 120 - dtgContenitori.Top
    dtgContenitori.Columns("Contenitore").Width = dtgPercRD.Width * 0.5
    dtgContenitori.Columns("Costo Unitario").Width = dtgPercRD.Width * 0.25
    
    dtgDensita.Width = pic(4).ScaleWidth - 120 - dtgDensita.Left
    dtgDensita.Height = pic(4).ScaleHeight - 120 - dtgDensita.Top
    dtgDensita.Columns("Cod Cer").Width = dtgDensita.Width * 0.1
    dtgDensita.Columns("Rifiuto").Width = dtgDensita.Width * 0.5
    dtgDensita.Columns("U.M. Da").Width = dtgDensita.Width * 0.08
    dtgDensita.Columns("U.M. A").Width = dtgDensita.Width * 0.08
    dtgDensita.Columns("Coefficiente").Width = dtgDensita.Width * 0.2
    
    'TabStrip.Tag = 0
End Sub


Public Sub Edita(plvw As ListView)
    Dim i As Integer
    Dim qry As ADODB.Command
    
    Set tvw(0).ImageList = MDIMain.imglist
    Set tvw(1).ImageList = MDIMain.imglist

    Set lvw = plvw
    lblCampi(0) = lvw.SelectedItem.SubItems(1)
    lblCampi(1) = lvw.SelectedItem.Text
    lblCampi(4) = lvw.SelectedItem.SubItems(2)
    lblCampi(2) = lvw.SelectedItem.SubItems(4)
    lblCampi(3) = lvw.SelectedItem.SubItems(5)
    lblCampi(25) = lvw.SelectedItem.SubItems(3)
    
    
    mCodComune = lblCampi(1)
    mAnnoRiferimento = lblCampi(4)
    
    
    Set qry = gADOX.Procedures("InsContenitoriComunali").Command
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = mCodComune
    qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
    qry.Parameters("pTipoUtenza") = "UA"
    qry.Execute

    Set qry.ActiveConnection = Nothing
    Set qry = Nothing
    
    Set qry = gADOX.Procedures("InsContenitoriComunali").Command
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = mCodComune
    qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
    qry.Parameters("pTipoUtenza") = "UD"
    qry.Execute

    Set qry.ActiveConnection = Nothing
    Set qry = Nothing
    
    
    AdodcContenitoriUpdate.CommandType = adCmdText
    AdodcContenitoriUpdate.RecordSource = "select ConCom.*,Con.Descrizione from ContenitoriComunali as ConCom  " & _
                                          "inner join contenitori as Con on  Con.id=ConCom.id  " & _
                                          "where flgSimulazione=false and CodComune='" & mCodComune & "' and AnnoRiferimento='" & mAnnoRiferimento & "' " & _
                                          "order by ConCom.TipoUtenza,ConCom.id "
    AdodcContenitoriUpdate.ConnectionString = gDBC.ConnectionString
    AdodcContenitoriUpdate.Refresh
    
    
    Set dtgContenitori.DataSource = AdodcContenitoriUpdate
    
'    For i = 0 To txt.Count - 1
'        txt(i) = lvw.SelectedItem.SubItems(i + 6)
'    Next i
    
    Dim oNodeP As Node
    
    tvw(0).Visible = False
    tvw(0).Nodes.Clear
    Set oNodeP = tvw(0).Nodes.Add(, tvwFirst, ROOT, "Utenze", ICON_DOMESTICA_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw(0).Nodes.Add(ROOT, tvwChild, UA, "Utenze Attivita", ICON_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw(0).Nodes.Add(ROOT, tvwChild, UD, "Utenze Domestica", ICON_DOMESTICA)
    oNodeP.Expanded = True
        
    tvw(1).Visible = False
    tvw(1).Nodes.Clear
    Set oNodeP = tvw(1).Nodes.Add(, tvwFirst, ROOT, "Utenze", ICON_DOMESTICA_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw(1).Nodes.Add(ROOT, tvwChild, UA, "Utenze Attivita", ICON_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw(1).Nodes.Add(ROOT, tvwChild, UD, "Utenze Domestica", ICON_DOMESTICA)
    oNodeP.Expanded = True
    
    CreaAlbero ContenitoriSingoliSvuotamentiUA, tvw(0)
    
    CreaAlbero ContenitoriSingoliSvuotamentiUD, tvw(0)

    CreaAlbero ContenitoriInDotazioneUD, tvw(1)
    
    CreaAlbero ContenitoriInDotazioneUA, tvw(1)

    tvw(0).Visible = True
    tvw(1).Visible = True
    
    
    chk(0).value = 1
    chk(1).value = lvw.SelectedItem.ListSubItems("ModalitaCalcoloPV")
    chk(2).value = 1
    chk(3).value = lvw.SelectedItem.ListSubItems("ModalitaCalcoloPF")
    
    InitForm = True
    dcmbUnitaMisura.BoundText = lvw.SelectedItem.ListSubItems("UnitaDiMisuraPV")
    
    dcmbUnitaMisura_Click AreaConstants.dbcAreaList
    InitForm = False
    
    tabStrip.Tag = 1
    tabStrip.SelectedItem = tabStrip.Tabs(1)
    frmEditRaccolta.Show vbModal
  
End Sub

Sub CreaAlbero(ByVal Tipo As TpoContenitoriSvuotaInDotazione, tvw As TreeView)
    Dim oNodeP As Node
    Dim i As Long
    Dim KEYNODE As String
    Dim Key As String
    Dim rs As Recordset
    Dim qry As Command
    Dim sQry As String
    Dim sTable As String
    Dim pKeySave As String
    Dim pKey As String
    
    
    Set oNodeP = Nothing
    
    On Error GoTo 0
    
    
    Select Case Tipo
        Case ContenitoriSingoliSvuotamentiUD
            sTable = "ContenitoriSvuotamentiUDComunali"
            sQry = "InsUtenzeDomesticheInContenitoriSvuotamentoUD"
        Case ContenitoriSingoliSvuotamentiUA
            sTable = "ContenitoriSvuotamentiCAComunali"
            sQry = "InsCategorieAttivitaInContenitoriSvuotamentoUA"
    
        Case ContenitoriInDotazioneUA
             sTable = "ContenitoriInDotazioneCAComunali"
             sQry = "InsCategorieAttivitaInContenitoriInDotazioneCA"
        Case ContenitoriInDotazioneUD
            sTable = "ContenitoriInDotazioneUDComunali"
            sQry = "InsUtenzeDomesticheInContenitoriInDotazioneUD"
    End Select
    
    'Inserisce se vi sono le nuove Utenze Domestiche
    Set qry = gADOX.Procedures(sQry).Command
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = mCodComune
    qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
    qry.Execute
    
    Set rs = SelezionaContenitoriSvuotamentiIncarico(Tipo, mCodComune, mAnnoRiferimento)

    pKeySave = CreaLaPkey(sTable)
    
    Do While Not rs.EOF
        On Error Resume Next
        If Tipo = ContenitoriSingoliSvuotamentiUA Or _
           Tipo = ContenitoriInDotazioneUA Then
            KEYNODE = UA
            Key = rs("codAttivita")
        ElseIf Tipo = ContenitoriSingoliSvuotamentiUD Or _
               Tipo = ContenitoriInDotazioneUD Then
            KEYNODE = UD
            Key = rs("IdNumeriComponenti")
        End If
        
        Set oNode = tvw.Nodes(KEYNODE & SEPARATORE & "Liv1" & SEPARATORE & Key)
        If Err > 0 Then
            Set oNodeP = tvw.Nodes.Add(KEYNODE, tvwChild, KEYNODE & SEPARATORE & "Liv1" & SEPARATORE & Key, rs("Descrizione"), IIf(KEYNODE = UA, ICON_FACTORY, ICON_MAN))
            oNodeP.Expanded = False
            Err.Clear
            
            pKey = pKeySave
            For i = 0 To rs.fields.Count - 1
                If InStr(pKeySave, SEPARATORE & rs.fields(i).Name) And rs.fields(i).Name <> "IdContenitore" Then
                    MettivaloreInPKEY pKey, SEPARATORE & rs.fields(i).Name, rs.fields(i).value, rs.fields(i).Type
                End If
            Next
            
            
            oNodeP.Tag = pKey
            
            
            oNodeP.Sorted = True
        End If
        
        If rs("valore") > 0 Then
            Set oNode = tvw.Nodes(oNodeP.Key)
            Set oNode = tvw.Nodes.Add(oNodeP.Key, tvwChild, oNodeP.Key & "Liv2" & SEPARATORE & rs("idContenitore"), rs("DesCont") & SEPARATOREVALORE & rs("valore"))
            
            AggiornoTotalePerNodo oNodeP, rs("valore") * rs("idContenitore")
            
            
            pKey = pKeySave
            For i = 0 To rs.fields.Count - 1
                If InStr(pKeySave, SEPARATORE & rs.fields(i).Name) Then
                    MettivaloreInPKEY pKey, SEPARATORE & rs.fields(i).Name, rs.fields(i).value, rs.fields(i).Type
                End If
            Next
            oNode.Tag = pKey
            oNodeP.ForeColor = vbBlue
        End If
        

        
        oNode.Expanded = False
        
        rs.MoveNext
        

    Loop

    Set qry = Nothing
    Set rs = Nothing


End Sub

Private Sub AggiornoTotalePerNodo(oNode As Node, value As Double)
    Dim i As Integer
    Dim ii As Integer
    Dim tot As Double
    
    i = InStr(1, oNode.Text, "Totale")
    If i > 0 Then
        ii = InStrRev(oNode.Text, " ")
        tot = Mid$(oNode.Text, ii + 1) + value
        oNode.Text = Trim$(Left$(oNode.Text, i - 1)) & " Totale: " & tot
    Else
        oNode.Text = oNode.Text & " Totale: " & value
    End If
End Sub

Private Sub tvw_NodeClick(index As Integer, ByVal Node As MSComctlLib.Node)
    Set oNode = Nothing
    Txt(index) = 0
    Livello = CheLivello(Node.FullPath, tvw(index))
    lbl(index) = "Valore"
    cmdContenitori(0).Enabled = False
    cmdContenitori(1).Enabled = False
    cmdContenitori(2).Enabled = False
    

    Select Case Livello
        Case 1
        Case 2
        Case 3
            Set oNode = Node
            
            cmdContenitori(0).Enabled = True
            cmdContenitori(1).Enabled = False
            cmdContenitori(2).Enabled = False
            
            cmdContenitori(5).Enabled = True
            cmdContenitori(3).Enabled = False
            cmdContenitori(4).Enabled = False
            
        Case 4
            Set oNode = Node

            Txt(index) = PrendiValore(Node.Text)
            dcmbContenitori(index).BoundText = PrendiContenitore(Node.Key)

            cmdContenitori(0).Enabled = False
            cmdContenitori(1).Enabled = True
            cmdContenitori(2).Enabled = True
            
            cmdContenitori(5).Enabled = False
            cmdContenitori(3).Enabled = True
            cmdContenitori(4).Enabled = True
            
    End Select
    
End Sub

Function CheLivello(ByVal s As String, tvw As TreeView) As Long
    Dim i As Long
    
    i = InStr(1, s, tvw.PathSeparator)
    CheLivello = 1
    Do While i > 0
        CheLivello = CheLivello + 1
        s = Mid$(s, i + 1)
        i = InStr(1, s, tvw.PathSeparator)
    Loop
End Function


Function PrendiValore(ByVal s As String) As String
    Dim ii As Long
    
    ii = InStr(1, s, SEPARATOREVALORE)
    PrendiValore = Mid$(Replace(s, ",", "."), ii + Len(SEPARATOREVALORE), Len(s) - ii)

End Function

Function PrendiContenitore(ByVal s As String) As String
    Dim i As Long
    i = InStrRev(s, "&")
    
    PrendiContenitore = Mid$(s, i + 1)
End Function


Private Sub TabStrip_Click()

    pic(tabStrip.Tag).Visible = False
    Select Case tabStrip.SelectedItem.Key
        Case "GEN"
            pic(0).Visible = True
            tabStrip.Tag = 0
        Case "PV"
            pic(1).Visible = True
            tabStrip.Tag = 1
        Case "PF"
            pic(2).Visible = True
            tabStrip.Tag = 2
        Case "CONT"
            pic(3).Visible = True
            tabStrip.Tag = 3
        Case "DENS"
            pic(4).Visible = True
            tabStrip.Tag = 4
    End Select
    

    
    
End Sub


Sub GestioneNodo(tvw As TreeView, ByVal Modifica As Boolean)
    Dim oNodeP As Node
    Dim oNodeIns As Node
    Dim Key As String
    Dim sMessage  As String
    Dim Des As String
    Dim ii As Integer
    Dim value As Integer
    
    Key = dcmbContenitori(tvw.index).BoundText
    Des = dcmbContenitori(tvw.index).Text
    'controllo valori
    If Val(Txt(tvw.index).Text) = 0 Or Val(Key) = 0 Then
        MsgBox "I valori inseriti non sono validi", vbExclamation, App.Title
        Exit Sub
    End If
    
    If Not oNode Is Nothing Then
        On Error Resume Next
        Select Case Livello
            Case 3
                'livello utenza
                Set oNodeIns = oNode
            Case 4
                'livello valore
                Set oNodeIns = oNode.Parent
        End Select
        
        sMessage = "Attenzione il contenitore:" & Key & " è già presente " + vbCrLf + "per l'utenza (" & oNodeIns.Text & ")"
        
        Set oNodeP = tvw.Nodes(oNodeIns.Key & "Liv2" & SEPARATORE & Key)
        
        value = 0
        If Modifica Then
            ii = InStrRev(oNodeP.Text, "-")
            value = CInt(Mid$(oNodeP.Text, ii + 1))
        End If
        
        
        AggiornoTotalePerNodo oNodeIns, Txt(tvw.index) * Key - value * Key
        
        If Err = 0 Then
            'se ho trovato il nodo e non sono in modifca o se sono in modifica e ho modifcato la percentuale non
            'deve esistere già
            If Not Modifica Or (oNodeP.Key <> oNode.Key) Then MsgBox sMessage, vbExclamation, App.Title: Exit Sub
        End If
        
        If Modifica Then
            oNode.Text = Des & SEPARATOREVALORE & Txt(tvw.index)
            oNode.Key = oNodeIns.Key & "Liv2" & SEPARATORE & Key
        Else
            Err.Clear
            
            Set oNodeP = tvw.Nodes(oNodeIns.Key)
            oNodeIns.ForeColor = vbBlue
            
            Set oNodeP = tvw.Nodes.Add(oNodeP.Key, tvwChild, oNodeP.Key & "Liv2" & SEPARATORE & Key, Des & SEPARATOREVALORE & Txt(tvw.index))
        End If
        oNodeP.EnsureVisible
        
    End If

    
End Sub

Private Sub txt_LostFocus(index As Integer)
    If index = 1 Then Txt(1).Text = CInt(Txt(1).Text)
    
End Sub

Private Sub txt_Validate(index As Integer, Cancel As Boolean)
    If index = 0 Or index = 1 Or index = 2 Then
        If Not IsNumeric(Txt(index).Text) Then
            MsgBox "Il valore deve essere numerico", vbExclamation, App.Title: Cancel = True
        ElseIf index = 2 Then
            Dim d As Double
            d = Txt(index).Text
            If d <= 0 Or d > 1 Then
                MsgBox "Il valore deve essere >0 e < 1", vbExclamation, App.Title: Cancel = True
            End If
        End If
    End If
    
End Sub

Sub SalvaSuDB(ByVal Tipo As TpoContenitoriSvuotaInDotazione, tvw As TreeView)
    Dim i As Long
    Dim sKey As String
    Dim rs As Recordset
    Dim sTable As String
    Dim s As String
    
    For i = 1 To tvw.Nodes.Count
        Set oNode = tvw.Nodes(i)
        Select Case CheLivello(oNode.FullPath, tvw)
            Case 4
                
                
                Select Case Tipo
                    Case ContenitoriSingoliSvuotamentiUD, ContenitoriSingoliSvuotamentiUA
                        If oNode.Parent.Parent.Key = "UA" Then
                            sTable = "ContenitoriSvuotamentiCAComunali"
                        Else
                            sTable = "ContenitoriSvuotamentiUDComunali"
                        End If
                        
                    Case ContenitoriInDotazioneUA, ContenitoriInDotazioneUD
                         
                        If oNode.Parent.Parent.Key = "UA" Then
                            sTable = "ContenitoriInDotazioneCAComunali"
                        Else
                            sTable = "ContenitoriInDotazioneUDComunali"
                        End If

                End Select
                
                
                If oNode.Tag <> vbNullString Then
                    sKey = oNode.Tag
                    'cancello il record per permettere anche la variazione della percentuale
                    gDBC.Execute ("Delete from " & sTable & "  where " & GetPkeyForWhereCondition(sKey))
                End If
                
                sKey = oNode.Parent.Tag
                MettivaloreInPKEY sKey, SEPARATORE & "IdContenitore", PrendiContenitore(oNode.Key), adInteger
                
                If oNode.ForeColor <> vbRed Then
                    gDBC.Execute "Insert into " & sTable & "(" & getCampiDapKeyPerInsert(sKey) & ",valore) values(" & getValoriDapKeyPerInsert(sKey) & "," & Replace(PrendiValore(oNode.Text), ",", ".") & ")"
                End If
                
        End Select
    Next
    
    Set rs = Nothing
End Sub


Private Sub SvuotoAlbero(tvw As TreeView)
    Dim oNode As Node
    Dim i As Long
    
    For i = tvw.Nodes.Count To 1 Step -1
        Set oNode = tvw.Nodes(i)
        If CheLivello(oNode.FullPath, tvw) = 4 Then
            oNode.Parent.Text = Mid$(oNode.Parent.Text, InStr(oNode.Parent.Text, "Totale") - 1)
            oNode.Parent.ForeColor = vbBlack
            tvw.Nodes.Remove oNode.Key
        End If
    Next
End Sub

