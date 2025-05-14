VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCodiciCER 
   Caption         =   "Lista Rifiuti"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   13230
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "Associa"
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
      Left            =   7680
      TabIndex        =   16
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ricerca"
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
      Left            =   9000
      TabIndex        =   13
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ComboBox cmbGrid 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Text            =   "cmbGrid"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Esportazione"
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
      Index           =   1
      Left            =   10320
      TabIndex        =   11
      Top             =   8160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   240
      ScaleHeight     =   2895
      ScaleWidth      =   12615
      TabIndex        =   2
      Top             =   120
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmCodiciCER.frx":0000
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
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
         DataMember      =   "cmdCodiciCER"
         ColumnCount     =   34
         BeginProperty Column00 
            DataField       =   "CategoriaCERDes"
            Caption         =   "CategoriaCERDes"
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
            DataField       =   "SottoCategoriaCERDES"
            Caption         =   "SottoCategoriaCERDES"
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
         BeginProperty Column03 
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
            DataField       =   "CodSubCer"
            Caption         =   "CodSubCer"
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
         BeginProperty Column07 
            DataField       =   "CodCer1"
            Caption         =   "Codcer1"
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
         BeginProperty Column08 
            DataField       =   "CodSottoCategoria1"
            Caption         =   "codsottocategoria1"
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
         BeginProperty Column09 
            DataField       =   "CodCategoria1"
            Caption         =   "codcategoria"
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
         BeginProperty Column10 
            DataField       =   "codiceCER"
            Caption         =   "Codice Cer"
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
         BeginProperty Column11 
            DataField       =   "Descrizione"
            Caption         =   "Tipo Rifiuto"
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
         BeginProperty Column12 
            DataField       =   "RifiutoPericoloso"
            Caption         =   "Pericoloso"
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
         BeginProperty Column13 
            DataField       =   "UnitaDiMisura"
            Caption         =   "U.M."
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
         BeginProperty Column14 
            DataField       =   "codStatoFisico"
            Caption         =   "Stato Fisico"
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
         BeginProperty Column15 
            DataField       =   "FlgAbilitatoInRicerca"
            Caption         =   "Gestito "
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
         BeginProperty Column16 
            DataField       =   "FlgBonus"
            Caption         =   "Attivo Bonus"
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
         BeginProperty Column17 
            DataField       =   "CorrispettivoBonus"
            Caption         =   "Corrispettivo Bonus"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """€"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "flgBonusAPunti"
            Caption         =   "Attivo Bonus Punti"
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
         BeginProperty Column19 
            DataField       =   "BonusPunti"
            Caption         =   "Corrispettivo Bonus Punti"
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
         BeginProperty Column20 
            DataField       =   "flgIngombrante"
            Caption         =   "Ingom brante"
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
         BeginProperty Column21 
            DataField       =   "flgAmmessoPerUD"
            Caption         =   "Ammesso Quando UD"
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
         BeginProperty Column22 
            DataField       =   "qtyMaxGGUD"
            Caption         =   "qty Max GG UD"
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
         BeginProperty Column23 
            DataField       =   "qtyMaxAAUD"
            Caption         =   "qty Max AA UD"
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
         BeginProperty Column24 
            DataField       =   "flgAmmessoPerUA"
            Caption         =   "Ammesso Quando UA"
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
         BeginProperty Column25 
            DataField       =   "qtyMaxGGUA"
            Caption         =   "qty Max GG UA"
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
         BeginProperty Column26 
            DataField       =   "qtyMaxAAUA"
            Caption         =   "qty Max AA UA"
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
         BeginProperty Column27 
            DataField       =   "FlgAbilitatoFatturazione"
            Caption         =   "Attivo Per Fatt."
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
         BeginProperty Column28 
            DataField       =   "qtyMaxFatturabile"
            Caption         =   "qty Max Fatt."
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
         BeginProperty Column29 
            DataField       =   "costounitario"
            Caption         =   "Costo Unitario"
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
         BeginProperty Column30 
            DataField       =   "ggMAssimiGiacenza"
            Caption         =   "GG Max Giac Gestionale"
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
         BeginProperty Column31 
            DataField       =   "ggMAssimiGiacenzaNormativa"
            Caption         =   "GG Max Giac Normativa"
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
         BeginProperty Column32 
            DataField       =   "qtyMassimaGiacenza"
            Caption         =   "qty Max. Giac Gestionale"
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
         BeginProperty Column33 
            DataField       =   "qtyMassimaGiacenzaNormativa"
            Caption         =   "qty Max. Giac Normativa"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1785,26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   5520,189
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column13 
               Button          =   -1  'True
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column17 
               Alignment       =   1
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column22 
               Alignment       =   1
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column23 
               Alignment       =   1
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column24 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column25 
               Alignment       =   1
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column26 
               Alignment       =   1
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column27 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column28 
               Alignment       =   1
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column29 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column30 
            EndProperty
            BeginProperty Column31 
            EndProperty
            BeginProperty Column32 
            EndProperty
            BeginProperty Column33 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   1
      Left            =   3720
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   3
      Top             =   3720
      Width           =   8895
      Begin VB.ComboBox cmbCodiciCer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbCodiciCer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   960
         Width           =   8535
      End
      Begin VB.ComboBox cmbCodiciCer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   360
         Width           =   8535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gestito"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sotto Categoria"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   825
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro"
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
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
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
      Left            =   11880
      TabIndex        =   0
      Top             =   8160
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
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
      TabIndex        =   15
      Top             =   7920
      Visible         =   0   'False
      Width           =   8340
   End
End
Attribute VB_Name = "frmCodiciCER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private rsSottoCategoria As Recordset
Private sourceElenco As String
Private rsUnitaDiMisura As Recordset
Private rsStatoFisico As Recordset

Private noEvent As Boolean

Private Sub cmbCodiciCer_Click(Index As Integer)
    
    Select Case Index
        Case 0
            creaComboSottoCategoria cmbCodiciCer(0)
    End Select
End Sub




Private Sub cmbGrid_Change()
    If noEvent Then Exit Sub
    If cmbGrid.Tag <> "" And cmbGrid.Locked = False Then
        DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 2
            ricaricaLista
        Case 1
            lblEsportazione.Visible = True
            ProgressBar1.Visible = True
            Esportazione
            lblEsportazione.Visible = False
            ProgressBar1.Visible = False
        Case 3
            frmAssociaCodiceCerAlComune.Show vbModal
    End Select
End Sub

Private Sub Esportazione()
    Dim ItemH As Column
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim mChecked As Long
    Dim i As Long
        
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gCodiciCERFile & IIf(gCodiciCERFileType = 0, ".TXT", ".CSV")
   
    If DataGrid1.ApproxCount <= 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = DevEnvMDB.rscmdCodiciCER.RecordCount
    ProgressBar1.Value = 0
    
    DataGrid1.Row = 0
    For Each ItemH In DataGrid1.Columns
        If ItemH.Visible Then
            txtStream.Write Chr$(34) & ItemH.Caption & Chr$(34)
            If ItemH.ColIndex < DataGrid1.Columns.Count Then
                txtStream.Write gCodiciCERSeparatore
            End If
        End If
    Next
    txtStream.Write vbCrLf
    

    DevEnvMDB.rscmdCodiciCER.MoveFirst
    Do While Not DevEnvMDB.rscmdCodiciCER.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        For Each ItemH In DataGrid1.Columns
            If ItemH.Visible Then
                
                txtStream.Write formatValue(ItemH)
                If ItemH.ColIndex < DataGrid1.Columns.Count Then
                    txtStream.Write gCodiciCERSeparatore
                End If
            End If
        Next
        txtStream.Write vbCrLf
        DevEnvMDB.rscmdCodiciCER.MoveNext
    Loop
    DevEnvMDB.rscmdCodiciCER.MoveFirst
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    
    fso.MoveFile fileEXP, gPercorsoEsp & "\" & gCodiciCERFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gCodiciCERFileType = 0, ".TXT", ".CSV")
    
    Set fso = Nothing
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title
End Sub

Private Function formatValue(item As Column) As Variant
    On Error Resume Next
    Dim v As Variant
    
    v = item.Value
    If Err.Number > 0 Then v = item.Text
    
    If (item.NumberFormat <> "") Then
        formatValue = toZero(v)
    Else
        formatValue = "=" & Chr$(34) & nullToBlank(v) & Chr$(34)
    End If
End Function

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    DoEvents
    
    If DataGrid1.col = -1 Or DataGrid1.Row = -1 Then Exit Sub
    
    cmbGrid.Visible = False
    cmbGrid.Move TabStrip1.ClientLeft + DataGrid1.Columns(DataGrid1.col).left + 7, _
                DataGrid1.RowTop(DataGrid1.Row) - 15 + TabStrip1.ClientTop, _
                DataGrid1.Columns(DataGrid1.col).width
                
    noEvent = True
    Select Case UCase$(DataGrid1.Columns(DataGrid1.col).DataField)
        Case UCase$("codStatoFisico")
            caricaComboStatoFisico cmbGrid, DataGrid1.col
        Case UCase$("UnitaDiMisura")
            caricaComboUnitaDiMisura cmbGrid, DataGrid1.col
        Case UCase$("FlgAbilitatoInRicerca")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("FlgAbilitatoFatturazione")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("FlgBonus")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("flgBonusAPunti")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("flgIngombrante")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("FlgAmmessoPerUD")
            caricaComboSINO cmbGrid, DataGrid1.col
        Case UCase$("FlgAmmessoPerUA")
            caricaComboSINO cmbGrid, DataGrid1.col
            
    End Select
    noEvent = False
End Sub

Private Sub caricaComboSINO(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    cmb.AddItem "Si"
    cmb.AddItem "No"
    
    cmb.Text = IIf(DataGrid1.Columns(DataGrid1.col).Value = -1, "Si", "No")
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub

Private Sub caricaComboUnitaDiMisura(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    rsUnitaDiMisura.MoveFirst
    Do While Not rsUnitaDiMisura.EOF
        cmb.AddItem rsUnitaDiMisura!codUnitaMisura
        rsUnitaDiMisura.MoveNext
    Loop
    
    cmb.Text = DataGrid1.Columns(DataGrid1.col).Value
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub

Private Sub caricaComboStatoFisico(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    rsStatoFisico.MoveFirst
    Do While Not rsStatoFisico.EOF
        cmb.AddItem rsStatoFisico!codStatoFisico
        rsStatoFisico.MoveNext
    Loop
    
    cmb.Text = DataGrid1.Columns(DataGrid1.col).Value
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub

Private Sub cmbGrid_Click()
    DoEvents
    DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
    If UCase$(DataGrid1.Columns(cmbGrid.Tag).DataField) = UCase$("flgBonus") And DataGrid1.Columns(cmbGrid.Tag).Value = 0 Then
        DataGrid1.Columns(cmbGrid.Tag + 1).Value = 0
    End If
    If UCase$(DataGrid1.Columns(cmbGrid.Tag).DataField) = UCase$("flgBonusapunti") And DataGrid1.Columns(cmbGrid.Tag).Value = 0 Then
        DataGrid1.Columns(cmbGrid.Tag + 1).Value = 0
    End If
    
End Sub


Private Sub DataGrid1_Scroll(Cancel As Integer)

    If (DataGrid1.col = -1 Or DataGrid1.Row = -1) Then Exit Sub
    cmbGrid.Move TabStrip1.ClientLeft + DataGrid1.Columns(DataGrid1.col).left + 7, _
            DataGrid1.RowTop(DataGrid1.Row) - 15 + TabStrip1.ClientTop, _
            DataGrid1.Columns(DataGrid1.col).width
            
            
    If (cmbGrid.Visible) Then cmbGrid.Visible = DataGrid1.Columns(DataGrid1.col).left > 0
End Sub

Private Sub Form_Load()
    On Error GoTo LoadERR
    Screen.MousePointer = vbHourglass
    DoEvents
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Screen.MousePointer = vbDefault
    
    creaComboCategoria
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = frmLogin.Livello <> LVL_USER
    
    cmbGrid.Locked = frmLogin.Livello <> LVL_ADMINISTRATOR
    
    cmbCodiciCer(2).AddItem "Tutti"
    cmbCodiciCer(2).AddItem "Si"
    cmbCodiciCer(2).AddItem "No"
    cmbCodiciCer(2).ListIndex = 1
    
    Set rsUnitaDiMisura = CreateObject("ADODB.Recordset")
    rsUnitaDiMisura.Source = "select * from UnitaDiMisura"
    rsUnitaDiMisura.LockType = adLockBatchOptimistic
    rsUnitaDiMisura.CursorLocation = adUseClient
    rsUnitaDiMisura.CursorType = adOpenStatic
    Set rsUnitaDiMisura.ActiveConnection = gDBC
    rsUnitaDiMisura.Open
    Set rsUnitaDiMisura.ActiveConnection = Nothing
    
    Set rsStatoFisico = CreateObject("ADODB.Recordset")
    rsStatoFisico.Source = "select * from StatoFisicoCER"
    rsStatoFisico.LockType = adLockBatchOptimistic
    rsStatoFisico.CursorLocation = adUseClient
    rsStatoFisico.CursorType = adOpenStatic
    Set rsStatoFisico.ActiveConnection = gDBC
    rsStatoFisico.Open
    Set rsStatoFisico.ActiveConnection = Nothing
    
    
    
    sourceElenco = DevEnvMDB.Commands("cmdCodiciCer").CommandText
    
    ricaricaLista

    cmbGrid.font = DataGrid1.font
    cmbGrid.font.size = DataGrid1.font.size
    
    
    Exit Sub
LoadERR:
    MsgBox "LoadCodiciCer:" & Err.Description, vbCritical, App.Title
End Sub

Private Sub creaComboCategoria()
    Dim rs As ADODB.Recordset
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    rs.Source = "select * from CategoriaCer order by codCategoria"
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    cmbCodiciCer(0).AddItem "Tutti"
    Do While Not rs.EOF
        cmbCodiciCer(0).AddItem rs!codCategoria & " - " & rs!Descrizione
        rs.MoveNext
    Loop
    
    Set rsSottoCategoria = CreateObject("ADODB.Recordset")
    rsSottoCategoria.LockType = adLockBatchOptimistic
    rsSottoCategoria.CursorLocation = adUseClient
    rsSottoCategoria.CursorType = adOpenStatic
    
    rsSottoCategoria.Source = "select * from SottoCategoriaCer order by codSottoCategoria"
    Set rsSottoCategoria.ActiveConnection = gDBC
    rsSottoCategoria.Open
    Set rsSottoCategoria.ActiveConnection = Nothing
    
    cmbCodiciCer(1).AddItem "Tutti"
    cmbCodiciCer(1).ListIndex = 0
    cmbCodiciCer(0).ListIndex = 0
End Sub


Private Sub creaComboSottoCategoria(Categoria As String)
    
    cmbCodiciCer(1).Clear
    cmbCodiciCer(1).AddItem "Tutti"
    
    If cmbCodiciCer(0).Text <> "Tutti" Then
        
        rsSottoCategoria.Filter = ""
        rsSottoCategoria.Filter = "codCategoria ='" & left(cmbCodiciCer(0).Text, InStr(1, cmbCodiciCer(0).Text, " - ") - 1) & "'"
        
        Do While Not rsSottoCategoria.EOF
            cmbCodiciCer(1).AddItem rsSottoCategoria!codSottoCategoria & " - " & rsSottoCategoria!Descrizione
            rsSottoCategoria.MoveNext
        Loop
    End If
    
    cmbCodiciCer(1).ListIndex = 0
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    resizeControl


    ProgressBar1.Move 15, Me.ScaleHeight - ProgressBar1.Height - 15, cmd(cmd.ubound).left - 30
    lblEsportazione.Move 15, ProgressBar1.top - lblEsportazione.Height - 15, cmd(cmd.ubound).left - 30
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
    
    Picture1(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    Picture1(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    
    DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSottoCategoria.Close
    Set rsSottoCategoria = Nothing
    rsUnitaDiMisura.Close
    Set rsUnitaDiMisura = Nothing
    
    If (DevEnvMDB.rscmdCodiciCER.state = adStateOpen) Then
       DevEnvMDB.rscmdCodiciCER.Close
    End If
    DevEnvMDB.Commands("cmdCodiciCER").CommandText = sourceElenco
    DevEnvMDB.cmdCodiciCer
    
    Set pfrmCodiciCER = Nothing
End Sub

Private Sub TabStrip1_Click()
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1
    
    If cmbGrid.Visible And TabStrip1.SelectedItem.Index = 2 Then cmbGrid.Visible = False
    
End Sub

Private Sub ricaricaLista()
    Dim where As String
    
    If (DevEnvMDB.rscmdCodiciCER.state = adStateOpen) Then
       DevEnvMDB.rscmdCodiciCER.Close
    End If
    
    where = ""
    
       
    DevEnvMDB.Commands("cmdCodiciCER").CommandText = sourceElenco
    
    If (cmbCodiciCer(0).Text <> "Tutti") Then
        where = where & "CodiciCerComunali.codCategoria='" & left(cmbCodiciCer(0).Text, InStr(1, cmbCodiciCer(0).Text, " - ") - 1) & "'"
    End If
    
    If (cmbCodiciCer(1).Text <> "Tutti") Then
        If (where <> "") Then where = where & " AND "
        where = where & "CodiciCerComunali.codSottoCategoria='" & left(cmbCodiciCer(1).Text, InStr(1, cmbCodiciCer(1).Text, " - ") - 1) & "'"
    End If
    
    If (cmbCodiciCer(2).Text <> "Tutti") Then
        If (where <> "") Then where = where & " AND "
        where = where & "FlgAbilitatoInRicerca=" & IIf(cmbCodiciCer(2).Text = "Si", -1, 0)
    End If
    
    where = where & IIf(where = vbNullString, "", " and ") & " CodiciCerComunali.codComune='" & GetCurrentComune!codistat & "'"
   
    
    DevEnvMDB.Commands("cmdCodiciCER").CommandText = sourceElenco & IIf(where = "", "", " WHERE " & where & " ") & " Order by CodiciCerComunali.codCategoria,CodiciCerComunali.codSottoCategoria,CodiciCerComunali.codCer,CodiciCerComunali.CodSubCer"
  
    DevEnvMDB.cmdCodiciCer
    
    
    Set DataGrid1.DataSource = Nothing
    
    Set DataGrid1.DataSource = DevEnvMDB
    DataGrid1.Refresh
    
    TabStrip1.Tabs(1).Selected = True
    
End Sub

