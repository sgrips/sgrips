VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImpianti 
   Caption         =   "Elenco Impianti"
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
      Height          =   7265
      Index           =   1
      Left            =   1440
      ScaleHeight     =   7230
      ScaleWidth      =   8505
      TabIndex        =   30
      Top             =   360
      Width           =   8535
      Begin VB.TextBox txtInterno 
         DataField       =   "Subalterno"
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
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   7
         Top             =   4650
         Width           =   420
      End
      Begin VB.TextBox txtSubalterno 
         DataField       =   "Interno"
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
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   8
         Top             =   4650
         Width           =   420
      End
      Begin VB.TextBox txtCap 
         DataField       =   "cap"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   87
         Top             =   3210
         Width           =   1020
      End
      Begin VB.ComboBox cmbAutorizzazione 
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
         Left            =   2760
         TabIndex        =   0
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtidAutorizzazione 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idAutorizzazione"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1320
         Width           =   990
      End
      Begin VB.TextBox txtCodSoggetto 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codSoggetto"
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
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   810
         Width           =   660
      End
      Begin VB.TextBox txtImpianto 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idImpianto"
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
         TabIndex        =   42
         Top             =   315
         Width           =   660
      End
      Begin VB.TextBox txtNumeroFax 
         DataField       =   "NumeroFax"
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
         MaxLength       =   15
         TabIndex        =   10
         Top             =   5655
         Width           =   2475
      End
      Begin VB.TextBox txtNumeroTelefono 
         DataField       =   "NumeroTelefono"
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
         MaxLength       =   15
         TabIndex        =   9
         Top             =   5145
         Width           =   2475
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
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
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4650
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
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
         MaxLength       =   100
         TabIndex        =   5
         Top             =   4155
         Width           =   4800
      End
      Begin VB.TextBox txtCodComune 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
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
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3645
         Width           =   990
      End
      Begin VB.TextBox txtPartitaIva 
         DataField       =   "PartitaIva"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         TabIndex        =   3
         Top             =   2790
         Width           =   2580
      End
      Begin VB.TextBox txtCodFiscale 
         DataField       =   "CodFiscale"
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
         TabIndex        =   2
         Top             =   2295
         Width           =   2640
      End
      Begin VB.TextBox txtRagioneSociale 
         DataField       =   "RagioneSociale"
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox txtidSoggetto 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSoggetto"
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
         TabIndex        =   32
         Top             =   810
         Width           =   660
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmImpianti.frx":0000
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   2760
         TabIndex        =   4
         Top             =   3645
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "Comune"
         BoundColumn     =   "CODISTAT"
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cap:"
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
         Index           =   12
         Left            =   1245
         TabIndex        =   88
         Top             =   3255
         Width           =   405
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Autorizzazione:"
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
         Index           =   11
         Left            =   270
         TabIndex        =   86
         Top             =   1365
         Width           =   1350
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
         Left            =   2505
         TabIndex        =   45
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idImpianto"
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
         Left            =   780
         TabIndex        =   43
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
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
         Index           =   8
         Left            =   1275
         TabIndex        =   41
         Top             =   5700
         Width           =   390
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
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
         Left            =   885
         TabIndex        =   40
         Top             =   5190
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nr Civico:"
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
         Index           =   6
         Left            =   840
         TabIndex        =   39
         Top             =   4695
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo:"
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
         Index           =   5
         Left            =   885
         TabIndex        =   38
         Top             =   4200
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comune:"
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
         Index           =   4
         Left            =   885
         TabIndex        =   36
         Top             =   3690
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Partita Iva:"
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
         Left            =   750
         TabIndex        =   35
         Top             =   2835
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Fiscale:"
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
         Left            =   570
         TabIndex        =   34
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ragione Sociale:"
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
         Left            =   195
         TabIndex        =   33
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idSoggetto:"
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
         Left            =   675
         TabIndex        =   31
         Top             =   855
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   12615
      TabIndex        =   21
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   240
         TabIndex        =   22
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "idImpianto"
            Caption         =   "idImpianto"
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
            DataField       =   "idSoggetto"
            Caption         =   "idSoggetto"
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
         BeginProperty Column03 
            DataField       =   "RagioneSociale"
            Caption         =   "Ragione Sociale"
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
            DataField       =   "CodFiscale"
            Caption         =   "Cod Fiscale"
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
            DataField       =   "PartitaIva"
            Caption         =   "Partita Iva"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "00000000000"
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
            DataField       =   "Comune"
            Caption         =   "Comune"
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
            DataField       =   "Indirizzo"
            Caption         =   "Indirizzo"
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
            DataField       =   "NumCivico"
            Caption         =   "Nr Civico"
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
            DataField       =   "NumeroTelefono"
            Caption         =   "Nr Telefono"
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
            DataField       =   "idAutorizzazione"
            Caption         =   "Autorizzazione"
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
            DataField       =   "NumeroFax"
            Caption         =   "Nr Fax"
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
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2670,236
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1709,858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1544,882
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2865,26
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2654,929
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1425,26
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1454,74
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   3
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   23
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   360
         Width           =   2055
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
         Index           =   3
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   4335
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
         Index           =   4
         Left            =   5640
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Fiscale"
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
         TabIndex        =   29
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Partita Iva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cognome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ragione Sociale"
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
         TabIndex        =   26
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6540
      Index           =   2
      Left            =   2520
      ScaleHeight     =   6510
      ScaleWidth      =   11265
      TabIndex        =   46
      Top             =   1320
      Width           =   11295
      Begin VB.Frame Frame1 
         Caption         =   "Smaltimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4815
         Index           =   0
         Left            =   0
         TabIndex        =   66
         Top             =   480
         Width           =   6735
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1815
            Index           =   0
            Left            =   3000
            ScaleHeight     =   1815
            ScaleWidth      =   735
            TabIndex        =   76
            Top             =   600
            Width           =   735
            Begin VB.CommandButton Command1 
               Caption         =   ">"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   80
               Tag             =   "Associa l'elemento selezionato"
               Top             =   480
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<<"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   79
               ToolTipText     =   "Disassocia tutti gli elementi"
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">>"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   78
               ToolTipText     =   "Associa tutti gli elementi"
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   77
               ToolTipText     =   "Disassocia l'elemento selezionato"
               Top             =   1440
               Width           =   495
            End
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   1
            Left            =   3840
            TabIndex        =   75
            ToolTipText     =   "Tipo operazione Smaltimento associata"
            Top             =   600
            Width           =   2655
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   0
            Left            =   240
            TabIndex        =   74
            ToolTipText     =   "Tipo operazione Smaltimento d'associare"
            Top             =   600
            Width           =   2655
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   7
            Left            =   240
            TabIndex        =   73
            ToolTipText     =   "Tipo operazione Smaltimento d'associare"
            Top             =   3000
            Width           =   2655
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   6
            Left            =   3840
            TabIndex        =   72
            ToolTipText     =   "Tipo operazione Smaltimento associata"
            Top             =   3000
            Width           =   2655
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Index           =   3
            Left            =   3000
            ScaleHeight     =   1935
            ScaleWidth      =   735
            TabIndex        =   67
            Top             =   3000
            Width           =   735
            Begin VB.CommandButton Command1 
               Caption         =   "<"
               Height          =   375
               Index           =   15
               Left            =   120
               TabIndex        =   71
               ToolTipText     =   "Disassocia l'elemento selezionato"
               Top             =   1440
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">>"
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   70
               ToolTipText     =   "Associa tutti gli elementi"
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<<"
               Height          =   375
               Index           =   13
               Left            =   120
               TabIndex        =   69
               ToolTipText     =   "Disassocia tutti gli elementi"
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   68
               Tag             =   "Associa l'elemento selezionato"
               Top             =   480
               Width           =   495
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Operazioni Associate"
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
            TabIndex        =   84
            Top             =   360
            Width           =   2010
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Operazioni da Associare"
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
            TabIndex        =   83
            Top             =   360
            Width           =   2325
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
            Index           =   7
            Left            =   240
            TabIndex        =   82
            Top             =   2760
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
            Index           =   6
            Left            =   3840
            TabIndex        =   81
            Top             =   2760
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Recupero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4815
         Index           =   1
         Left            =   4080
         TabIndex        =   47
         Top             =   480
         Width           =   6735
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   2
            Left            =   240
            TabIndex        =   61
            ToolTipText     =   "Tipo operazione Recupero da associata"
            Top             =   480
            Width           =   2655
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   3
            Left            =   3840
            TabIndex        =   60
            ToolTipText     =   "Tipo operazione Recupero associata"
            Top             =   360
            Width           =   2655
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1815
            Index           =   1
            Left            =   3000
            ScaleHeight     =   1815
            ScaleWidth      =   735
            TabIndex        =   55
            Top             =   480
            Width           =   735
            Begin VB.CommandButton Command1 
               Caption         =   "<"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   59
               ToolTipText     =   "Disassocia l'elemento selezionato"
               Top             =   1440
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">>"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   58
               ToolTipText     =   "Associa tutti gli elementi"
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<<"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   57
               ToolTipText     =   "Disassocia tutti gli elementi"
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   56
               Tag             =   "Associa l'elemento selezionato"
               Top             =   480
               Width           =   495
            End
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   4
            Left            =   240
            TabIndex        =   54
            ToolTipText     =   "Tipo operazione Recupero da associata"
            Top             =   2640
            Width           =   2655
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Index           =   2
            Left            =   2640
            ScaleHeight     =   1935
            ScaleWidth      =   735
            TabIndex        =   49
            Top             =   2520
            Width           =   735
            Begin VB.CommandButton Command1 
               Caption         =   "<"
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   53
               ToolTipText     =   "Disassocia l'elemento selezionato"
               Top             =   1440
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">>"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   52
               ToolTipText     =   "Associa tutti gli elementi"
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "<<"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   51
               ToolTipText     =   "Disassocia tutti gli elementi"
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton Command1 
               Caption         =   ">"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   50
               Tag             =   "Associa l'elemento selezionato"
               Top             =   480
               Width           =   495
            End
         End
         Begin VB.ListBox List1 
            Height          =   1815
            Index           =   5
            Left            =   3720
            TabIndex        =   48
            ToolTipText     =   "Tipo operazione Recupero associata"
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Operazioni Associate"
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
            Index           =   3
            Left            =   3960
            TabIndex        =   65
            Top             =   120
            Width           =   2010
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Operazioni da Associare"
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
            Index           =   2
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   2325
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
            Index           =   8
            Left            =   3840
            TabIndex        =   63
            Top             =   2400
            Width           =   1455
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
            Index           =   9
            Left            =   240
            TabIndex        =   62
            Top             =   2400
            Width           =   1830
         End
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   16
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
      TabIndex        =   17
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   24
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
            Caption         =   "Operazioni"
            Key             =   "OPERAZIONI"
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
Attribute VB_Name = "frmImpianti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private pidSoggetto As Long
Private pcodSoggetto As String
Private pErrore As Boolean
Private pModifiche  As Boolean
Private pInsert As Boolean


Const TXT_CODICEFISCALE = 0
Const TXT_COGNOME = 2
Const TXT_NOME = 4
Const TXT_PARTITAIVA = 1
Const TXT_RAGIONESOCIALE = 3

Const LIST_OPERAZIONE_SMALTIMENTO_DA_ASS = 0
Const LIST_OPERAZIONE_SMALTIMENTO_ASS = 1
Const CMD_LST_ASSOCIA_S_ALL = 0
Const CMD_LST_ASSOCIA_S_ONE = 1
Const CMD_LST_DISASSOCIA_S_ONE = 2
Const CMD_LST_DISASSOCIA_S_ALL = 3

Const LIST_OPERAZIONE_RECUPERO_DA_ASS = 2
Const LIST_OPERAZIONE_RECUPERO_ASS = 3
Const CMD_LST_ASSOCIA_R_ALL = 4
Const CMD_LST_ASSOCIA_R_ONE = 5
Const CMD_LST_DISASSOCIA_R_ONE = 6
Const CMD_LST_DISASSOCIA_R_ALL = 7

Const LIST_RIFIUTI_SMALTIMENTO_DA_ASS = 7
Const LIST_RIFIUTI_SMALTIMENTO_ASS = 6
Const CMD_LST_RIFIUTI_ASSOCIA_S_ALL = 14
Const CMD_LST_RIFIUTI_ASSOCIA_S_ONE = 12
Const CMD_LST_RIFIUTI_DISASSOCIA_S_ALL = 13
Const CMD_LST_RIFIUTI_DISASSOCIA_S_ONE = 15

Const LIST_RIFIUTI_RECUPERO_DA_ASS = 4
Const LIST_RIFIUTI_RECUPERO_ASS = 5
Const CMD_LST_RIFIUTI_ASSOCIA_R_ALL = 9
Const CMD_LST_RIFIUTI_ASSOCIA_R_ONE = 11
Const CMD_LST_RIFIUTI_DISASSOCIA_R_ALL = 10
Const CMD_LST_RIFIUTI_DISASSOCIA_R_ONE = 8

Private colPicture As New Collection

Private rsComuni As Recordset
Private pNoEvent As Boolean

Private rsImpianti As ADODB.Recordset


Private Sub cmbAutorizzazione_Click()

    txtidAutorizzazione.SetFocus
    If cmbAutorizzazione.ListCount = 0 Then Exit Sub
    
    If Val(txtidAutorizzazione.Text) <> cmbAutorizzazione.ItemData(cmbAutorizzazione.ListIndex) Then
        txtidAutorizzazione.Text = cmbAutorizzazione.ItemData(cmbAutorizzazione.ListIndex)
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Dim bCommit  As Boolean
    
    
    On Error GoTo cmd_ClickERR
    Select Case Index
        Case 0
            Unload Me
        Case 2
            ricaricaLista
        Case 3
            If (validaDati) Then
                If pModifiche And Not (TabStrip1.SelectedItem Is Nothing) Then
                    
                    gDBC.BeginTrans
                    bCommit = True
                    If pInsert Then
                        gDBC.Execute "insert into impianti (Cap,CodComune,CodFiscale,CodSoggetto,idAutorizzazione,idSoggetto,Indirizzo, " & _
                                      "NumCivico,NumeroFax,NumeroTelefono,PartitaIva,RagioneSociale,Interno,SubAlterno) values('" & _
                                       txtCap.Text & "','" & txtCodComune.Text & "','" & txtCodFiscale.Text & "','" & txtCodSoggetto.Text & "'," & _
                                       txtidAutorizzazione.Text & "," & txtidSoggetto.Text & ",'" & Replace(txtIndirizzo.Text, "'", "''") & "'," & toZero(txtNumCivico.Text) & ",'" & _
                                       txtNumeroFax.Text & "','" & txtNumeroTelefono.Text & "'," & txtPartitaIva.Text & ",'" & Replace(txtRagioneSociale.Text, "'", "''") & "','" & _
                                       txtInterno.Text & "','" & txtSubalterno.Text & "')"
                        
        
                    Else
                        gDBC.Execute "update impianti set  Cap='" & txtCap.Text & "' ,CodComune='" & txtCodComune.Text & "',CodFiscale='" & txtCodFiscale.Text & "' ,CodSoggetto='" & txtCodSoggetto.Text & "'," & _
                                       "idAutorizzazione=" & txtidAutorizzazione.Text & ",idSoggetto=" & txtidSoggetto.Text & ",Indirizzo='" & Replace(txtIndirizzo.Text, "'", "''") & "'," & _
                                       "NumCivico=" & toZero(txtNumCivico.Text) & ",NumeroFax='" & txtNumeroFax.Text & "',NumeroTelefono='" & txtNumeroTelefono.Text & "'," & _
                                       "PartitaIva=" & txtPartitaIva.Text & ",RagioneSociale='" & Replace(txtRagioneSociale.Text, "'", "''") & "',Interno='" & txtInterno.Text & "' ,SubAlterno='" & txtSubalterno.Text & "' where idImpianto=" & txtImpianto.Text
                    End If
                    
                    gDBC.CommitTrans
                    
                    RefreshGrid
                
                    pInsert = False
                    pModifiche = False
                    
                    If (TabStrip1.SelectedItem.key = "OPERAZIONI") Then
                        salvaAssociazioni
                    End If
                    
                    
                    If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
                    
                    TabStrip1.SelectedItem = TabStrip1.Tabs(1)
                    TabStrip1_Click
                    
                    pErrore = False
                End If
            End If
        Case 4
            pInsert = True
            
            TabStrip1.Tabs("DETTAGLIO").Selected = True
            TabStrip1_Click
            
            txtCap.Text = vbNullString
            txtCodComune.Text = vbNullString
            txtCodFiscale.Text = vbNullString
            txtCodSoggetto.Text = vbNullString
            txtidAutorizzazione.Text = vbNullString
            txtidSoggetto.Text = vbNullString
            txtImpianto.Text = vbNullString
            txtIndirizzo.Text = vbNullString
            txtNumCivico.Text = vbNullString
            txtNumeroFax.Text = vbNullString
            txtNumeroTelefono.Text = vbNullString
            txtPartitaIva.Text = vbNullString
            txtRagioneSociale.Text = vbNullString
            

            
            If pidSoggetto > 0 Then
                txtidSoggetto.Text = pidSoggetto
                txtCodSoggetto.Text = pcodSoggetto
                disabilitaAssociazione False
                
                cmbAutorizzazione_Click
            End If
            
        Case 5
            If (MsgBox("Attenzione, se sicuro di voler cancellare l'impianto: " & rsImpianti("RagioneSociale"), vbQuestion + vbYesNo) = vbYes) Then
                gDBC.BeginTrans
                gDBC.Execute ("delete * from Impianti Where idImpianto =" & rsImpianti!idImpianto)
                gDBC.CommitTrans
                
                txtCap.Text = vbNullString
                txtCodComune.Text = vbNullString
                txtCodFiscale.Text = vbNullString
                txtCodSoggetto.Text = vbNullString
                txtidAutorizzazione.Text = vbNullString
                txtidSoggetto.Text = vbNullString
                txtImpianto.Text = vbNullString
                txtIndirizzo.Text = vbNullString
                txtNumCivico.Text = vbNullString
                txtNumeroFax.Text = vbNullString
                txtNumeroTelefono.Text = vbNullString
                txtPartitaIva.Text = vbNullString
                txtRagioneSociale.Text = vbNullString
                
                
                RefreshGrid
                pModifiche = False
                pInsert = False

                TabStrip1.Tabs("LISTA").Selected = True
                TabStrip1_Click
                
                
            End If
    End Select
    Exit Sub
cmd_ClickERR:
    
    If bCommit Then gDBC.RollbackTrans
    MsgBox Err.Description, vbCritical + vbOKOnly, "frmImpianti"
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
    
    If txt(TXT_CODICEFISCALE) <> vbNullString Then
        search = " CodFiscale like '" & txt(TXT_CODICEFISCALE) & "%'"
    End If
    

    If txt(TXT_COGNOME) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " Cognome like '" & txt(TXT_COGNOME) & "%'"
    End If
    
    If txt(TXT_NOME) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " Nome like '" & txt(TXT_NOME) & "%'"
    End If
    
    
    If txt(TXT_PARTITAIVA) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " PartitaIva = " & txt(TXT_PARTITAIVA)
    End If
    
    If txt(TXT_RAGIONESOCIALE) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " RagioneSociale like '" & txt(TXT_RAGIONESOCIALE) & "%'"
    End If
    
    If pidSoggetto > 0 Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " idSoggetto=" & pidSoggetto & " and codSoggetto='" & pcodSoggetto & "'"
    
        
    End If
    
    rsImpianti.Filter = search
    If search = "" Then
        rsImpianti.Filter = search & " RagioneSociale <> '123'"
    End If
    
    TabStrip1.Tabs("LISTA").Selected = True

    Screen.MousePointer = vbDefault
End Sub


Private Sub ctlComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtCodComune.SetFocus
        txtCodComune.Text = ctlComune.BoundText
        txtCap.Text = findCap(txtCodComune.Text)
    End If
End Sub

Private Sub ctlComune_GotFocus()
    cmd(3).default = False
End Sub

Private Sub ctlComune_LostFocus()
    cmd(3).default = True
End Sub

Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs("DETTAGLIO").Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then ctlComune.Text = nullToBlank(rsImpianti!Comune)
End Sub

Private Sub RefreshGrid()
    Set DataGrid1.DataSource = Nothing
    
    Set rsImpianti = New ADODB.Recordset
    rsImpianti.LockType = adLockBatchOptimistic
    rsImpianti.CursorLocation = adUseClient
    rsImpianti.CursorType = adOpenStatic
    
    rsImpianti.Source = "SELECT Impianti.*, Comuni.Logo , Comuni.Comune FROM Impianti inner join Comuni on Impianti.CodComune = Comuni.CODISTAT"
    Set rsImpianti.ActiveConnection = gDBC
    rsImpianti.Open
    
    Set rsImpianti.ActiveConnection = Nothing
       
    If pidSoggetto > 0 Then
        rsImpianti.Filter = " idSoggetto=" & pidSoggetto & " and codSoggetto='" & pcodSoggetto & "'"
    End If
    
    Set DataGrid1.DataSource = rsImpianti
End Sub


Private Sub Form_Load()
    On Error GoTo LoadERR
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    RefreshGrid
    
    
    Screen.MousePointer = vbHourglass
    pModifiche = False
    DoEvents
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    Picture1(3).Visible = False
    
    Set ctlComune.DataSource = Nothing
    Set ctlComune.RowSource = Nothing
    
    Set rsComuni = CreateObject("ADODB.Recordset")
    rsComuni.LockType = adLockBatchOptimistic
    rsComuni.CursorLocation = adUseClient
    rsComuni.CursorType = adOpenStatic
    
    Set rsComuni.ActiveConnection = gDBC
    
    rsComuni.Source = "Select comune,CODISTAT,cap from comuni order by comune"
    rsComuni.Open
    
    Set rsComuni.ActiveConnection = Nothing
    
    ctlComune.CausesValidation = False
    Set ctlComune.DataSource = rsComuni
    Set ctlComune.RowSource = rsComuni
    
    DataGrid1_RowColChange -1, -1

    Exit Sub
LoadERR:
    MsgBox "Load frmImpianti:" & Err.Description, vbCritical, App.Title
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
    If rsImpianti.state = adStateOpen Then
        rsImpianti.Close
    End If
    
    Set colPicture = Nothing
End Sub

Private Sub Picture1_Resize(Index As Integer)
    Dim widthList As Long

    Select Case Index
        Case 0
            DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
        Case 1
        Case 2
            Frame1(0).Move offsetX, offsetY, Picture1(Index).ScaleWidth - 2 * offsetX
            widthList = (Frame1(0).width - Picture2(0).ScaleWidth - 3 * offsetX) / 2
            
            'smaltimento
            Label2(0).Move offsetX, offsetY * 2
            List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS).Move offsetX, Label2(0).top + Label2(0).Height, widthList
            Picture2(0).Move offsetX + widthList, List1(0).top
            Label2(1).Move widthList + Picture2(0).width + offsetX, offsetY * 2
            List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).Move Label2(1).left, List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS).top, widthList
            'rifiuti per lo smaltimento
            Label2(7).Move offsetX, List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).top + List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).Height + offsetY
            List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).Move offsetX, Label2(7).top + Label2(7).Height, widthList
            Picture2(3).Move offsetX + widthList, List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).top
            Label2(6).Move widthList + Picture2(3).width + offsetX, Label2(7).top
            List1(LIST_RIFIUTI_SMALTIMENTO_ASS).Move Label2(6).left, List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).top, widthList
            
            'recupero
            Frame1(1).Move offsetY, Frame1(0).top + Frame1(0).Height + offsetY, Picture1(Index).ScaleWidth - 2 * offsetX
            
            Label2(2).Move offsetX, offsetY * 2
            List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).Move offsetX, Label2(2).top + Label2(2).Height, widthList
            Picture2(1).Move offsetX + widthList, List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).top
            Label2(3).Move widthList + Picture2(1).width + offsetX, offsetY * 2
            List1(LIST_OPERAZIONE_RECUPERO_ASS).Move Label2(3).left, List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).top, widthList
        
            'rifiuti per lo smaltimento
            Label2(9).Move offsetX, List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).top + List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).Height + offsetY
            List1(LIST_RIFIUTI_RECUPERO_DA_ASS).Move offsetX, Label2(7).top + Label2(7).Height, widthList
            Picture2(2).Move offsetX + widthList, List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).top
            Label2(8).Move widthList + Picture2(2).width + offsetX, Label2(9).top
            List1(LIST_RIFIUTI_RECUPERO_ASS).Move Label2(8).left, List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).top, widthList
            
    End Select
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If pModifiche Then
        If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
            cmd_Click 3
            
            If pErrore Then
                pNoEvent = True
                TabStrip1.Tabs("DETTAGLIO").Selected = True
                pNoEvent = False
            End If
        Else
            pInsert = False
            pModifiche = False
        End If
    End If
    
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    cmd(2).enabled = TabStrip1.SelectedItem.key = "FILTRO"
    cmd(2).default = TabStrip1.SelectedItem.key = "FILTRO"
    
    cmd(3).enabled = (TabStrip1.SelectedItem.key = "DETTAGLIO" Or TabStrip1.SelectedItem.key = "OPERAZIONI") And (pcodSoggetto <> "S" Or (pcodSoggetto = "S" And frmLogin.Livello <> LVL_USER))
    cmd(3).default = (TabStrip1.SelectedItem.key = "DETTAGLIO" Or TabStrip1.SelectedItem.key = "OPERAZIONI") And (pcodSoggetto <> "S" Or (pcodSoggetto = "S" And frmLogin.Livello <> LVL_USER))
    cmd(0).default = TabStrip1.SelectedItem.key = "LISTA"
    cmd(4).enabled = TabStrip1.SelectedItem.key <> "OPERAZIONI" And (pcodSoggetto <> "S" Or (pcodSoggetto = "S" And frmLogin.Livello <> LVL_USER))
    cmd(5).enabled = TabStrip1.SelectedItem.key <> "OPERAZIONI" And (pcodSoggetto <> "S" Or (pcodSoggetto = "S" And frmLogin.Livello <> LVL_USER))
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    colPicture(TabStrip1.SelectedItem.key).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1

    If (TabStrip1.SelectedItem.key = "DETTAGLIO") Then
        caricaAutorizzazioni Val(txtidAutorizzazione.Text)
        
        If Not rsImpianti.EOF Then
            txtCap.Text = rsImpianti!Cap
            txtCodComune.Text = rsImpianti!codComune
            txtCodFiscale.Text = rsImpianti!CodFiscale
            txtCodSoggetto.Text = rsImpianti!codSoggetto
            txtidAutorizzazione.Text = rsImpianti!idAutorizzazione
            txtidSoggetto.Text = rsImpianti!idSoggetto
            txtImpianto.Text = rsImpianti!idImpianto
            txtIndirizzo.Text = nullToBlank(rsImpianti!Indirizzo)
            txtNumCivico.Text = nullToBlank(rsImpianti!NumCivico)
            txtNumeroFax.Text = nullToBlank(rsImpianti!NumeroFax)
            txtNumeroTelefono.Text = nullToBlank(rsImpianti!NumeroTelefono)
            txtPartitaIva.Text = nullToBlank(rsImpianti!PartitaIva)
            txtRagioneSociale.Text = nullToBlank(rsImpianti!RagioneSociale)
        End If
        
        
        txtRagioneSociale.SetFocus
        
        pModifiche = False
    End If

    If (TabStrip1.SelectedItem.key = "OPERAZIONI") Then
        disabilitaAssociazione txtImpianto <> ""
        List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS).SetFocus
        
        If List1(0).Tag <> txtImpianto.Text Then
            caricaListe
            List1(0).Tag = txtImpianto.Text
        End If
        
    End If


    If (TabStrip1.SelectedItem.key = "FILTRO") Then
        txt(0).SetFocus
    End If



End Sub

Public Sub showMe(Optional idSoggetto As Long = 0, Optional codSoggetto As String = "", Optional Des As String = "")
    
    pidSoggetto = idSoggetto
    pcodSoggetto = codSoggetto
    
    colPicture.Add Picture1(0), TabStrip1.Tabs(1).key
    colPicture.Add Picture1(1), TabStrip1.Tabs(2).key
    colPicture.Add Picture1(2), TabStrip1.Tabs(3).key
    colPicture.Add Picture1(3), TabStrip1.Tabs(4).key
    
    If idSoggetto > 0 Then Me.Caption = "Elenco Impianti di " & Des
    
    
'    If (codSoggetto = "S") Then
'        TabStrip1.Tabs.Remove ("OPERAZIONI")
'    End If
    
    TabStrip1.Tabs("LISTA").Selected = True

    Screen.MousePointer = vbDefault
    
    ricaricaLista
    
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
    
    If DevEnvMDB.rscmdOperazioniSRDaAssociare.state = adStateOpen Then
        DevEnvMDB.rscmdOperazioniSRDaAssociare.Close
    End If
    
    If DevEnvMDB.rscmdCodCerSRDaAssociare.state = adStateOpen Then
        DevEnvMDB.rscmdCodCerSRDaAssociare.Close
    End If
    
        
    DevEnvMDB.cmdOperazioniSRDaAssociare txtImpianto, pidSoggetto, pcodSoggetto
    DevEnvMDB.cmdCodCerSRDaAssociare txtImpianto, pidSoggetto, pcodSoggetto
        
    With DevEnvMDB.rscmdOperazioniSRDaAssociare
        Do While Not .EOF
            If .Fields("DaAssociare") = -1 Then
                If .Fields("recupero") = True Then
                    List1(LIST_OPERAZIONE_RECUPERO_DA_ASS).AddItem .Fields("codTipo") & " - " & .Fields("Descrizione")
                Else
                    List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS).AddItem .Fields("codTipo") & " - " & .Fields("Descrizione")
                End If
            Else
                If .Fields("recupero") = True Then
                    List1(LIST_OPERAZIONE_RECUPERO_ASS).AddItem .Fields("codTipo") & " - " & .Fields("Descrizione")
                Else
                    List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).AddItem .Fields("codTipo") & " - " & .Fields("Descrizione")
                End If
            End If
            .MoveNext
        Loop
    End With


    With DevEnvMDB.rscmdCodCerSRDaAssociare
        Do While Not .EOF
            If .Fields("DaRecupero") = -1 Then
                If IsNull(.Fields("idImpianto")) Then
                    List1(LIST_RIFIUTI_RECUPERO_DA_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
                Else
                    List1(LIST_RIFIUTI_RECUPERO_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
                End If
            Else
                If IsNull(.Fields("idImpianto")) Then
                    List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
                Else
                    List1(LIST_RIFIUTI_SMALTIMENTO_ASS).AddItem .Fields("codCategoria") & .Fields("codSottoCategoria") & .Fields("codCer") & " - " & .Fields("Descrizione")
                End If
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
            
        Case CMD_LST_ASSOCIA_S_ALL, CMD_LST_ASSOCIA_S_ONE
            spostaElemento List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS), List1(LIST_OPERAZIONE_SMALTIMENTO_ASS), Index = CMD_LST_ASSOCIA_S_ALL
        Case CMD_LST_DISASSOCIA_S_ONE, CMD_LST_DISASSOCIA_S_ALL
            spostaElemento List1(LIST_OPERAZIONE_SMALTIMENTO_ASS), List1(LIST_OPERAZIONE_SMALTIMENTO_DA_ASS), Index = CMD_LST_DISASSOCIA_S_ALL
        Case CMD_LST_ASSOCIA_R_ALL, CMD_LST_ASSOCIA_R_ONE
            spostaElemento List1(LIST_OPERAZIONE_RECUPERO_DA_ASS), List1(LIST_OPERAZIONE_RECUPERO_ASS), Index = CMD_LST_ASSOCIA_R_ALL
        Case CMD_LST_DISASSOCIA_R_ONE, CMD_LST_DISASSOCIA_R_ALL
            spostaElemento List1(LIST_OPERAZIONE_RECUPERO_ASS), List1(LIST_OPERAZIONE_RECUPERO_DA_ASS), Index = CMD_LST_DISASSOCIA_R_ALL
    
        Case CMD_LST_RIFIUTI_ASSOCIA_S_ALL, CMD_LST_RIFIUTI_ASSOCIA_S_ONE
            spostaElemento List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS), List1(LIST_RIFIUTI_SMALTIMENTO_ASS), Index = CMD_LST_RIFIUTI_ASSOCIA_S_ALL
        Case CMD_LST_RIFIUTI_DISASSOCIA_S_ALL, CMD_LST_RIFIUTI_DISASSOCIA_S_ONE
            spostaElemento List1(LIST_RIFIUTI_SMALTIMENTO_ASS), List1(LIST_RIFIUTI_SMALTIMENTO_DA_ASS), Index = CMD_LST_RIFIUTI_DISASSOCIA_S_ALL
        Case CMD_LST_RIFIUTI_ASSOCIA_R_ALL, CMD_LST_RIFIUTI_ASSOCIA_R_ONE
            spostaElemento List1(LIST_RIFIUTI_RECUPERO_DA_ASS), List1(LIST_RIFIUTI_RECUPERO_ASS), Index = CMD_LST_RIFIUTI_ASSOCIA_R_ALL
        Case CMD_LST_RIFIUTI_DISASSOCIA_R_ALL, CMD_LST_RIFIUTI_DISASSOCIA_R_ONE
            spostaElemento List1(LIST_RIFIUTI_RECUPERO_ASS), List1(LIST_RIFIUTI_RECUPERO_DA_ASS), Index = CMD_LST_RIFIUTI_DISASSOCIA_R_ALL
    
    
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
        Case LIST_OPERAZIONE_SMALTIMENTO_DA_ASS
            Command1_Click (CMD_LST_ASSOCIA_S_ONE)
        Case LIST_OPERAZIONE_SMALTIMENTO_ASS
            Command1_Click (CMD_LST_DISASSOCIA_S_ONE)
        Case LIST_OPERAZIONE_RECUPERO_DA_ASS
            Command1_Click (CMD_LST_ASSOCIA_R_ONE)
        Case LIST_OPERAZIONE_RECUPERO_ASS
            Command1_Click (CMD_LST_DISASSOCIA_R_ONE)
    
        Case LIST_RIFIUTI_SMALTIMENTO_ASS
            Command1_Click (CMD_LST_RIFIUTI_DISASSOCIA_S_ONE)
        Case LIST_RIFIUTI_SMALTIMENTO_DA_ASS
            Command1_Click (CMD_LST_RIFIUTI_ASSOCIA_S_ONE)
        Case LIST_RIFIUTI_RECUPERO_DA_ASS
            Command1_Click (CMD_LST_RIFIUTI_ASSOCIA_R_ONE)
        Case LIST_RIFIUTI_RECUPERO_ASS
            Command1_Click (CMD_LST_RIFIUTI_DISASSOCIA_R_ONE)
    
    End Select
End Sub

Private Sub salvaAssociazioni()
    Dim i As Integer
    Dim qry As ADODB.Command
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codCer As String
    
    gDBC.Execute "delete * from AssImpiantoCodCerPerSmaltimentoRecupero where idImpianto=" & txtImpianto & " and  idSoggetto=" & pidSoggetto & " and  codSoggetto='" & pcodSoggetto & "'"
    gDBC.Execute "delete * from AssImpiantoOperazioniSmaltimentoRecupero where idImpianto=" & txtImpianto & " and  idSoggetto=" & pidSoggetto & " and  codSoggetto='" & pcodSoggetto & "'"
    
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "insert into AssImpiantoOperazioniSmaltimentoRecupero (idImpianto,idSoggetto,codSoggetto,codOperazioneSR) values (?,?,?,?)"
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 1)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 3)
    
    Set qry.ActiveConnection = gDBC
    
    For i = 0 To List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).ListCount - 1
        qry.Parameters(0).Value = txtImpianto
        qry.Parameters(1).Value = pidSoggetto
        qry.Parameters(2).Value = pcodSoggetto
        qry.Parameters(3).Value = left$(List1(LIST_OPERAZIONE_SMALTIMENTO_ASS).List(i), 3)
        qry.Execute
    Next
    
    For i = 0 To List1(LIST_OPERAZIONE_RECUPERO_ASS).ListCount - 1
        qry.Parameters(0).Value = txtImpianto
        qry.Parameters(1).Value = pidSoggetto
        qry.Parameters(2).Value = pcodSoggetto
        qry.Parameters(3).Value = left$(List1(LIST_OPERAZIONE_RECUPERO_ASS).List(i), 3)
        qry.Execute
    Next
    
    Set qry.ActiveConnection = Nothing
    Set qry = Nothing
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "insert into AssImpiantoCodCerPerSmaltimentoRecupero (codCategoria,codSottoCategoria,codCer " _
        & ",idImpianto,idSoggetto,codSoggetto,recupero) values (?,?,?,?,?,?,?)"
    qry.CommandType = adCmdText
    qry.Prepared = True
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 2)
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adVarChar, adParamInput, 1)
    qry.Parameters.Append qry.CreateParameter(, adBoolean, adParamInput)
    
    Set qry.ActiveConnection = gDBC
    
    For i = 0 To List1(LIST_RIFIUTI_SMALTIMENTO_ASS).ListCount - 1
        
        codCer = left$(List1(LIST_RIFIUTI_SMALTIMENTO_ASS).List(i), 6)
        codCategoria = left$(codCer, 2)
        codSottoCategoria = Mid$(codCer, 3, 2)
        codCer = Right$(codCer, 2)
        
        qry.Parameters(0).Value = codCategoria
        qry.Parameters(1).Value = codSottoCategoria
        qry.Parameters(2).Value = codCer
        qry.Parameters(3).Value = txtImpianto
        qry.Parameters(4).Value = pidSoggetto
        qry.Parameters(5).Value = pcodSoggetto
        qry.Parameters(6).Value = 0
        qry.Execute
    Next
    
    For i = 0 To List1(LIST_RIFIUTI_RECUPERO_ASS).ListCount - 1
        
        codCer = left$(List1(LIST_RIFIUTI_RECUPERO_ASS).List(i), 6)
        codCategoria = left$(codCer, 2)
        codSottoCategoria = Mid$(codCer, 3, 2)
        codCer = Right$(codCer, 2)
        
        qry.Parameters(0).Value = codCategoria
        qry.Parameters(1).Value = codSottoCategoria
        qry.Parameters(2).Value = codCer
        qry.Parameters(3).Value = txtImpianto
        qry.Parameters(4).Value = pidSoggetto
        qry.Parameters(5).Value = pcodSoggetto
        qry.Parameters(6).Value = -1
        qry.Execute
    Next
    
End Sub

Public Sub caricaAutorizzazioni(idAutorizzazione As Long)
    cmbAutorizzazione.Clear
    
    If DevEnvMDB.rscmdAutorizzazioni.state = adStateOpen Then
        DevEnvMDB.rscmdAutorizzazioni.Close
    End If
    
    DevEnvMDB.cmdAutorizzazioni pcodSoggetto, pidSoggetto
    DevEnvMDB.rscmdAutorizzazioni.Filter = ""
    DevEnvMDB.rscmdAutorizzazioni.Filter = "DataScadenza > '" & Format$(Now, "yyyy-mm-dd") & "'"
    
    Do While Not DevEnvMDB.rscmdAutorizzazioni.EOF
        cmbAutorizzazione.AddItem DevEnvMDB.rscmdAutorizzazioni!numero & " del " & Format$(DevEnvMDB.rscmdAutorizzazioni!datainiziovalidita, "dd/mm/yyyy") & " scad. " & Format$(DevEnvMDB.rscmdAutorizzazioni!datascadenza, "dd/mm/yyyy")
        cmbAutorizzazione.ItemData(cmbAutorizzazione.ListCount - 1) = DevEnvMDB.rscmdAutorizzazioni!idAutorizzazione
        
        If DevEnvMDB.rscmdAutorizzazioni!idAutorizzazione = idAutorizzazione Then
            cmbAutorizzazione.ListIndex = cmbAutorizzazione.ListCount - 1
        End If
        
        DevEnvMDB.rscmdAutorizzazioni.MoveNext
    Loop
    
    If cmbAutorizzazione.ListCount > 0 And cmbAutorizzazione.ListIndex < 0 Then
        cmbAutorizzazione.ListIndex = 0
    End If
    
    DevEnvMDB.rscmdAutorizzazioni.Filter = ""
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
End Sub


Private Sub txtNumCivico_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If

End Sub

Private Sub txtPartitaIva_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub



Private Function validaDati() As Boolean
    Dim campoInErrore As Boolean
    
    validaDati = False
    campoInErrore = False
    
    txtRagioneSociale.ToolTipText = ""
    txtCodFiscale.ToolTipText = ""
    txtCodComune.ToolTipText = ""
    txtCap.ToolTipText = ""
    txtIndirizzo.ToolTipText = ""
    txtPartitaIva.ToolTipText = ""
    txtidAutorizzazione.ToolTipText = ""

    
    txtRagioneSociale.BackColor = -2147483643
    txtCodFiscale.BackColor = -2147483643
    txtCodComune.BackColor = -2147483643
    txtCap.BackColor = -2147483643
    txtIndirizzo.BackColor = -2147483643
    txtPartitaIva.BackColor = -2147483643
    txtidAutorizzazione.BackColor = &HE0E0E0
    
    
    If txtIndirizzo.Text = "" Then
        txtIndirizzo.ToolTipText = "Attenzione l'Indirizzo è obbligatorio"
        txtIndirizzo.BackColor = vbRed
        txtIndirizzo.SetFocus
        
        campoInErrore = True
    End If
    
    If txtCodComune.Text = "" Then
        txtCodComune.ToolTipText = "Attenzione il Comune è obbligatorio"
        txtCodComune.BackColor = vbRed
        txtCodComune.SetFocus
        
        campoInErrore = True
    End If
    
    If txtCap.Text = "" Then
        txtCap.ToolTipText = "Attenzione il Cap è obbligatorio"
        txtCap.BackColor = vbRed
        txtCap.SetFocus
        
        campoInErrore = True
    End If
    
    If txtPartitaIva.Text = "" And txtCodFiscale.Text = "" Then
        txtPartitaIva.ToolTipText = "Attenzione il (PartitaIva o Codice Fiscale) è obbligatorio"
        txtPartitaIva.BackColor = vbRed
        txtPartitaIva.SetFocus
        
        txtCodFiscale.ToolTipText = "Attenzione il (Ragione Sociale o Codice Fiscale) è obbligatorio"
        txtCodFiscale.BackColor = vbRed
        txtCodFiscale.SetFocus
        
        campoInErrore = True
    End If
    
    If txtRagioneSociale.Text = "" Then
        txtRagioneSociale.ToolTipText = "Attenzione la Ragione Sociale è obbligatoria"
        txtRagioneSociale.BackColor = vbRed
        txtRagioneSociale.SetFocus
        
        campoInErrore = True
    End If
    
    If txtidAutorizzazione.Text = "" Then
        txtidAutorizzazione.ToolTipText = "Attenzione l'Autorizzazione obbligatoria"
        txtidAutorizzazione.BackColor = vbRed
        txtidAutorizzazione.SetFocus
        
        campoInErrore = True
    End If
    
    
    validaDati = Not campoInErrore

End Function

Private Sub txtRagioneSociale_Change()
    pModifiche = True
End Sub

Private Sub txtSubalterno_Change()
    pModifiche = True
End Sub

Private Sub txtCap_Change()
    pModifiche = True
End Sub

Private Sub txtCodComune_Change()
    pModifiche = True
End Sub

Private Sub txtCodFiscale_Change()
    pModifiche = True
End Sub

Private Sub txtidAutorizzazione_Change()
    pModifiche = True
End Sub

Private Sub txtIndirizzo_Change()
    pModifiche = True
End Sub

Private Sub txtInterno_Change()
    pModifiche = True
End Sub

Private Sub txtNumCivico_Change()
    pModifiche = True
End Sub

Private Sub txtNumeroFax_Change()
    pModifiche = True
End Sub

Private Sub txtNumeroTelefono_Change()
    pModifiche = True
End Sub

Private Sub txtPartitaIva_Change()
    pModifiche = True
End Sub
