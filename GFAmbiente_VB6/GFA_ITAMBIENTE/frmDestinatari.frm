VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDestinatari 
   Caption         =   "Elenco Destinatari"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   13245
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6060
      Index           =   1
      Left            =   1440
      ScaleHeight     =   6030
      ScaleWidth      =   11265
      TabIndex        =   33
      Top             =   360
      Width           =   11295
      Begin VB.TextBox Text1 
         DataField       =   "Subalterno"
         DataMember      =   "cmdDestinatari"
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
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3810
         Width           =   420
      End
      Begin VB.TextBox Text2 
         DataField       =   "Interno"
         DataMember      =   "cmdDestinatari"
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
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   9
         Top             =   3810
         Width           =   420
      End
      Begin VB.TextBox txtCap 
         DataField       =   "Cap"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   5
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2280
         Width           =   990
      End
      Begin VB.TextBox txtNumeroFax 
         DataField       =   "NumeroFax"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   15
         TabIndex        =   11
         Top             =   4815
         Width           =   2475
      End
      Begin VB.TextBox txtNumeroTelefono 
         DataField       =   "NumeroTelefono"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   15
         TabIndex        =   10
         Top             =   4305
         Width           =   2475
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3810
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3315
         Width           =   4800
      End
      Begin VB.TextBox txtCodComune 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         DataMember      =   "cmdDestinatari"
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
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2805
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
         DataMember      =   "cmdDestinatari"
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
         TabIndex        =   3
         Top             =   1830
         Width           =   2580
      End
      Begin VB.TextBox txtCodFiscale 
         DataField       =   "CodFiscale"
         DataMember      =   "cmdDestinatari"
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
         TabIndex        =   2
         Top             =   1335
         Width           =   2640
      End
      Begin VB.TextBox txtRagioneSociale 
         DataField       =   "RagioneSociale"
         DataMember      =   "cmdDestinatari"
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   825
         Width           =   3375
      End
      Begin VB.TextBox txtidDestinatario 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idDestinatario"
         DataMember      =   "cmdDestinatari"
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
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   660
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmDestinatari.frx":0000
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         Top             =   2805
         Width           =   3615
         _ExtentX        =   6376
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
         Index           =   9
         Left            =   1215
         TabIndex        =   44
         Top             =   2325
         Width           =   405
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
         TabIndex        =   43
         Top             =   4860
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
         TabIndex        =   42
         Top             =   4350
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
         TabIndex        =   41
         Top             =   3855
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
         TabIndex        =   40
         Top             =   3360
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
         TabIndex        =   38
         Top             =   2850
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
         TabIndex        =   37
         Top             =   1875
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
         TabIndex        =   36
         Top             =   1380
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
         TabIndex        =   35
         Top             =   870
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idDestinatario:"
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
         Left            =   420
         TabIndex        =   34
         Top             =   375
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   2160
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   26
      Top             =   4440
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
         TabIndex        =   12
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Autorizzazioni"
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
      Left            =   8640
      TabIndex        =   19
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Impianti"
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
      TabIndex        =   18
      Top             =   8040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   12615
      TabIndex        =   24
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmDestinatari.frx":0024
         Height          =   2415
         Left            =   240
         TabIndex        =   25
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
         DataMember      =   "cmdDestinatari"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "idDestinatario"
            Caption         =   "idDestinatario"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "comune"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
               ColumnWidth     =   2954,835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2910,047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2415,118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1409,953
            EndProperty
         EndProperty
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
      Index           =   6
      Left            =   1800
      TabIndex        =   23
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
      Index           =   5
      Left            =   3120
      TabIndex        =   22
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
      Index           =   4
      Left            =   4440
      TabIndex        =   21
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
      TabIndex        =   17
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
      Index           =   3
      Left            =   5760
      TabIndex        =   20
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Elenco dei trasportatori"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dettaglio"
            Object.ToolTipText     =   "Modifica dettaglio trasportatore"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro"
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
Attribute VB_Name = "frmDestinatari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private sourceElenco As String

Const TXT_CODICEFISCALE = 0
Const TXT_COGNOME = 2
Const TXT_NOME = 4
Const TXT_PARTITAIVA = 1
Const TXT_RAGIONESOCIALE = 3


Const CMD_CHIUDI = 0
Const CMD_IMPIANTI = 1
Const CMD_AUTORIZZAZIONI = 2
Const CMD_RICERCA = 3
Const CMD_SALVA = 4
Const CMD_NUOVO = 5
Const CMD_CANCELLA = 6

Private rsComuni As Recordset
Private pErrore As Boolean
Private pNoEvent As Boolean

Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Select Case Index
        Case CMD_CHIUDI
            cancelCurrentOperation
            Unload Me
        Case CMD_IMPIANTI
            If DataGrid1.Row <> -1 Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                frmImpianti.showMe DataGrid1.Columns("idDestinatario").Value, "D", Des
            End If
        Case CMD_RICERCA
            ricaricaLista
        Case CMD_SALVA
            
            pErrore = DevEnvMDB.rscmdDestinatari.EditMode <> adEditNone
                        
            If DevEnvMDB.rscmdDestinatari.EditMode <> adEditNone And validaDati Then
               DevEnvMDB.rscmdDestinatari.UpdateBatch adAffectAllChapters
               DataGrid1.Refresh
               If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
               
               pErrore = False
            End If
            
            If Not pErrore Then
                cmd(CMD_AUTORIZZAZIONI).enabled = True
                cmd(CMD_IMPIANTI).enabled = True
            End If
        Case CMD_NUOVO
            
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdDestinatari.AddNew
            
            cmd(CMD_AUTORIZZAZIONI).enabled = False
            cmd(CMD_IMPIANTI).enabled = False
            
        Case CMD_CANCELLA
            If (MsgBox("Attenzione, se sicuro di voler cancellare il Destinatario: " & DevEnvMDB.rscmdDestinatari("RagioneSociale"), vbQuestion + vbYesNo) = vbYes) Then
                'cancella autorizzazioni
                gDBC.Execute "delete * from Autorizzazioni where idsoggetto=" & DevEnvMDB.rscmdDestinatari("idDestinatario") & " and codSoggetto='D'"
                'cancella impianti
                gDBC.Execute "delete * from Impianti where idsoggetto=" & DevEnvMDB.rscmdDestinatari("idDestinatario") & " and codSoggetto='D'"
                                
                'destinatari
                DevEnvMDB.rscmdDestinatari.Delete
                DevEnvMDB.rscmdDestinatari.UpdateBatch adAffectAllChapters
            End If
        Case CMD_AUTORIZZAZIONI
            If DataGrid1.Row <> -1 Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                frmAutorizzazioni.showMe DataGrid1.Columns("idDestinatario").Value, "D", Des
            End If
        
    End Select
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
    
    
    DevEnvMDB.rscmdDestinatari.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdDestinatari.Filter = search & " RagioneSociale <> '123'"
    End If
    
    
    
    TabStrip1.Tabs(1).Selected = True

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

Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs(2).Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (5)
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then ctlComune.Text = nullToBlank(DevEnvMDB.rscmdDestinatari.Fields("comune"))
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    On Error GoTo LoadERR
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    For i = Picture1.LBound To Picture1.ubound
        Picture1(i).Visible = False
    Next
    
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Screen.MousePointer = vbDefault
    
    sourceElenco = DevEnvMDB.rscmdDestinatari.Source
        
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
    MsgBox "FrmDestinatari:" & Err.Description, vbCritical, App.Title
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
    If DevEnvMDB.rscmdDestinatari.state = adStateOpen Then
        DevEnvMDB.rscmdDestinatari.Close
    End If
    DevEnvMDB.rscmdDestinatari.Filter = vbNullString
    
    Set pfrmDestinatari = Nothing
End Sub


Private Sub Picture1_Resize(Index As Integer)
    Dim widthList  As Long
    
    Select Case Index
        Case 0
            DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
            
    End Select
End Sub

Private Sub cancelCurrentOperation()
    On Error Resume Next
    If DevEnvMDB.rscmdDestinatari.EditMode = adEditAdd Then
        DevEnvMDB.rscmdDestinatari.CancelBatch
        If (Not DevEnvMDB.rscmdDestinatari.EOF And Not DevEnvMDB.rscmdDestinatari.BOF) Then
            DevEnvMDB.rscmdDestinatari.MoveFirst
        End If
    ElseIf DevEnvMDB.rscmdDestinatari.EditMode = adEditInProgress Then
        DevEnvMDB.rscmdDestinatari.CancelUpdate
    End If
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If (Not DevEnvMDB.rscmdDestinatari.EOF And Not DevEnvMDB.rscmdDestinatari.BOF) Then
        If DevEnvMDB.rscmdDestinatari.EditMode <> adEditNone Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click 4
                If pErrore Then
                    pNoEvent = True
                    TabStrip1.Tabs(2).Selected = True
                    pNoEvent = False
                End If
            Else
                cancelCurrentOperation
            End If
        End If
    End If
    
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    cmd(CMD_RICERCA).enabled = TabStrip1.SelectedItem.Index = 3
    cmd(CMD_RICERCA).default = TabStrip1.SelectedItem.Index = 3
    cmd(CMD_SALVA).enabled = TabStrip1.SelectedItem.Index = 2
    cmd(CMD_SALVA).default = TabStrip1.SelectedItem.Index = 2
    cmd(CMD_CHIUDI).default = TabStrip1.SelectedItem.Index = 1
    
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1

    If (TabStrip1.SelectedItem.Index = 2) Then
        txtRagioneSociale.SetFocus
        
        cmd(CMD_AUTORIZZAZIONI).enabled = True
        cmd(CMD_IMPIANTI).enabled = True
    End If

    If (TabStrip1.SelectedItem.Index = 3) Then
        txt(0).SetFocus
        
        cmd(CMD_AUTORIZZAZIONI).enabled = False
        cmd(CMD_IMPIANTI).enabled = False
        
    End If



End Sub

Private Sub ctlComune_GotFocus()
    cmd(CMD_SALVA).default = False
End Sub

Private Sub ctlComune_LostFocus()
    cmd(CMD_SALVA).default = True
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

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
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
    
    
    txtRagioneSociale.BackColor = -2147483643
    txtCodFiscale.BackColor = -2147483643
    txtCodComune.BackColor = -2147483643
    txtCap.BackColor = -2147483643
    txtIndirizzo.BackColor = -2147483643
    txtPartitaIva.BackColor = -2147483643
    
    
    
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
    
    validaDati = Not campoInErrore

    
End Function

