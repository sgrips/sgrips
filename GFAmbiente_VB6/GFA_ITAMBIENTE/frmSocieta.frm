VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSocieta 
   Caption         =   "Elenco Societa"
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
      Height          =   7265
      Index           =   1
      Left            =   1440
      ScaleHeight     =   7230
      ScaleWidth      =   8505
      TabIndex        =   35
      Top             =   360
      Width           =   8535
      Begin VB.TextBox Text3 
         DataField       =   "CapSociale"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   49
         Top             =   2280
         Width           =   3300
      End
      Begin VB.TextBox Text6 
         DataField       =   "Email"
         DataMember      =   "cmdSocieta"
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
         MaxLength       =   100
         TabIndex        =   7
         Top             =   4080
         Width           =   4740
      End
      Begin VB.TextBox Text4 
         DataField       =   "Logo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdSocieta"
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
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2760
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         DataField       =   "Subalterno"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   10
         Top             =   5010
         Width           =   420
      End
      Begin VB.TextBox Text2 
         DataField       =   "Interno"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   11
         Top             =   5010
         Width           =   420
      End
      Begin VB.TextBox txtCap 
         DataField       =   "Cap"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   5
         Top             =   3225
         Width           =   1020
      End
      Begin VB.TextBox txtNumeroFax 
         DataField       =   "NumeroFax"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   13
         Top             =   6015
         Width           =   3195
      End
      Begin VB.TextBox txtNumeroTelefono 
         DataField       =   "NumeroTelefono"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   12
         Top             =   5505
         Width           =   3195
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   9
         Top             =   5010
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
         DataMember      =   "cmdSocieta"
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
         MaxLength       =   100
         TabIndex        =   8
         Top             =   4515
         Width           =   4800
      End
      Begin VB.TextBox txtCodComune 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         DataMember      =   "cmdSocieta"
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
         TabIndex        =   41
         TabStop         =   0   'False
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
         DataMember      =   "cmdSocieta"
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
         Width           =   3300
      End
      Begin VB.TextBox txtCodFiscale 
         DataField       =   "CodFiscale"
         DataMember      =   "cmdSocieta"
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
         Width           =   3360
      End
      Begin VB.TextBox txtRagioneSociale 
         DataField       =   "RagioneSociale"
         DataMember      =   "cmdSocieta"
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
      Begin VB.TextBox txtidSocieta 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         DataMember      =   "cmdSocieta"
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
         Top             =   330
         Width           =   660
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmSocieta.frx":0000
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   2880
         TabIndex        =   6
         Top             =   3645
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
         Caption         =   "Cap.Sociale:"
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
         Left            =   525
         TabIndex        =   50
         Top             =   2325
         Width           =   1110
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email:"
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
         Index           =   13
         Left            =   1050
         TabIndex        =   48
         Top             =   4125
         Width           =   555
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Logo:"
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
         Left            =   1155
         TabIndex        =   47
         Top             =   2805
         Width           =   480
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
         Left            =   1245
         TabIndex        =   46
         Top             =   3270
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
         TabIndex        =   45
         Top             =   6060
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
         TabIndex        =   44
         Top             =   5550
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
         TabIndex        =   43
         Top             =   5055
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
         TabIndex        =   42
         Top             =   4560
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   870
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idSocieta:"
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
         Left            =   795
         TabIndex        =   36
         Top             =   375
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   2160
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   28
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
         TabIndex        =   14
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   26
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmSocieta.frx":0024
         Height          =   2415
         Left            =   240
         TabIndex        =   27
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
         DataMember      =   "cmdSocieta"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "idSocieta"
            Caption         =   "idSocieta"
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
      Left            =   3720
      TabIndex        =   25
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
      Left            =   5040
      TabIndex        =   24
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
      Left            =   6360
      TabIndex        =   23
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
      TabIndex        =   19
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
      Left            =   7680
      TabIndex        =   22
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   29
      Top             =   3720
      Width           =   3015
      _ExtentX        =   5318
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
Attribute VB_Name = "frmSocieta"
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

Private pErrore As Boolean
Private pNoEvent As Boolean

Private rsComuni As Recordset

Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Select Case Index
        Case CMD_CHIUDI
            cancelCurrentOperation
            Unload Me
        Case CMD_IMPIANTI
            If DataGrid1.Row <> -1 Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                frmImpianti.showMe DataGrid1.Columns("idSocieta").Value, "S", Des
            End If
        Case CMD_RICERCA
            ricaricaLista
        Case CMD_SALVA
            On Error Resume Next
            pErrore = DevEnvMDB.rscmdSocieta.EditMode <> adEditNone
        
            If DevEnvMDB.rscmdSocieta.EditMode <> adEditNone Then
               
               DevEnvMDB.rscmdSocieta.UpdateBatch adAffectAllChapters
               
               DataGrid1.Refresh
               
               If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
               
               pErrore = False
            End If
            
            cmd(CMD_AUTORIZZAZIONI).enabled = True
            cmd(CMD_IMPIANTI).enabled = True
        Case CMD_NUOVO
            
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdSocieta.AddNew
            
            cmd(CMD_AUTORIZZAZIONI).enabled = False
            cmd(CMD_IMPIANTI).enabled = False
        Case CMD_CANCELLA
            If (MsgBox("Attenzione, se sicuro di voler cancellare la società: " & DevEnvMDB.rscmdSocieta("RagioneSociale"), vbQuestion + vbYesNo) = vbYes) Then
                
                'cancella autorizzazioni
                gDBC.Execute "delete * from Autorizzazioni where idsoggetto=" & DevEnvMDB.rscmdSocieta("idSocieta") & " and codSoggetto='S'"
                'cancella impianti
                gDBC.Execute "delete * from Impianti where idsoggetto=" & DevEnvMDB.rscmdSocieta("idSocieta") & " and codSoggetto='S'"
                
                'cancella società
                DevEnvMDB.rscmdSocieta.Delete
                DevEnvMDB.rscmdSocieta.UpdateBatch adAffectAllChapters
            End If
        Case CMD_AUTORIZZAZIONI
            If DataGrid1.Row <> -1 Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                frmAutorizzazioni.showMe DataGrid1.Columns("idSocieta").Value, "S", Des
            End If
        
    End Select
End Sub


Private Function formatValue(Item As Column) As Variant
    On Error Resume Next
    Dim v As Variant
    
    v = Item.Value
    If err.Number > 0 Then v = Item.Text
    
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
    
    
    DevEnvMDB.rscmdSocieta.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdSocieta.Filter = search & " RagioneSociale <> '123'"
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
    If DataGrid1.Row >= 0 Then ctlComune.Text = nullToBlank(DevEnvMDB.rscmdSocieta.Fields("comune"))
End Sub

Private Sub Form_Load()
    On Error GoTo LoadERR
    
    If DevEnvMDB.rscmdSocieta.state = adStateOpen Then
        DevEnvMDB.rscmdSocieta.Close
        DevEnvMDB.cmdSocieta
    Else
        DevEnvMDB.cmdSocieta
    End If
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    
    sourceElenco = DevEnvMDB.rscmdSocieta.Source
        
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

    
    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            Dim t As TextBox
            Set t = c
            If t.DataMember = "cmdSocieta" Then
                Set t.DataSource = Nothing
                Set t.DataSource = DevEnvMDB
            End If
        End If
    Next
    
    Set DataGrid1.DataSource = DevEnvMDB
    
    DataGrid1_RowColChange -1, -1
    Screen.MousePointer = vbDefault

    Exit Sub
LoadERR:
    MsgBox "Load frmSocieta:" & err.Description, vbCritical, App.Title
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
    If DevEnvMDB.rscmdSocieta.state = adStateOpen Then
        DevEnvMDB.rscmdSocieta.Close
    End If
    DevEnvMDB.rscmdSocieta.Filter = vbNullString
    
    Set pfrmSocieta = Nothing
End Sub

Private Sub Picture1_Resize(Index As Integer)
    Select Case Index
        Case 0
            DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
        Case 1
        Case 2
            
    End Select
End Sub

Private Sub cancelCurrentOperation()
    If DevEnvMDB.rscmdSocieta.EOF Or DevEnvMDB.rscmdSocieta.BOF Then Exit Sub

    On Error Resume Next
    If DevEnvMDB.rscmdSocieta.EditMode = adEditAdd Then
        DevEnvMDB.rscmdSocieta.CancelBatch
        If (Not DevEnvMDB.rscmdSocieta.EOF And Not DevEnvMDB.rscmdSocieta.BOF) Then
            DevEnvMDB.rscmdSocieta.MoveFirst
        End If
    ElseIf DevEnvMDB.rscmdSocieta.EditMode = adEditInProgress Then
        DevEnvMDB.rscmdSocieta.CancelUpdate
    End If
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If (Not DevEnvMDB.rscmdSocieta.EOF And Not DevEnvMDB.rscmdSocieta.BOF) Then
        If DevEnvMDB.rscmdSocieta.EditMode <> adEditNone Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click 3
                
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
    cmd(CMD_SALVA).enabled = TabStrip1.SelectedItem.Index = 2 And frmLogin.Livello <> LVL_USER
    cmd(CMD_SALVA).default = TabStrip1.SelectedItem.Index = 2 And frmLogin.Livello <> LVL_USER
    cmd(CMD_CHIUDI).default = TabStrip1.SelectedItem.Index = 1
    cmd(CMD_NUOVO).enabled = frmLogin.Livello <> LVL_USER
    cmd(CMD_CANCELLA).enabled = frmLogin.Livello <> LVL_USER
    
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


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
End Sub



