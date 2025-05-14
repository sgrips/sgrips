VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmIntermediari 
   Caption         =   "Elenco Intermediari"
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
      Height          =   6975
      Index           =   1
      Left            =   3720
      ScaleHeight     =   6945
      ScaleWidth      =   7665
      TabIndex        =   31
      Top             =   240
      Width           =   7695
      Begin VB.TextBox Text1 
         DataField       =   "Subalterno"
         DataMember      =   "cmdIntermediari"
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
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   7
         Top             =   3765
         Width           =   420
      End
      Begin VB.TextBox Text2 
         DataField       =   "Interno"
         DataMember      =   "cmdIntermediari"
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
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3765
         Width           =   420
      End
      Begin VB.TextBox txtCap 
         DataField       =   "Cap"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2235
         MaxLength       =   5
         TabIndex        =   3
         Top             =   2220
         Width           =   1005
      End
      Begin VB.TextBox txtCodComuneSedeLegale 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2700
         Width           =   990
      End
      Begin VB.TextBox txtNote 
         DataField       =   "Note"
         DataMember      =   "cmdIntermediari"
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
         Height          =   1725
         Left            =   2220
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   5355
         Width           =   5055
      End
      Begin VB.TextBox txtNumeroFax 
         DataField       =   "NumeroFax"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   15
         TabIndex        =   10
         Top             =   4770
         Width           =   2475
      End
      Begin VB.TextBox txtNumeroTelefono 
         DataField       =   "NumeroTelefono"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   15
         TabIndex        =   9
         Top             =   4260
         Width           =   2475
      End
      Begin VB.TextBox txtNumCivicoSedeLegale 
         DataField       =   "NumCivico"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3765
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzoSedeLegale 
         DataField       =   "Indirizzo"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   100
         TabIndex        =   5
         Top             =   3270
         Width           =   5175
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
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   11
         TabIndex        =   2
         Top             =   1785
         Width           =   2685
      End
      Begin VB.TextBox txtRagioneSociale 
         DataField       =   "RagioneSociale"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   0
         Top             =   810
         Width           =   3375
      End
      Begin VB.TextBox txtCodFiscale 
         DataField       =   "CodFiscale"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1260
         Width           =   2640
      End
      Begin VB.TextBox txtidTrasportatore 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idIntermediario"
         DataMember      =   "cmdIntermediari"
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
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   278
         Width           =   660
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmIntermediario.frx":0000
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   3360
         TabIndex        =   4
         Top             =   2700
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
         Index           =   4
         Left            =   1800
         TabIndex        =   45
         Top             =   2235
         Width           =   405
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
         Index           =   10
         Left            =   1410
         TabIndex        =   41
         Top             =   2745
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Note:"
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
         Left            =   1725
         TabIndex        =   40
         Top             =   5400
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Fax:"
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
         Left            =   1740
         TabIndex        =   39
         Top             =   4785
         Width           =   450
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
         Left            =   1410
         TabIndex        =   38
         Top             =   4275
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Num Civico:"
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
         Left            =   1155
         TabIndex        =   37
         Top             =   3780
         Width           =   1035
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
         Left            =   1410
         TabIndex        =   36
         Top             =   3285
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
         Left            =   1275
         TabIndex        =   35
         Top             =   1800
         Width           =   915
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
         Index           =   2
         Left            =   720
         TabIndex        =   34
         Top             =   825
         Width           =   1470
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Codice Fiscale:"
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
         Left            =   840
         TabIndex        =   33
         Top             =   1275
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Trasportatore:"
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
         Left            =   990
         TabIndex        =   32
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   2160
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   24
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
         TabIndex        =   18
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
         TabIndex        =   44
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
         TabIndex        =   19
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
      Index           =   1
      Left            =   9960
      TabIndex        =   43
      Top             =   8040
      Width           =   1575
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
      Left            =   2160
      TabIndex        =   17
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
      Left            =   3480
      TabIndex        =   16
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
      Left            =   4800
      TabIndex        =   15
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
      TabIndex        =   13
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
      TabIndex        =   21
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmIntermediario.frx":0024
         Height          =   2415
         Left            =   240
         TabIndex        =   23
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
         DataMember      =   "cmdIntermediari"
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "idIntermediario"
            Caption         =   "Codice"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "PartitaIva"
            Caption         =   "Partita Iva"
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
            DataField       =   "CodComune"
            Caption         =   "idComune"
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
            Caption         =   "Num Civico"
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
            Caption         =   "Numero Telefono"
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
            Caption         =   "Numero Fax"
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
            DataField       =   "Note"
            Caption         =   "Note"
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
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2970,142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2174,74
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2624,882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2085,166
            EndProperty
         EndProperty
      End
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
      Left            =   6120
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   25
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
Attribute VB_Name = "frmIntermediari"
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
Const CMD_AUTORIZZAZIONI = 1
Const CMD_RICERCA = 2
Const CMD_SALVA = 3
Const CMD_NUOVO = 4
Const CMD_CANCELLA = 5

Private rsComuni As Recordset
Private pNoEvent  As Boolean
Private pErrore As Boolean


Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Select Case Index
        Case CMD_CHIUDI
            cancelCurrentOperation
            Unload Me
        Case CMD_RICERCA
            ricaricaLista
        Case CMD_SALVA
        
            pErrore = DevEnvMDB.rscmdIntermediari.EditMode <> adEditNone
            If DevEnvMDB.rscmdIntermediari.EditMode <> adEditNone Then
                DevEnvMDB.rscmdIntermediari.UpdateBatch adAffectAllChapters
                DataGrid1.Refresh
                
                cmd(CMD_AUTORIZZAZIONI).enabled = True
                If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
                pErrore = False
            End If
        Case CMD_NUOVO
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdIntermediari.AddNew
            
            cmd(CMD_AUTORIZZAZIONI).enabled = False
            
        Case CMD_CANCELLA
            If (MsgBox("Attenzione, se sicuro di voler cancellare l'intermediario:" & DevEnvMDB.rscmdIntermediari("RagioneSociale"), vbQuestion + vbYesNo) = vbYes) Then
                'cancella autorizzazioni
                gDBC.Execute "delete * from Autorizzazioni where idsoggetto=" & DevEnvMDB.rscmdIntermediari("idIntermediario") & " and codSoggetto='I'"

                
                DevEnvMDB.rscmdIntermediari.Delete
                DevEnvMDB.rscmdIntermediari.UpdateBatch adAffectAllChapters
            End If
        Case CMD_AUTORIZZAZIONI
            If DataGrid1.Row <> -1 Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                frmAutorizzazioni.showMe DataGrid1.Columns("codice").Value, "I", Des
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
    
    
    DevEnvMDB.rscmdIntermediari.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdIntermediari.Filter = search & " RagioneSociale <> '123'"
    End If
    
    TabStrip1.Tabs(1).Selected = True

    Screen.MousePointer = vbDefault

End Sub


Private Sub ctlComune_Click(Area As Integer)
    If Area = 2 Then
        On Error Resume Next
        txtCodComuneSedeLegale.SetFocus
        txtCodComuneSedeLegale.Text = ctlComune.BoundText
        txtCap.Text = findCap(txtCodComuneSedeLegale.Text)
    End If
End Sub

Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs(2).Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then ctlComune.Text = nullToBlank(DevEnvMDB.rscmdIntermediari.Fields("comune"))
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
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Screen.MousePointer = vbDefault
    
    sourceElenco = DevEnvMDB.rscmdIntermediari.Source
    
    
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
    MsgBox "Load FrmIntermediari:" & Err.Description, vbCritical, App.Title
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
    If DevEnvMDB.rscmdIntermediari.state = adStateOpen Then
        DevEnvMDB.rscmdIntermediari.Close
    End If
    
    DevEnvMDB.rscmdIntermediari.Filter = vbNullString
    
    Set pfrmIntermediari = Nothing
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
    On Error Resume Next
    If DevEnvMDB.rscmdIntermediari.EditMode = adEditAdd Then
        DevEnvMDB.rscmdIntermediari.CancelBatch
                If (Not DevEnvMDB.rscmdIntermediari.EOF And Not DevEnvMDB.rscmdIntermediari.BOF) Then
            DevEnvMDB.rscmdIntermediari.MoveFirst
        End If
    ElseIf DevEnvMDB.rscmdIntermediari.EditMode = adEditInProgress Then
        DevEnvMDB.rscmdIntermediari.CancelUpdate
    End If
End Sub


Private Sub TabStrip1_Click()
    If pNoEvent Then Exit Sub

    If Not DevEnvMDB.rscmdIntermediari.EOF And Not DevEnvMDB.rscmdIntermediari.BOF Then
        If (DevEnvMDB.rscmdIntermediari.EditMode <> adEditNone) Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click CMD_SALVA
                
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
    End If

    If (TabStrip1.SelectedItem.Index = 3) Then
        txt(0).SetFocus
        cmd(CMD_AUTORIZZAZIONI).enabled = False
    End If



End Sub

Private Sub ctlComune_GotFocus()
    cmd(CMD_SALVA).default = False
End Sub

Private Sub ctlComune_LostFocus()
    cmd(CMD_SALVA).default = True
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (CMD_CANCELLA)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
End Sub

Private Sub txtNumCivicoSedeLegale_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPartitaIva_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub
