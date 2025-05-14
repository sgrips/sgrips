VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConducenti 
   Caption         =   "Elenco Conducenti"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13245
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7265
      Index           =   1
      Left            =   2640
      ScaleHeight     =   7230
      ScaleWidth      =   8505
      TabIndex        =   29
      Top             =   240
      Width           =   8535
      Begin VB.TextBox Text2 
         DataField       =   "Interno"
         DataMember      =   "cmdConducenti"
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
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   7
         Top             =   3585
         Width           =   420
      End
      Begin VB.TextBox Text1 
         DataField       =   "Subalterno"
         DataMember      =   "cmdConducenti"
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
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3585
         Width           =   420
      End
      Begin VB.TextBox txtCap 
         DataField       =   "Cap"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   5
         TabIndex        =   2
         Top             =   2160
         Width           =   990
      End
      Begin VB.TextBox txtidTrasportatore 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idTrasportatore"
         DataMember      =   "cmdConducenti"
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
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtNote 
         DataField       =   "Note"
         DataMember      =   "cmdConducenti"
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
         Height          =   2310
         Index           =   1
         Left            =   1890
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4590
         Width           =   4695
      End
      Begin VB.TextBox txtMatricola 
         DataField       =   "Matricola"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   8
         Top             =   4080
         Width           =   2475
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3585
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   4
         Top             =   3090
         Width           =   4695
      End
      Begin VB.TextBox txtCodComune 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2580
         Width           =   990
      End
      Begin VB.TextBox txtNome 
         DataField       =   "Nome"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   1
         Top             =   1710
         Width           =   3375
      End
      Begin VB.TextBox txtCognome 
         DataField       =   "Cognome"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   0
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtidConducente 
         BackColor       =   &H00E0E0E0&
         DataField       =   "idConducente"
         DataMember      =   "cmdConducenti"
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
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   705
         Width           =   660
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmConducenti.frx":0000
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   3000
         TabIndex        =   3
         Top             =   2580
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
         Caption         =   "Cap"
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
         Left            =   1470
         TabIndex        =   42
         Top             =   2175
         Width           =   345
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idTrasportatore:"
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
         Left            =   495
         TabIndex        =   41
         Top             =   270
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Matricola:"
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
         Left            =   960
         TabIndex        =   39
         Top             =   4080
         Width           =   855
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
         Index           =   18
         Left            =   1395
         TabIndex        =   38
         Top             =   4635
         Width           =   465
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nr.Civico/Esp/Int"
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
         Index           =   16
         Left            =   420
         TabIndex        =   37
         Top             =   3615
         Width           =   1440
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
         Index           =   15
         Left            =   1080
         TabIndex        =   36
         Top             =   3135
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
         Index           =   14
         Left            =   1080
         TabIndex        =   34
         Top             =   2625
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   1290
         TabIndex        =   33
         Top             =   1740
         Width           =   570
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cognome:"
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
         Left            =   975
         TabIndex        =   32
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "idConducente:"
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
         Left            =   615
         TabIndex        =   30
         Top             =   735
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
      TabIndex        =   22
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
         TabIndex        =   10
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   12615
      TabIndex        =   20
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmConducenti.frx":0024
         Height          =   2415
         Left            =   240
         TabIndex        =   21
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
         DataMember      =   "cmdConducenti"
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "idConducente"
            Caption         =   "idConducente"
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
            DataField       =   "idTrasportatore"
            Caption         =   "idTrasportatore"
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
            DataField       =   "Cognome"
            Caption         =   "Cognome"
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
            DataField       =   "Nome"
            Caption         =   "Nome"
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
            DataField       =   "Sesso"
            Caption         =   "Sesso"
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
            DataField       =   "CodComune"
            Caption         =   "codComune"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
            DataField       =   "Matricola"
            Caption         =   "Matricola"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3734,929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2489,953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2085,166
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
      Index           =   5
      Left            =   6360
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   15
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
      TabIndex        =   16
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   23
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
Attribute VB_Name = "frmConducenti"
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

Private midTrasportatore As Long
Private mcodSoggetto As String

Private rsComuni As Recordset
Private pNoEvent As Boolean
Private pErrore As Boolean


Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Select Case Index
        Case 0
            cancelCurrentOperation
            Unload Me
        Case 2
            ricaricaLista
        Case 3
                    
            pErrore = DevEnvMDB.rscmdConducenti.EditMode <> adEditNone
                    
            If DevEnvMDB.rscmdConducenti.EditMode <> adEditNone And validaDati Then
                DevEnvMDB.rscmdConducenti.UpdateBatch adAffectAllChapters
                DataGrid1.Refresh
                
                pErrore = False
                If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
            End If
            
        Case 4
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdConducenti.AddNew
            
            txtidTrasportatore = midTrasportatore
        Case 5
            If (MsgBox("Attenzione, se sicuro di voler cancellare il conducente:" & DevEnvMDB.rscmdConducenti("Cognome") & " " & DevEnvMDB.rscmdConducenti("Nome"), vbQuestion + vbYesNo) = vbYes) Then
                DevEnvMDB.rscmdConducenti.Delete
                DevEnvMDB.rscmdConducenti.UpdateBatch adAffectAllChapters
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
    
    
    DevEnvMDB.rscmdConducenti.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdConducenti.Filter = search & " Cognome <> '123'"
    End If
    
    TabStrip1.Tabs(1).Selected = True

    Screen.MousePointer = vbDefault

End Sub


Private Sub ctlComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        DoEvents
        txtCodComune.Text = ctlComune.BoundText
        txtCap.Text = findCap(txtCodComune.Text)
        txtCodComune.SetFocus
    End If
End Sub



Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs(2).Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then ctlComune.Text = nullToBlank(DevEnvMDB.rscmdConducenti.Fields("comune"))
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
    
    sourceElenco = DevEnvMDB.rscmdConducenti.Source
        
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
    Set ctlComune.RowSource = ctlComune.DataSource

    
    DataGrid1_RowColChange -1, -1
    

    Exit Sub
LoadERR:
    MsgBox "Load frmConducenti:" & Err.Description, vbCritical, App.Title
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
    DevEnvMDB.rscmdConducenti.Filter = vbNullString
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
    If (Not DevEnvMDB.rscmdConducenti.EOF And Not DevEnvMDB.rscmdConducenti.BOF) Then
        If DevEnvMDB.rscmdConducenti.EditMode = adEditAdd Then
            DevEnvMDB.rscmdConducenti.CancelBatch
            If (Not DevEnvMDB.rscmdConducenti.EOF And Not DevEnvMDB.rscmdConducenti.BOF) Then
                DevEnvMDB.rscmdConducenti.MoveFirst
            End If
        ElseIf DevEnvMDB.rscmdConducenti.EditMode = adEditInProgress Then
            DevEnvMDB.rscmdConducenti.CancelUpdate
        End If
    End If
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If (Not DevEnvMDB.rscmdConducenti.EOF And Not DevEnvMDB.rscmdConducenti.BOF) Then
        If DevEnvMDB.rscmdConducenti.EditMode <> adEditNone Then
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
    
    cmd(2).enabled = TabStrip1.SelectedItem.Index = 3
    cmd(2).default = TabStrip1.SelectedItem.Index = 3
    cmd(3).enabled = TabStrip1.SelectedItem.Index = 2
    cmd(3).default = TabStrip1.SelectedItem.Index = 2
    cmd(0).default = TabStrip1.SelectedItem.Index = 1
    
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1

    If (TabStrip1.SelectedItem.Index = 2) Then
        txtCognome.SetFocus
    End If

    If (TabStrip1.SelectedItem.Index = 3) Then
        txt(0).SetFocus
    End If



End Sub

Private Sub ctlComune_GotFocus()
    cmd(3).default = False
End Sub

Private Sub ctlComune_LostFocus()
    cmd(3).default = True
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (5)
End Sub


Public Sub showMe(idTrasportatore As Long, Des As String, codSoggetto As String)

    If (DevEnvMDB.rscmdConducenti.state = adStateOpen) Then DevEnvMDB.rscmdConducenti.Close
    DevEnvMDB.cmdConducenti idTrasportatore
    midTrasportatore = idTrasportatore
    mcodSoggetto = codSoggetto
    Label1(0).Caption = Des
    
    Me.Show vbModal
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



Private Function validaDati() As Boolean
    Dim campoInErrore As Boolean
    
    validaDati = False
    campoInErrore = False
    
    txtCognome.ToolTipText = ""
    txtCognome.BackColor = -2147483643
    txtNome.ToolTipText = ""
    txtNome.BackColor = -2147483643
    
    If txtCognome.Text = "" Then
        txtCognome.ToolTipText = "Attenzione il Cognome è obbligatorio"
        txtCognome.BackColor = vbRed
        txtCognome.SetFocus
        
        campoInErrore = True
    End If
    
    If txtNome.Text = "" Then
        txtNome.ToolTipText = "Attenzione il Nome è obbligatorio"
        txtNome.BackColor = vbRed
        txtNome.SetFocus
        
        campoInErrore = True
    End If
    
    validaDati = Not campoInErrore

End Function
