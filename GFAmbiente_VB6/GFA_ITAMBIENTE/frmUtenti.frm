VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUtenti 
   Caption         =   "Elenco Utenti"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   13245
   Begin VB.CommandButton cmd 
      Caption         =   "Immobili"
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
      Left            =   4800
      TabIndex        =   25
      Top             =   8040
      Width           =   1215
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
      Index           =   3
      Left            =   7440
      TabIndex        =   10
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Badge"
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
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Automezzi"
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
      TabIndex        =   8
      Top             =   8040
      Width           =   1215
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
      Left            =   11640
      TabIndex        =   7
      Top             =   8040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   1
      Left            =   3480
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   14
      Top             =   3600
      Width           =   8895
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   6
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Badge"
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   22
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Utente"
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Fiscale"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Partita Iva"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cognome"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ragione Sociale"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   17
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Index           =   4
         Left            =   5640
         TabIndex        =   16
         Top             =   840
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
      TabIndex        =   12
      Top             =   600
      Width           =   12615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmUtenti.frx":0000
         Height          =   2415
         Left            =   225
         TabIndex        =   13
         Top             =   300
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         DataMember      =   "cmdUtenti"
         ColumnCount     =   39
         BeginProperty Column00 
            DataField       =   "CodUtente"
            Caption         =   "Utente"
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
            DataField       =   "CodAnagrafe"
            Caption         =   "CodAnagrafe"
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
            DataField       =   "CodBadge"
            Caption         =   "Badge"
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
            DataField       =   "DataBadgeInizio"
            Caption         =   "Data Badge Inizio"
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
            DataField       =   "DataBadgeFine"
            Caption         =   "Data Badge Fine"
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
            DataField       =   "ProgressivoBadge"
            Caption         =   "Progressivo Badge"
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
            DataField       =   "TipoAnagrafe"
            Caption         =   "Tipo"
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
            DataField       =   "EnteCreditore"
            Caption         =   "Ente Creditore"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
            DataField       =   "DataDiNascita"
            Caption         =   "Data Di Nascita"
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
            DataField       =   "CodComuneDiNascita"
            Caption         =   "CodComuneDiNascita"
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
            DataField       =   "ComuneDinascita.Comune"
            Caption         =   "ComuneDinascita.Comune"
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
         BeginProperty Column16 
            DataField       =   "CodFiscale"
            Caption         =   "CodFiscale"
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "CodiceISTAT"
            Caption         =   "CodiceIstat"
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
         BeginProperty Column19 
            DataField       =   "CodComuneResidenza"
            Caption         =   "CodComuneResidenza"
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
         BeginProperty Column20 
            DataField       =   "ComuneDiResidenza.Comune"
            Caption         =   "Comune Di Residenza"
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
         BeginProperty Column21 
            DataField       =   "CodViaResidenza"
            Caption         =   "CodViaResidenza"
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
         BeginProperty Column22 
            DataField       =   "IndirizzoResidenza"
            Caption         =   "Indirizzo Residenza"
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
         BeginProperty Column23 
            DataField       =   "NumCivicoResidenza"
            Caption         =   "Num Civico Residenza"
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
         BeginProperty Column24 
            DataField       =   "SubalternoResidenza"
            Caption         =   "Subalterno Residenza"
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
         BeginProperty Column25 
            DataField       =   "InternoResidenza"
            Caption         =   "Interno Residenza"
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
         BeginProperty Column26 
            DataField       =   "ScalaResidenza"
            Caption         =   "Scala Residenza"
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
         BeginProperty Column27 
            DataField       =   "PianoResidenza"
            Caption         =   "Piano Residenza"
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
         BeginProperty Column28 
            DataField       =   "TipoDomicilio"
            Caption         =   "Tipo Domicilio"
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
         BeginProperty Column29 
            DataField       =   "NominativoRecapito"
            Caption         =   "Nominativo Recapito"
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
         BeginProperty Column30 
            DataField       =   "CodComuneRecapito"
            Caption         =   "CodComuneRecapito"
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
            DataField       =   "ComuneRecapito.Comune"
            Caption         =   "Comune Recapito"
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
            DataField       =   "CodViaRecapito"
            Caption         =   "CodViaRecapito"
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
            DataField       =   "IndirizzoRecapito"
            Caption         =   "Indirizzo Recapito"
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
         BeginProperty Column34 
            DataField       =   "NumCivicoRecapito"
            Caption         =   "Num Civico Recapito"
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
         BeginProperty Column35 
            DataField       =   "SubalternoRecapito"
            Caption         =   "Subalterno Recapito"
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
         BeginProperty Column36 
            DataField       =   "InternoRecapito"
            Caption         =   "Interno Recapito"
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
         BeginProperty Column37 
            DataField       =   "ScalaRecapito"
            Caption         =   "Scala Recapito"
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
         BeginProperty Column38 
            DataField       =   "PianoRecapito"
            Caption         =   "Piano Recapito"
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
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   510,236
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2924,788
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1275,024
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1980,284
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   2385,071
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   2055,118
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   2670,236
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   1620,284
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column24 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column25 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column26 
               Object.Visible         =   0   'False
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column27 
               Object.Visible         =   0   'False
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column28 
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column29 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column30 
               Object.Visible         =   0   'False
               ColumnWidth     =   1904,882
            EndProperty
            BeginProperty Column31 
               ColumnWidth     =   2340,284
            EndProperty
            BeginProperty Column32 
               Object.Visible         =   0   'False
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column33 
               ColumnWidth     =   2849,953
            EndProperty
            BeginProperty Column34 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column35 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column36 
               ColumnWidth     =   884,976
            EndProperty
            BeginProperty Column37 
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column38 
               Object.Visible         =   0   'False
               ColumnWidth     =   900,284
            EndProperty
         EndProperty
      End
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
      Index           =   4
      Left            =   6120
      TabIndex        =   11
      Top             =   8040
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   3720
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   23
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
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   8340
   End
End
Attribute VB_Name = "frmUtenti"
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
Const TXT_CODUTENTE = 5
Const TXT_BADGE = 6



Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If DataGrid1.Row <> -1 Then
                If (DataGrid1.Columns("Cognome").Text = "") Then
                    Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                Else
                    Des = nullToBlank(DataGrid1.Columns("Nome").Value) & " " & nullToBlank(DataGrid1.Columns("Cognome").Value)
                End If
                frmAutomezzi.ShowMe DataGrid1.Columns("Utente").Value, Des, "U"
            End If
            
        Case 2
            If DataGrid1.Row <> -1 Then
                If (DataGrid1.Columns("Cognome").Text = "") Then
                    Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
                Else
                    Des = nullToBlank(DataGrid1.Columns("Nome").Value) & " " & nullToBlank(DataGrid1.Columns("Cognome").Value)
                End If
                frmBadge.ShowMe DataGrid1.Columns("Utente").Value, Des
            End If
        Case 3
            lblEsportazione.Visible = True
            ProgressBar1.Visible = True
            DoEvents
            Esportazione
            
            lblEsportazione.Visible = False
            ProgressBar1.Visible = False
        
        Case 4
            ricaricaLista
        Case 5
            If (DataGrid1.Columns("Cognome").Text = "") Then
                Des = nullToBlank(DataGrid1.Columns("Ragione Sociale").Value)
            Else
                Des = nullToBlank(DataGrid1.Columns("Nome").Value) & " " & nullToBlank(DataGrid1.Columns("Cognome").Value)
            End If
        
            frmImmobili.ShowMe DataGrid1.Columns("Utente").Value, Des
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
    fileEXP = App.Path & "\" & gUtentiFile & IIf(gUtentiFileType = 0, ".TXT", ".CSV")
   
    If DataGrid1.ApproxCount <= 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = DevEnvMDB.rscmdUtenti.RecordCount
    ProgressBar1.Value = 0
    
    DataGrid1.Row = 0
    For Each ItemH In DataGrid1.Columns
        If ItemH.Visible Then
            txtStream.Write Chr$(34) & ItemH.Caption & Chr$(34)
            If ItemH.ColIndex < DataGrid1.Columns.Count Then
                txtStream.Write gUtentiSeparatore
            End If
        End If
    Next
    txtStream.Write vbCrLf
    
    
    DevEnvMDB.rscmdUtenti.MoveFirst
    Do While Not DevEnvMDB.rscmdUtenti.EOF
        ProgressBar1.Value = ProgressBar1.Value + 1
        For Each ItemH In DataGrid1.Columns
            If ItemH.Visible Then
                
                txtStream.Write formatValue(ItemH)
                If ItemH.ColIndex < DataGrid1.Columns.Count Then
                    txtStream.Write gUtentiSeparatore
                End If
            End If
        Next
        txtStream.Write vbCrLf
        DevEnvMDB.rscmdUtenti.MoveNext
    Loop
    DevEnvMDB.rscmdUtenti.MoveFirst
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    
    fso.MoveFile fileEXP, gPercorsoEsp & "\" & gUtentiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gUtentiFileType = 0, ".TXT", ".CSV")
    
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


Private Sub ricaricaLista()
    Dim search As String
    
    On Error GoTo ricaricaListaERR
    
    Screen.MousePointer = vbHourglass
    
    search = vbNullString
    
    If txt(TXT_CODICEFISCALE) <> vbNullString Then
        search = " CodFiscale like '" & txt(TXT_CODICEFISCALE) & "%'"
    End If
    
    If txt(TXT_CODUTENTE) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " CodUtente =" & txt(TXT_CODUTENTE)
    End If
    
    If txt(TXT_BADGE) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " codBadge like '" & txt(TXT_BADGE) & "%'"
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
    
    
    DevEnvMDB.rscmdUtenti.Filter = IIf(search = vbNullString, vbNullString, search & " AND ") & "CodComune='" & GetCurrentComune("CodIstat").Value & "'"
    'DevEnvMDB.rscmdUtenti.Open
    
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = DevEnvMDB
    
    
    TabStrip1.Tabs(1).Selected = True

    Screen.MousePointer = vbDefault
Exit Sub
ricaricaListaERR:
    MsgBox Error, vbCritical, App.Title
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
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Screen.MousePointer = vbDefault
    
    sourceElenco = DevEnvMDB.rscmdUtenti.Source
        
    ricaricaLista

    Exit Sub
LoadERR:
    MsgBox "Load frmutenti:" & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    resizeControl
    
    ProgressBar1.Move 15, Me.ScaleHeight - ProgressBar1.Height - 15, cmd(cmd.ubound).left - 30
    lblEsportazione.Move 15, ProgressBar1.top - lblEsportazione.Height - 15, cmd(cmd.ubound).left - 30
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
    
    Picture1(0).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    Picture1(1).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    
    DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DevEnvMDB.rscmdUtenti.Filter = vbNullString
    
    Set pfrmUtenti = Nothing
End Sub

Private Sub TabStrip1_Click()
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1, 5
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
End Sub
