VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSocietaTab 
   Caption         =   "Gestione Società"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16605
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   16605
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicSocieta 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6330
      Index           =   2
      Left            =   75
      ScaleHeight     =   6330
      ScaleWidth      =   5985
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   5985
      Begin VB.TextBox txtImpPrefissoFax 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6225
         MaxLength       =   5
         TabIndex        =   75
         Top             =   4800
         Width           =   690
      End
      Begin VB.TextBox txtImpIstatAttivita 
         DataField       =   "Cap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   64
         Top             =   4800
         Width           =   2745
      End
      Begin VB.TextBox txtImpNumAddetti 
         Alignment       =   1  'Right Justify
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6225
         MaxLength       =   50
         TabIndex        =   71
         Top             =   3465
         Width           =   3870
      End
      Begin VB.TextBox txtImpNrRea 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   9
         TabIndex        =   63
         Top             =   4815
         Width           =   2820
      End
      Begin VB.TextBox txtImpPrefissoTel 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6225
         MaxLength       =   5
         TabIndex        =   72
         Top             =   4110
         Width           =   690
      End
      Begin VB.TextBox txtImpTelefono 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7290
         MaxLength       =   10
         TabIndex        =   74
         Top             =   4110
         Width           =   2790
      End
      Begin VB.TextBox txtImpFax 
         DataField       =   "NumeroFax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   77
         Top             =   4800
         Width           =   2760
      End
      Begin VB.TextBox txtImpNumeroCivico 
         Alignment       =   1  'Right Justify
         DataField       =   "NumCivico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6225
         MaxLength       =   6
         TabIndex        =   68
         Top             =   2775
         Width           =   660
      End
      Begin VB.TextBox txtImpSubAlterno 
         DataField       =   "Interno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7875
         MaxLength       =   4
         TabIndex        =   70
         Top             =   2775
         Width           =   495
      End
      Begin VB.TextBox txtImpInterno 
         DataField       =   "Subalterno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7125
         MaxLength       =   4
         TabIndex        =   69
         Top             =   2775
         Width           =   495
      End
      Begin VB.TextBox txtImpIndirizzo 
         DataField       =   "Indirizzo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   67
         Top             =   2100
         Width           =   3885
      End
      Begin VB.TextBox txtImpCodComune 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   5490
         Width           =   990
      End
      Begin VB.TextBox txtImpRagioneSociale 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   50
         TabIndex        =   59
         Top             =   3480
         Width           =   5625
      End
      Begin VB.TextBox txtImpPartitaIva 
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   62
         Top             =   4110
         Width           =   2745
      End
      Begin VB.TextBox txtImpIdAutorizzazione 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2775
         Width           =   1245
      End
      Begin VB.TextBox txtImpCodSoggetto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2100
         Width           =   405
      End
      Begin VB.TextBox txtImpIdImpianto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2100
         Width           =   1230
      End
      Begin VB.TextBox txtImpIdSoggetto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox txtImpCodFiscale 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   16
         TabIndex        =   61
         Top             =   4110
         Width           =   2820
      End
      Begin VB.TextBox txtImpCap 
         DataField       =   "Cap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   5
         TabIndex        =   65
         Top             =   5490
         Width           =   1020
      End
      Begin MSComctlLib.ListView lvwImpianti 
         Height          =   1590
         Left            =   75
         TabIndex        =   50
         Top             =   75
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "idSoggetto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "codSoggetto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "idAutorizzazione"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "RagioneSociale"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodFiscale"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PartitaIva"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cap"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CodComune"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Indirizzo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "NumCivico"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Subalterno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Interno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "NumeroTelefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "NumeroFax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "CodIstat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Nr.REA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Tot.Addetti"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataListLib.DataCombo ctlImpComune 
         Bindings        =   "FrmSocietaTab.frx":0000
         DataField       =   "CODISTAT"
         Height          =   375
         Left            =   2400
         TabIndex        =   66
         Top             =   5490
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   661
         _Version        =   393216
         ListField       =   "Comune"
         BoundColumn     =   "CODISTAT"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo ctlImpAutorizzazioni 
         Bindings        =   "FrmSocietaTab.frx":0024
         DataField       =   "idAutorizzazione"
         Height          =   375
         Left            =   1395
         TabIndex        =   58
         Top             =   2760
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   661
         _Version        =   393216
         ListField       =   "Descrizione"
         BoundColumn     =   "idAutorizzazione"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Impianto"
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
         Index           =   22
         Left            =   75
         TabIndex        =   91
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   43
         Left            =   7035
         TabIndex        =   90
         Top             =   4725
         Width           =   120
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "ISTAT Attività"
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
         Index           =   42
         Left            =   2985
         TabIndex        =   89
         Top             =   4575
         Width           =   1245
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Totale Addetti"
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
         Index           =   41
         Left            =   6240
         TabIndex        =   88
         Top             =   3225
         Width           =   1200
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nr Iscrizione REA"
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
         Index           =   40
         Left            =   90
         TabIndex        =   87
         Top             =   4575
         Width           =   1545
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   7035
         TabIndex        =   86
         Top             =   4050
         Width           =   120
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
         Index           =   27
         Left            =   6240
         TabIndex        =   85
         Top             =   3900
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   6240
         TabIndex        =   84
         Top             =   4575
         Width           =   330
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nr Civico/Interno/Sub.Alt"
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
         Index           =   25
         Left            =   6240
         TabIndex        =   83
         Top             =   2550
         Width           =   2115
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo"
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
         Index           =   24
         Left            =   6240
         TabIndex        =   82
         Top             =   1875
         Width           =   1170
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comune"
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
         Index           =   23
         Left            =   1245
         TabIndex        =   81
         Top             =   5265
         Width           =   795
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   19
         Left            =   3000
         TabIndex        =   80
         Top             =   3870
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Autorizzazione"
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
         Index           =   21
         Left            =   90
         TabIndex        =   79
         Top             =   2550
         Width           =   1290
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "RagioneSociale"
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
         Index           =   20
         Left            =   90
         TabIndex        =   78
         Top             =   3225
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "CodFiscale"
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
         Left            =   90
         TabIndex        =   76
         Top             =   3870
         Width           =   975
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
         Index           =   28
         Left            =   90
         TabIndex        =   73
         Top             =   5265
         Width           =   345
      End
      Begin VB.Shape shpImp 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   75
         Top             =   1725
         Width           =   5790
      End
   End
   Begin VB.PictureBox picRight 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   6480
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   113
      Top             =   7920
      Width           =   4815
      Begin VB.CommandButton cmd 
         Caption         =   "Nuovo"
         DisabledPicture =   "FrmSocietaTab.frx":0048
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   3
         Left            =   0
         Picture         =   "FrmSocietaTab.frx":689A
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Inserisce un nuovo elemento"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Cancella"
         DisabledPicture =   "FrmSocietaTab.frx":D0EC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   2
         Left            =   1230
         Picture         =   "FrmSocietaTab.frx":1393E
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Cancella l'elemento corrente"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Salva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   1
         Left            =   2475
         Picture         =   "FrmSocietaTab.frx":1A190
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Salva le modifiche correnti"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Chiudi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   0
         Left            =   3705
         Picture         =   "FrmSocietaTab.frx":209E2
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Annulla senza salvare"
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox PicSocieta 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5355
      Index           =   1
      Left            =   6600
      ScaleHeight     =   5355
      ScaleWidth      =   9615
      TabIndex        =   34
      Top             =   5160
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CheckBox ChkAutoDefault 
         Caption         =   "DEFAULT"
         DownPicture     =   "FrmSocietaTab.frx":27234
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   4875
         Picture         =   "FrmSocietaTab.frx":27667
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3675
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker DTPAutoInizioValidita 
         Height          =   390
         Left            =   75
         TabIndex        =   44
         Top             =   3675
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57475073
         CurrentDate     =   39934
      End
      Begin VB.TextBox txtAutoRilasciataDa 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         MaxLength       =   50
         TabIndex        =   40
         Top             =   4500
         Width           =   4635
      End
      Begin VB.TextBox txtAutoNumero 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         MaxLength       =   50
         TabIndex        =   39
         Top             =   2925
         Width           =   4635
      End
      Begin VB.TextBox txtAutoCodSoggetto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2250
         Width           =   405
      End
      Begin VB.TextBox txtAutoIDSoggetto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2250
         Width           =   930
      End
      Begin VB.TextBox txtAutoIdAutorizzazione 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2250
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvwAutorizzazioni 
         Height          =   1740
         Left            =   75
         TabIndex        =   35
         Top             =   75
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3069
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "id"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "idSoggetto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "codSoggetto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Numero"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DataInizioValidita"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DataScadenza"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "RilasciataDa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "flgDefault"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPAutoFineValidita 
         Height          =   390
         Left            =   2700
         TabIndex        =   46
         Top             =   3675
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57475073
         CurrentDate     =   39934
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   75
         Top             =   1875
         Width           =   5790
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Rilasciata Da"
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
         Index           =   17
         Left            =   75
         TabIndex        =   47
         Top             =   4200
         Width           =   1155
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Data Fine Validita"
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
         Left            =   2700
         TabIndex        =   45
         Top             =   3375
         Width           =   1545
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Data Inizio Validita"
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
         Left            =   75
         TabIndex        =   43
         Top             =   3375
         Width           =   1620
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
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
         Left            =   75
         TabIndex        =   42
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Autorizzazione"
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
         Left            =   105
         TabIndex        =   41
         Top             =   2025
         Width           =   1290
      End
   End
   Begin VB.PictureBox PicSocieta 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Index           =   0
      Left            =   150
      ScaleHeight     =   6855
      ScaleWidth      =   13905
      TabIndex        =   18
      Top             =   150
      Visible         =   0   'False
      Width           =   13905
      Begin VB.Frame Frame1 
         Caption         =   "Sede Legale"
         ForeColor       =   &H000000FF&
         Height          =   3090
         Left            =   5310
         TabIndex        =   93
         Top             =   1545
         Width           =   5040
         Begin VB.TextBox txtCodComuneLegale 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "CodComune"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox txtIndirizzoLegale 
            DataField       =   "Indirizzo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   100
            Top             =   1230
            Width           =   3600
         End
         Begin VB.TextBox txtCapLegale 
            DataField       =   "Cap"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   5
            TabIndex        =   99
            Top             =   1215
            Width           =   1020
         End
         Begin VB.TextBox txtNrInternoLegale 
            DataField       =   "Subalterno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            MaxLength       =   4
            TabIndex        =   98
            Top             =   1845
            Width           =   495
         End
         Begin VB.TextBox txtSubAlternoLegale 
            DataField       =   "Interno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1515
            MaxLength       =   4
            TabIndex        =   97
            Top             =   1845
            Width           =   495
         End
         Begin VB.TextBox txtNrCivicoLegale 
            DataField       =   "NumCivico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   6
            TabIndex        =   96
            Top             =   1845
            Width           =   660
         End
         Begin VB.TextBox txtNumTelefonoLegale 
            DataField       =   "NumeroTelefono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1095
            MaxLength       =   10
            TabIndex        =   95
            Top             =   2580
            Width           =   3720
         End
         Begin VB.TextBox txtPrefissoTelefonoLegale 
            DataField       =   "NumeroTelefono"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   5
            TabIndex        =   94
            Top             =   2580
            Width           =   690
         End
         Begin MSDataListLib.DataCombo ctlComuneLegale 
            Bindings        =   "FrmSocietaTab.frx":27AB7
            DataField       =   "CODISTAT"
            Height          =   375
            Left            =   1200
            TabIndex        =   102
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            _Version        =   393216
            ListField       =   "Comune"
            BoundColumn     =   "CODISTAT"
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Indirizzo"
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
            Index           =   31
            Left            =   1275
            TabIndex        =   108
            Top             =   975
            Width           =   720
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comune"
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
            Index           =   30
            Left            =   120
            TabIndex        =   107
            Top             =   225
            Width           =   720
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   32
            Left            =   150
            TabIndex        =   106
            Top             =   990
            Width           =   345
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nr Civico/Interno/Sub.Alt"
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
            Index           =   33
            Left            =   150
            TabIndex        =   105
            Top             =   1635
            Width           =   2115
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
            Index           =   34
            Left            =   120
            TabIndex        =   104
            Top             =   2310
            Width           =   780
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   855
            TabIndex        =   103
            Top             =   2475
            Width           =   120
         End
      End
      Begin VB.TextBox txtPrefissoFax 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   11
         Top             =   6000
         Width           =   690
      End
      Begin VB.TextBox txtNomeLegale 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         MaxLength       =   25
         TabIndex        =   17
         Top             =   5040
         Width           =   2220
      End
      Begin VB.TextBox txtCognomeLegale 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5460
         MaxLength       =   25
         TabIndex        =   16
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox txtPrefissoTel 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5460
         MaxLength       =   5
         TabIndex        =   13
         Top             =   375
         Width           =   690
      End
      Begin VB.TextBox txtCapSociale 
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2430
         Width           =   4650
      End
      Begin VB.TextBox txtemail 
         DataField       =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5460
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1080
         Width           =   4650
      End
      Begin VB.TextBox txtLogo 
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   30
         TabIndex        =   4
         Top             =   3090
         Width           =   4650
      End
      Begin VB.TextBox txtInterno 
         DataField       =   "Subalterno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   4
         TabIndex        =   9
         Top             =   5325
         Width           =   495
      End
      Begin VB.TextBox txtSubAlterno 
         DataField       =   "Interno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         MaxLength       =   4
         TabIndex        =   10
         Top             =   5325
         Width           =   495
      End
      Begin VB.TextBox txtCap 
         DataField       =   "Cap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   6
         Top             =   4560
         Width           =   1020
      End
      Begin VB.TextBox txtNumeroFax 
         DataField       =   "NumeroFax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   12
         Top             =   6000
         Width           =   3585
      End
      Begin VB.TextBox txtNumeroTelefono 
         DataField       =   "NumeroTelefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6435
         MaxLength       =   10
         TabIndex        =   14
         Top             =   375
         Width           =   3015
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   6
         TabIndex        =   8
         Top             =   5325
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4605
         Width           =   3525
      End
      Begin VB.TextBox txtCodComune 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3840
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   11
         TabIndex        =   2
         Top             =   1725
         Width           =   2205
      End
      Begin VB.TextBox txtCodFiscale 
         DataField       =   "CodFiscale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1725
         Width           =   2295
      End
      Begin VB.TextBox txtRagioneSociale 
         DataField       =   "RagioneSociale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1080
         Width           =   4635
      End
      Begin VB.TextBox txtidSocieta 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "idSocieta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   375
         Width           =   1980
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "FrmSocietaTab.frx":27ADB
         DataField       =   "CODISTAT"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         ListField       =   "Comune"
         BoundColumn     =   "CODISTAT"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
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
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   930
         TabIndex        =   92
         Top             =   5925
         Width           =   120
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nome Leg.Rap."
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
         Index           =   38
         Left            =   7785
         TabIndex        =   53
         Top             =   4770
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cognome Leg.Rap."
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
         Index           =   37
         Left            =   5460
         TabIndex        =   52
         Top             =   4770
         Width           =   1665
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   6195
         TabIndex        =   51
         Top             =   300
         Width           =   120
      End
      Begin VB.Image ImageLogo 
         Appearance      =   0  'Flat
         Height          =   4335
         Left            =   10305
         Top             =   75
         Width           =   3600
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cap.Sociale"
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
         Left            =   120
         TabIndex        =   33
         Top             =   2145
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Email"
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
         Left            =   5460
         TabIndex        =   32
         Top             =   810
         Width           =   495
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Logo"
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
         Left            =   120
         TabIndex        =   31
         Top             =   2820
         Width           =   420
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
         Left            =   105
         TabIndex        =   30
         Top             =   4335
         Width           =   405
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Left            =   120
         TabIndex        =   29
         Top             =   5715
         Width           =   330
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
         Left            =   5460
         TabIndex        =   28
         Top             =   135
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nr Civico/Interno/Sub.Alt"
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
         Left            =   150
         TabIndex        =   27
         Top             =   5055
         Width           =   2115
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo"
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
         Left            =   1245
         TabIndex        =   26
         Top             =   4365
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comune"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3570
         Width           =   720
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   24
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Fiscale"
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
         TabIndex        =   23
         Top             =   1455
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   810
         Width           =   1410
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Societa"
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
         TabIndex        =   21
         Top             =   135
         Width           =   660
      End
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1560
      ScaleHeight     =   1215
      ScaleWidth      =   3615
      TabIndex        =   109
      Top             =   7680
      Width           =   3615
      Begin VB.CommandButton cmd 
         Caption         =   "Autorizz."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   5
         Left            =   1200
         Picture         =   "FrmSocietaTab.frx":27AFF
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Pagina Avanti"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Societa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   6
         Left            =   0
         Picture         =   "FrmSocietaTab.frx":2E351
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Pagina Indietro"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Impianto"
         DisabledPicture =   "FrmSocietaTab.frx":34BA3
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Index           =   4
         Left            =   2400
         Picture         =   "FrmSocietaTab.frx":3B3F5
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Visualizza elementi come Icone"
         Top             =   0
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":41C47
            Key             =   "Autorizzazione"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":484A9
            Key             =   "Impianto"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":4ED0B
            Key             =   "Società"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":5556D
            Key             =   "Upd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":5BDCF
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSocietaTab.frx":62631
            Key             =   "New"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSocietaTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHdrIcons As New cLVHeaderSortIcons

Dim rsSocieta As Recordset
Dim rsAutorizzazioni As Recordset
Dim rsImpianti As Recordset
Dim rsComuni1 As Recordset
Dim rsComuni2 As Recordset
Dim rsComuni3 As Recordset
Dim rsAutorizzazioniSoc As Recordset
Dim flgModRecord As Boolean
Dim LastIndexPage As Integer

Public Sub ShowMe()
    Load Me
    
    PicSocieta(0).Visible = True
    
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        Me.Show
    Else
        Me.Show
    End If
    
    cmd_Click 6
   
End Sub




'[10.35.10] Riccardo Sammito Ufficio : 1024x600
Private Sub cmd_Click(Index As Integer)
    
    If flgModRecord And (Index = 4 Or Index = 5 Or Index = 6 Or Index = 0) Then
        If MsgBox("Attenzione Esistono delle modifiche non salvate, continuando andranno perse,si vuole continuare?", vbQuestion + vbYesNo) = vbYes Then
            If Not rsAutorizzazioni.EOF And Not rsAutorizzazioni.BOF Then
                rsAutorizzazioni.CancelUpdate
            End If
            If Not rsImpianti.EOF And Not rsImpianti.BOF Then
                rsImpianti.CancelUpdate
            End If
            flgModRecord = False
        End If
    End If
    
    Select Case Index
        Case 0
            Unload Me
        Case 1, 2, 3
            If PicSocieta(1).Visible Then
                cmdAutorizzazioni_Click Index
            ElseIf PicSocieta(2).Visible Then
                cmdImpianti_click Index
            ElseIf Index = 1 Then
                SalvaSocieta
                
                CaricaSocieta
            End If
        Case 4, 5, 6
            PicSocieta(LastIndexPage).Visible = False
            PicSocieta(6 - Index).Visible = True
            LastIndexPage = 6 - Index
    End Select
End Sub


Private Sub RefreshCommand()
    cmd(4).Visible = rsAutorizzazioni.RecordCount > 0
    cmd(5).Visible = rsSocieta.RecordCount > 0
End Sub


Private Function ValidaImpianti() As Boolean
    
    ctlImpComune.BackColor = &H80000005
    ctlImpComune.ToolTipText = vbNullString
    ctlImpAutorizzazioni.BackColor = &H80000005
    ctlImpAutorizzazioni.ToolTipText = ""
    txtImpCap.BackColor = &H80000005
    txtImpCap.ToolTipText = ""
    txtImpIndirizzo.BackColor = &H80000005
    txtImpIndirizzo.ToolTipText = ""
    txtImpNumeroCivico.BackColor = &H80000005
    txtImpNumeroCivico.ToolTipText = ""
    txtImpTelefono.BackColor = &H80000005
    txtImpTelefono.ToolTipText = ""
    txtImpIstatAttivita.BackColor = &H80000005
    txtImpIstatAttivita.ToolTipText = ""
    txtImpPrefissoTel.BackColor = &H80000005
    txtImpPrefissoTel.ToolTipText = ""
    txtImpNrRea.BackColor = &H80000005
    txtImpNrRea.ToolTipText = ""
    txtImpNumAddetti.BackColor = &H80000005
    txtImpNumAddetti.ToolTipText = ""
    txtImpCodFiscale.BackColor = &H80000005
    txtImpCodFiscale.ToolTipText = ""
    txtImpPartitaIva.BackColor = &H80000005
    txtImpPartitaIva.ToolTipText = ""
    
    ValidaImpianti = True
    If txtImpCodFiscale.Text = vbNullString And txtImpPartitaIva.Text = vbNullString Then
        txtImpCodFiscale.BackColor = &H8080FF
        txtImpCodFiscale.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
        ValidaImpianti = False
        txtImpPartitaIva.BackColor = &H8080FF
        txtImpPartitaIva.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
    End If
    
    If txtImpCodComune = vbNullString Then
        ctlImpComune.BackColor = &H8080FF
        ctlImpComune.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
    
    If txtImpIdAutorizzazione = vbNullString Then
        ctlImpAutorizzazioni.BackColor = &H8080FF
        ctlImpAutorizzazioni.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
        

    If txtImpCap = vbNullString Then
        txtImpCap.BackColor = &H8080FF
        txtImpCap.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If

    If txtImpIndirizzo = vbNullString Then
        txtImpIndirizzo.BackColor = &H8080FF
        txtImpIndirizzo.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
    
    
    If txtImpNumeroCivico = vbNullString Then
        txtImpNumeroCivico.BackColor = &H8080FF
        txtImpNumeroCivico.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
    
    If txtImpTelefono = vbNullString Then
        txtImpTelefono.BackColor = &H8080FF
        txtImpTelefono.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
        
    If txtImpIstatAttivita = vbNullString Then
        txtImpIstatAttivita.BackColor = &H8080FF
        txtImpIstatAttivita.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
    
    If txtImpPrefissoTel = vbNullString Then
        txtImpPrefissoTel.BackColor = &H8080FF
        txtImpPrefissoTel.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
        
    If txtImpNrRea = vbNullString Then
        txtImpNrRea.BackColor = &H8080FF
        txtImpNrRea.ToolTipText = "Il Campo è Obbligatorio"
        ValidaImpianti = False
    End If
    
    If gflgAttiviMUD Then
        If txtImpNumAddetti = vbNullString Then
            txtImpNumAddetti.BackColor = &H8080FF
            txtImpNumAddetti.ToolTipText = "Il Campo è Obbligatorio"
            ValidaImpianti = False
        End If
    End If
    
End Function

Private Function ValidaAutorizzazioni() As Boolean
    ValidaAutorizzazioni = True

    txtAutoNumero.BackColor = &H80000005
    txtAutoNumero.ToolTipText = ""
    txtAutoRilasciataDa.BackColor = &H80000005
    txtAutoRilasciataDa.ToolTipText = ""
    ChkAutoDefault.BackColor = &H80000005
    ChkAutoDefault.ToolTipText = ""
    DTPAutoInizioValidita.CalendarBackColor = &H80000005
    DTPAutoInizioValidita.ToolTipText = ""
    
    If txtAutoRilasciataDa.Text = vbNullString Then
        txtAutoRilasciataDa.BackColor = &H8080FF
        txtAutoRilasciataDa.ToolTipText = "Il Campo è Obbligatorio"
        ValidaAutorizzazioni = False
    End If

    If txtAutoNumero.Text = vbNullString Then
        txtAutoNumero.BackColor = &H8080FF
        txtAutoNumero.ToolTipText = "Il Campo è Obbligatorio"
        ValidaAutorizzazioni = False
    End If
    
    If DTPAutoInizioValidita > DTPAutoFineValidita Then
        DTPAutoInizioValidita.CalendarBackColor = &H8080FF
        DTPAutoInizioValidita.ToolTipText = "La data inizio validità deve essere minore della data fine validità"
        ValidaAutorizzazioni = False
    
    End If
    
    If getAutorizzazioneDiDefault = -1 And ChkAutoDefault.Value = 0 Then
        ChkAutoDefault.BackColor = &H8080FF
        ChkAutoDefault.ToolTipText = "Attenzione deve esistere almeno un autorizzazione di DEFAULT"
        ValidaAutorizzazioni = False
    End If

End Function

Private Sub SalvaAutorizzazioni()
    Set rsAutorizzazioni.ActiveConnection = gDBC
    
    rsAutorizzazioni.UpdateBatch adAffectAllChapters
    
    Set rsAutorizzazioni.ActiveConnection = Nothing
    
    caricaListaAutorizzazioni rsSocieta!idSocieta, rsSocieta!codSoggetto
End Sub


Private Sub SalvaImpianti()
    
    Set rsImpianti.ActiveConnection = gDBC
    
    rsImpianti.UpdateBatch adAffectAllChapters
    
    Set rsImpianti.ActiveConnection = Nothing
End Sub



Private Sub SalvaSocieta()
    'validazione campi
    txtCap.BackColor = &H80000005
    txtCap.ToolTipText = vbNullString
    ctlComune.BackColor = &H80000005
    ctlComune.ToolTipText = vbNullString
    txtCodFiscale.BackColor = &H80000005
    txtCodFiscale.ToolTipText = vbNullString
    txtPartitaIva.BackColor = &H80000005
    txtPartitaIva.ToolTipText = vbNullString
    txtIndirizzo.BackColor = &H80000005
    txtIndirizzo.ToolTipText = vbNullString
    txtNumCivico.BackColor = &H80000005
    txtNumCivico.ToolTipText = vbNullString
    txtPrefissoTel.BackColor = &H80000005
    txtPrefissoTel.ToolTipText = vbNullString
    txtNumeroTelefono.BackColor = &H80000005
    txtNumeroTelefono.ToolTipText = vbNullString
    txtRagioneSociale.BackColor = &H80000005
    txtRagioneSociale.ToolTipText = vbNullString
    txtCapSociale.BackColor = &H80000005
    txtCapSociale.ToolTipText = vbNullString
    txtCapLegale.BackColor = &H80000005
    txtCapLegale.ToolTipText = vbNullString
    ctlComuneLegale.BackColor = &H80000005
    ctlComuneLegale.ToolTipText = vbNullString
    txtIndirizzoLegale.BackColor = &H80000005
    txtIndirizzoLegale.ToolTipText = vbNullString
    txtNrCivicoLegale.BackColor = &H80000005
    txtNrCivicoLegale.ToolTipText = vbNullString
    txtNrInternoLegale.BackColor = &H80000005
    txtNrInternoLegale.ToolTipText = vbNullString
    txtPrefissoTelefonoLegale.BackColor = &H80000005
    txtPrefissoTelefonoLegale.ToolTipText = vbNullString
    txtCognomeLegale.BackColor = &H80000005
    txtCognomeLegale.ToolTipText = vbNullString
    txtCognomeLegale.BackColor = &H80000005
    txtCognomeLegale.ToolTipText = vbNullString
    txtNomeLegale.BackColor = &H80000005
    txtNomeLegale.ToolTipText = vbNullString
    
    If txtCap.Text = vbNullString Then
        txtCap.BackColor = &H8080FF
        txtCap.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtCodComune.Text = vbNullString Then
        ctlComune.BackColor = &H8080FF
        ctlComune.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtCodFiscale.Text = vbNullString And txtPartitaIva.Text = vbNullString Then
        txtCodFiscale.BackColor = &H8080FF
        txtCodFiscale.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
        
        txtPartitaIva.BackColor = &H8080FF
        txtPartitaIva.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
    End If
    
    If txtIndirizzo.Text = vbNullString Then
        txtIndirizzo.BackColor = &H8080FF
        txtIndirizzo.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtNumCivico.Text = vbNullString Then
        txtNumCivico.BackColor = &H8080FF
        txtNumCivico.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtPrefissoTel.Text = vbNullString Then
        txtPrefissoTel.BackColor = &H8080FF
        txtPrefissoTel.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtNumeroTelefono.Text = vbNullString Then
        txtNumeroTelefono.BackColor = &H8080FF
        txtNumeroTelefono.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtRagioneSociale.Text = vbNullString Then
        txtRagioneSociale.BackColor = &H8080FF
        txtRagioneSociale.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    If txtCapSociale.Text = vbNullString Then
        txtCapSociale.BackColor = &H8080FF
        txtCapSociale.ToolTipText = "Il Campo è Obbligatorio"
    End If
    
    
    'validare solo se attiva MUD
    If gflgAttiviMUD Then
        If txtCapLegale.Text = vbNullString Then
            txtCapLegale.BackColor = &H8080FF
            txtCapLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtCodComuneLegale.Text = vbNullString Then
            ctlComuneLegale.BackColor = &H8080FF
            ctlComuneLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtIndirizzoLegale.Text = vbNullString Then
            txtIndirizzoLegale.BackColor = &H8080FF
            txtIndirizzoLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtNrCivicoLegale.Text = vbNullString Then
            txtNrCivicoLegale.BackColor = &H8080FF
            txtNrCivicoLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtNrInternoLegale.Text = vbNullString Then
            txtNrInternoLegale.BackColor = &H8080FF
            txtNrInternoLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtPrefissoTelefonoLegale.Text = vbNullString Then
            txtPrefissoTelefonoLegale.BackColor = &H8080FF
            txtPrefissoTelefonoLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtCognomeLegale.Text = vbNullString Then
            txtCognomeLegale.BackColor = &H8080FF
            txtCognomeLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtCognomeLegale.Text = vbNullString Then
            txtCognomeLegale.BackColor = &H8080FF
            txtCognomeLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
        
        If txtNomeLegale.Text = vbNullString Then
            txtNomeLegale.BackColor = &H8080FF
            txtNomeLegale.ToolTipText = "Il Campo è Obbligatorio"
        End If
    End If
    
    rsSocieta!Cap = nullToBlank(txtCap.Text)
    rsSocieta!codComune = txtCodComune.Text
    rsSocieta!CodFiscale = txtCodFiscale.Text
    rsSocieta!PartitaIva = IIf(txtPartitaIva.Text = vbNullString, Null, txtPartitaIva.Text)
    rsSocieta!Indirizzo = txtIndirizzo.Text
    
    rsSocieta!NumCivico = toZero(txtNumCivico.Text)
    rsSocieta!Subalterno = txtSubalterno.Text
    rsSocieta!Interno = txtInterno.Text
    rsSocieta!PrefissoTel = txtPrefissoTel.Text
    rsSocieta!NumeroTelefono = txtNumeroTelefono.Text
    rsSocieta!NumeroFax = txtNumeroFax.Text
    rsSocieta!PrefissoFax = txtPrefissoFax.Text
    rsSocieta!RagioneSociale = txtRagioneSociale.Text
    rsSocieta!Logo = txtLogo.Text
    rsSocieta!Email = txtemail.Text
    rsSocieta!CapSociale = toZero(txtCapSociale.Text)
    
    
    rsSocieta!CapSL = txtCapLegale.Text
    rsSocieta!CodComuneSL = IIf(txtCodComuneLegale.Text = vbNullString, Null, txtCodComuneLegale.Text)
    rsSocieta!IndirizzoSL = IIf(txtIndirizzoLegale.Text = vbNullString, Null, txtIndirizzoLegale.Text)
    rsSocieta!NumCivicoSL = IIf(txtNrCivicoLegale.Text = vbNullString, Null, txtNrCivicoLegale.Text)
    rsSocieta!InternoSL = IIf(txtNrInternoLegale.Text = vbNullString, Null, txtNrInternoLegale.Text)
    rsSocieta!SubalternoSL = IIf(txtSubAlternoLegale.Text = vbNullString, Null, txtSubAlternoLegale.Text)
    rsSocieta!PrefissoTelSL = IIf(txtPrefissoTelefonoLegale.Text = vbNullString, Null, txtPrefissoTelefonoLegale.Text)
    rsSocieta!NumeroTelefonoSL = IIf(txtNumTelefonoLegale.Text = vbNullString, Null, txtNumTelefonoLegale.Text)
    rsSocieta!CognomeLegaleR = IIf(txtCognomeLegale.Text = vbNullString, Null, txtCognomeLegale.Text)
    rsSocieta!NomeLegaleR = IIf(txtNomeLegale.Text = vbNullString, Null, txtNomeLegale.Text)
    
    Set rsSocieta.ActiveConnection = gDBC
    rsSocieta.UpdateBatch adAffectAllChapters
    Set rsSocieta.ActiveConnection = Nothing
    
End Sub



Private Sub ctlComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtCodComune.SetFocus
        txtCodComune.Text = ctlComune.BoundText
        txtCap.Text = findCap(txtCodComune.Text)
    End If
End Sub

Private Sub ctlComuneLegale_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtCodComuneLegale.SetFocus
        txtCodComuneLegale.Text = ctlComuneLegale.BoundText
        txtCapLegale.Text = findCap(txtCodComuneLegale.Text)
    End If
End Sub

Private Sub ctlImpAutorizzazioni_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtImpIdAutorizzazione.SetFocus
        txtImpIdAutorizzazione.Text = ctlImpAutorizzazioni.BoundText
    End If
End Sub

Private Sub ctlImpComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtImpCodComune.SetFocus
        txtImpCodComune.Text = ctlImpComune.BoundText
        txtImpCap.Text = findCap(txtImpCodComune.Text)
    End If
End Sub



Private Sub Form_Load()
    lvwAutorizzazioni.Icons = ImageList1
    lvwImpianti.Icons = ImageList1
    lvwAutorizzazioni.SmallIcons = ImageList1
    lvwImpianti.SmallIcons = ImageList1
    
    
    
    Set ctlComune.DataSource = Nothing
    Set ctlComune.RowSource = Nothing
    
    Set rsComuni1 = CreateObject("ADODB.Recordset")
    rsComuni1.LockType = adLockBatchOptimistic
    rsComuni1.CursorLocation = adUseClient
    rsComuni1.CursorType = adOpenStatic
    
    Set rsComuni1.ActiveConnection = gDBC
    
    rsComuni1.Source = "Select comune,CODISTAT,cap from comuni order by comune"
    rsComuni1.Open
    
    Set rsComuni1.ActiveConnection = Nothing
    
    ctlComune.CausesValidation = False
    
    Set ctlComune.DataSource = rsComuni1
    Set ctlComune.RowSource = rsComuni1
    
    'comune 2
    Set ctlComuneLegale.DataSource = Nothing
    Set ctlComuneLegale.RowSource = Nothing
    Set rsComuni2 = rsComuni1.Clone
    Set ctlComuneLegale.DataSource = rsComuni2
    Set ctlComuneLegale.RowSource = rsComuni2
    
    'comune 3
    Set ctlImpComune.DataSource = Nothing
    Set ctlImpComune.RowSource = Nothing
    Set rsComuni3 = rsComuni1.Clone
    Set ctlImpComune.DataSource = rsComuni3
    Set ctlImpComune.RowSource = rsComuni3
    

    ImageList1.ListImages.Add , "AutorizzazioneNew", ImageList1.Overlay("Autorizzazione", "New")
    ImageList1.ListImages.Add , "AutorizzazioneUpd", ImageList1.Overlay("Autorizzazione", "Upd")
    ImageList1.ListImages.Add , "AutorizzazioneDel", ImageList1.Overlay("Autorizzazione", "Del")
    
    
    ImageList1.ListImages.Add , "ImpiantoNew", ImageList1.Overlay("Impianto", "New")
    ImageList1.ListImages.Add , "ImpiantoUpd", ImageList1.Overlay("Impianto", "Upd")
    ImageList1.ListImages.Add , "ImpiantoDel", ImageList1.Overlay("Impianto", "Del")
    
    
    lvwAutorizzazioni.Arrange = lvwAutoLeft
    lvwImpianti.Arrange = lvwAutoLeft

    CaricaSocieta

    
    RefreshCommand
    
    If lvwImpianti.ListItems.Count > 0 Then
        lvwImpianti.SelectedItem = lvwImpianti.ListItems(1)
        lvwImpianti_ItemClick lvwImpianti.SelectedItem
    End If
    
    If lvwAutorizzazioni.ListItems.Count > 0 Then
        lvwAutorizzazioni.SelectedItem = lvwAutorizzazioni.ListItems(1)
        lvwAutorizzazioni_ItemClick lvwAutorizzazioni.SelectedItem
    End If


End Sub


Private Sub caricaListaAutorizzazioni(idSoggetto As Long, codSogg As String)


    Set rsAutorizzazioniSoc = CreateObject("ADODB.Recordset")
    rsAutorizzazioniSoc.LockType = adLockBatchOptimistic
    rsAutorizzazioniSoc.CursorLocation = adUseClient
    rsAutorizzazioniSoc.CursorType = adOpenStatic
    
    Set rsAutorizzazioniSoc.ActiveConnection = gDBC
    
    rsAutorizzazioniSoc.Source = "select *,numero + ' del ' + Format(datainiziovalidita, 'dd/mm/yyyy') + ' scad. ' + Format(datascadenza, 'dd/mm/yyyy') as Descrizione from Autorizzazioni where CodSoggetto='" & codSogg & "' and idSoggetto=" & idSoggetto
    rsAutorizzazioniSoc.Open
    
    Set ctlImpAutorizzazioni.DataSource = Nothing
    Set ctlImpAutorizzazioni.RowSource = Nothing
    Set ctlImpAutorizzazioni.DataSource = rsAutorizzazioniSoc
    Set ctlImpAutorizzazioni.RowSource = rsAutorizzazioniSoc
    
    
    If rsAutorizzazioniSoc.EOF Then Exit Sub
    
    Set rsAutorizzazioniSoc.ActiveConnection = Nothing




End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    picLeft.Move 0, Me.ScaleHeight - picLeft.Height - 150
    picRight.Move Me.ScaleWidth - picRight.width - 150, picLeft.top
    For i = PicSocieta.LBound To PicSocieta.ubound
        PicSocieta(i).Move Me.ScaleLeft + 150, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight - picLeft.Height - 150
    Next

End Sub




Private Sub Form_Unload(Cancel As Integer)
    Dim lStyle As Long
    
    If gStartMode = 0 Or gStartMode = 1 Then
        lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
        lStyle = lStyle Or MY_WS_DOCK
        SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    
        If gStartMode = 1 Then
            SetParent Me.hWnd, hOldParent
        End If
    End If
    
    On Error Resume Next
    rsSocieta.Close
    Set rsSocieta = Nothing

    rsAutorizzazioni.Close
    Set rsAutorizzazioni = Nothing

    rsComuni1.Close
    Set rsComuni1 = Nothing

    rsImpianti.Close
    Set rsImpianti = Nothing

    rsComuni3.Close
    Set rsComuni3 = Nothing
    
    rsComuni2.Close
    Set rsComuni2 = Nothing

    rsAutorizzazioniSoc.Close
    Set rsAutorizzazioniSoc = Nothing

    Set pfrmSocieta = Nothing

End Sub

Private Sub txtImpFax_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtImpNumAddetti_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtImpNumeroCivico_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtImpPartitaIva_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtImpNrRea_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

End Sub



Private Sub txtImpTelefono_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLogo_change()
    PicSocieta_Resize (0)
    
End Sub

Private Function CaricaSocieta() As Boolean
    
    Set rsSocieta = CreateObject("ADODB.Recordset")
    rsSocieta.LockType = adLockBatchOptimistic
    rsSocieta.CursorLocation = adUseClient
    rsSocieta.CursorType = adOpenStatic
    
    Set rsSocieta.ActiveConnection = gDBC
    
    rsSocieta.Source = "select * from societa"
    rsSocieta.Open
    Set rsSocieta.ActiveConnection = Nothing
    
    CaricaSocieta = False
    
    If rsSocieta.EOF Then
        caricaAutorizzazioni 0, "A"
        caricaImpianti 0, "A"
        caricaListaAutorizzazioni 0, "A"
        
    Else
        caricaAutorizzazioni rsSocieta!idSocieta, rsSocieta!codSoggetto
        caricaImpianti rsSocieta!idSocieta, rsSocieta!codSoggetto
        caricaListaAutorizzazioni rsSocieta!idSocieta, rsSocieta!codSoggetto
    End If
    
    
    If rsSocieta.EOF Then rsSocieta.AddNew
    
    CaricaSocieta = True
    
    FillSocieta
End Function


Private Sub FillSocieta()
    txtCap.Text = nullToBlank(rsSocieta!Cap)
    txtCodComune.Text = nullToBlank(rsSocieta!codComune)
    txtCodFiscale.Text = nullToBlank(rsSocieta!CodFiscale)
    txtidSocieta.Text = nullToBlank(rsSocieta!idSocieta)
    txtIndirizzo.Text = nullToBlank(rsSocieta!Indirizzo)
    txtNumCivico.Text = nullToBlank(rsSocieta!NumCivico)
    txtSubalterno.Text = nullToBlank(rsSocieta!Subalterno)
    txtInterno.Text = nullToBlank(rsSocieta!Interno)
    txtPrefissoTel.Text = nullToBlank(rsSocieta!PrefissoTel)
    txtNumeroFax.Text = nullToBlank(rsSocieta!NumeroFax)
    txtPrefissoFax.Text = nullToBlank(rsSocieta!PrefissoFax)
    txtNumeroTelefono.Text = nullToBlank(rsSocieta!NumeroTelefono)
    txtPartitaIva.Text = nullToBlank(rsSocieta!PartitaIva)
    txtRagioneSociale.Text = nullToBlank(rsSocieta!RagioneSociale)
    txtLogo.Text = nullToBlank(rsSocieta!Logo)
    txtemail.Text = nullToBlank(rsSocieta!Email)
    txtCapSociale.Text = nullToBlank(rsSocieta!CapSociale)
    
    txtCapLegale.Text = nullToBlank(rsSocieta!CapSL)
    txtCodComuneLegale.Text = nullToBlank(rsSocieta!CodComuneSL)
    txtIndirizzoLegale.Text = nullToBlank(rsSocieta!IndirizzoSL)
    txtNrCivicoLegale.Text = nullToBlank(rsSocieta!NumCivicoSL)
    txtNrInternoLegale.Text = nullToBlank(rsSocieta!InternoSL)
    txtSubAlternoLegale.Text = nullToBlank(rsSocieta!SubalternoSL)
    txtPrefissoTelefonoLegale.Text = nullToBlank(rsSocieta!PrefissoTelSL)
    txtNumTelefonoLegale.Text = nullToBlank(rsSocieta!NumeroTelefonoSL)
    txtCognomeLegale.Text = nullToBlank(rsSocieta!CognomeLegaleR)
    txtNomeLegale.Text = nullToBlank(rsSocieta!NomeLegaleR)

    ctlComune.BoundText = nullToBlank(rsSocieta!codComune)
    ctlComuneLegale.BoundText = nullToBlank(rsSocieta!CodComuneSL)
    
    
    
End Sub

Private Sub txtNrCivicoLegale_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumCivico_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumeroFax_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumeroTelefono_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumTelefonoLegale_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPartitaIva_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

End Sub



'----------------------------------
'-- AUTORIZZAZIONI
'----------------------------------
Private Sub ChkAutoDefault_Click()
    Dim lvwItem As ListItem
    
    
    If ChkAutoDefault.Value = 1 Then
        For Each lvwItem In lvwAutorizzazioni.ListItems
            If lvwItem <> lvwAutorizzazioni.SelectedItem Then
                If lvwItem.Tag <> "DEL" Then
                    If lvwItem.ListSubItems("flgDefault") = "SI" Then
                        MsgBox "Attenzione esiste già una Autorizzazione impostata come default", vbExclamation + vbOK, App.Title
                        ChkAutoDefault.Value = 0
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
End Sub


Function getAutorizzazioneDiDefault() As Long
    Dim lvwItem As ListItem
    
    For Each lvwItem In lvwAutorizzazioni.ListItems
        If lvwItem.ListSubItems("flgDefault") = "SI" Then
            On Error Resume Next
            getAutorizzazioneDiDefault = lvwItem.ListSubItems("idAutorizzazione")
            Exit Function
        End If
    Next
    
    getAutorizzazioneDiDefault = -1
End Function


Private Function EsisteAssociazioneConImpianto(idAutorizzazione As Long, ragSocImpianto As String) As Boolean
    Dim lvwItem As ListItem
    
    For Each lvwItem In lvwImpianti.ListItems
        If lvwItem.ListSubItems("idAutorizzazione") = idAutorizzazione Then
            ragSocImpianto = lvwItem.ListSubItems("ragioneSociale")
            EsisteAssociazioneConImpianto = True
        End If
    Next
End Function


Private Sub cmdAutorizzazioni_Click(Index As Integer)
    Dim ragSocImpianto As String
    Select Case Index
        Case 3
            rsAutorizzazioni.AddNew

            rsAutorizzazioni!codSoggetto = rsSocieta!codSoggetto
            rsAutorizzazioni!idSoggetto = rsSocieta!idSocieta
            rsAutorizzazioni!datainiziovalidita = Now
            rsAutorizzazioni!datascadenza = Now
            rsAutorizzazioni!RilasciataDa = ""
            rsAutorizzazioni!numero = 0
            rsAutorizzazioni!flgDefault = False

            FillAutorizzazione
            
            Set lvwAutorizzazioni.SelectedItem = lvwAutorizzazioni.ListItems.Add(, , "")
            
            UpdateNewListAutorizzazioni lvwAutorizzazioni.SelectedItem, False
            
            lvwAutorizzazioni.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwAutorizzazioni.SelectedItem Is Nothing Then Exit Sub
            
            If EsisteAssociazioneConImpianto(rsAutorizzazioni!idAutorizzazione, ragSocImpianto) Then
                MsgBox "Attenzione l'autorizzazione è stata associata ad un impianto [" & ragSocImpianto & "],rimuovere tale associazione prima di continuare"
                Exit Sub
            End If
            
            lvwAutorizzazioni.ListItems.Remove (lvwAutorizzazioni.SelectedItem.Index)
                
            rsAutorizzazioni.Delete
                
            SalvaAutorizzazioni
                
            lvwAutorizzazioni_ItemClick lvwAutorizzazioni.SelectedItem
        Case 1
            If lvwAutorizzazioni.SelectedItem Is Nothing Or Not ValidaAutorizzazioni Then Exit Sub
            
            rsAutorizzazioni!datainiziovalidita = DTPAutoInizioValidita.Value
            rsAutorizzazioni!datascadenza = DTPAutoFineValidita.Value
            rsAutorizzazioni!RilasciataDa = txtAutoRilasciataDa
            rsAutorizzazioni!numero = txtAutoNumero
            rsAutorizzazioni!flgDefault = IIf(ChkAutoDefault.Value = 1, True, False)
            
            UpdateNewListAutorizzazioni lvwAutorizzazioni.SelectedItem, True
             
            lvwAutorizzazioni.SelectedItem.Tag = "MOD"
                        
            SalvaAutorizzazioni
            
            RefreshCommand
            
            flgModRecord = False
    End Select
    
End Sub



Private Sub lvwAutorizzazioni_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwAutorizzazioni, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwAutorizzazioni_ItemClick(ByVal item As MSComctlLib.ListItem)
    rsAutorizzazioni.MoveFirst
    rsAutorizzazioni.Move item.Index - 1
    
    FillAutorizzazione
End Sub

Private Sub caricaAutorizzazioni(idSocieta As Long, codSogg As String)
    
    Set rsAutorizzazioni = CreateObject("ADODB.Recordset")
    rsAutorizzazioni.LockType = adLockBatchOptimistic
    rsAutorizzazioni.CursorLocation = adUseClient
    rsAutorizzazioni.CursorType = adOpenStatic
    
    Set rsAutorizzazioni.ActiveConnection = gDBC
    
    rsAutorizzazioni.Source = "select * from Autorizzazioni where CodSoggetto='" & codSogg & "' and idSoggetto=" & idSocieta
    rsAutorizzazioni.Open
    
    If rsAutorizzazioni.EOF Then Exit Sub
    
    Set rsAutorizzazioni.ActiveConnection = Nothing
    
    FillListAutorizzazioni
    
End Sub

Private Sub FillAutorizzazione()
    txtAutoIdAutorizzazione.Text = nullToBlank(rsAutorizzazioni!idAutorizzazione)
    txtAutoCodSoggetto.Text = nullToBlank(rsAutorizzazioni!codSoggetto)
    txtAutoIDSoggetto.Text = nullToBlank(rsAutorizzazioni!idSoggetto)
    DTPAutoInizioValidita.Value = nullToBlank(rsAutorizzazioni!datainiziovalidita)
    DTPAutoFineValidita.Value = nullToBlank(rsAutorizzazioni!datascadenza)
    txtAutoRilasciataDa.Text = nullToBlank(rsAutorizzazioni!RilasciataDa)
    txtAutoNumero.Text = nullToBlank(rsAutorizzazioni!numero)
    ChkAutoDefault.Value = IIf(rsAutorizzazioni!flgDefault, 1, 0)
End Sub


Private Sub UpdateNewListAutorizzazioni(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsAutorizzazioni!idAutorizzazione) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "AutorizzazioneNew"
            ListItem.SmallIcon = "AutorizzazioneNew"
            
        ElseIf (Not IsNull(rsAutorizzazioni!idAutorizzazione) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Autorizzazione"
            ListItem.SmallIcon = "Autorizzazione"
        Else
            ListItem.Icon = "AutorizzazioneDel"
            ListItem.SmallIcon = "AutorizzazioneDel"
        End If
        
        ListItem.Text = rsAutorizzazioni!numero
        If Upd Then
            ListItem.ListSubItems("idAutorizzazione").Text = nullToBlank(rsAutorizzazioni!idAutorizzazione)
            ListItem.ListSubItems("idSoggetto").Text = rsAutorizzazioni!idSoggetto
            ListItem.ListSubItems("codSoggetto").Text = rsAutorizzazioni!codSoggetto
            ListItem.ListSubItems("Numero").Text = rsAutorizzazioni!numero
            ListItem.ListSubItems("DataInizioValidita").Text = rsAutorizzazioni!datainiziovalidita
            ListItem.ListSubItems("DataScadenza").Text = rsAutorizzazioni!datascadenza
            ListItem.ListSubItems("RilasciataDa").Text = rsAutorizzazioni!RilasciataDa
            ListItem.ListSubItems("flgDefault").Text = IIf(rsAutorizzazioni!flgDefault, "SI", "NO")
        Else
            ListItem.ListSubItems.Add , "idAutorizzazione", nullToBlank(rsAutorizzazioni!idAutorizzazione)
            ListItem.ListSubItems.Add , "idSoggetto", rsAutorizzazioni!idSoggetto
            ListItem.ListSubItems.Add , "codSoggetto", rsAutorizzazioni!codSoggetto
            ListItem.ListSubItems.Add , "Numero", rsAutorizzazioni!numero
            ListItem.ListSubItems.Add , "DataInizioValidita", rsAutorizzazioni!datainiziovalidita
            ListItem.ListSubItems.Add , "DataScadenza", rsAutorizzazioni!datascadenza
            ListItem.ListSubItems.Add , "RilasciataDa", rsAutorizzazioni!RilasciataDa
            ListItem.ListSubItems.Add , "flgDefault", IIf(rsAutorizzazioni!flgDefault, "SI", "NO")
        End If
        
        
        
End Sub

Private Sub FillListAutorizzazioni()
    Dim ListItem As ListItem
    
    lvwAutorizzazioni.ListItems.Clear
    
    If rsAutorizzazioni.RecordCount = 0 Then Exit Sub
    
    rsAutorizzazioni.MoveFirst
    
    Do While Not rsAutorizzazioni.EOF
        Set ListItem = lvwAutorizzazioni.ListItems.Add
        
        UpdateNewListAutorizzazioni ListItem, False
        
        rsAutorizzazioni.MoveNext
        
    Loop
    
    
End Sub

'----------------------------------
'-- IMPIANTI
'----------------------------------

Private Sub UpdateNewListImpianti(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsImpianti!idImpianto) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "ImpiantoNew"
            ListItem.SmallIcon = "ImpiantoNew"
        ElseIf (Not IsNull(rsImpianti!idImpianto) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Impianto"
            ListItem.SmallIcon = "Impianto"
        Else
            ListItem.Icon = "ImpiantoDel"
            ListItem.SmallIcon = "ImpiantoDel"
        End If
        
        ListItem.Text = nullToBlank(rsImpianti!RagioneSociale)
        
        If Upd Then
            ListItem.ListSubItems("idImpianto").Text = nullToBlank(rsImpianti!idAutorizzazione)
            ListItem.ListSubItems("idSoggetto").Text = nullToBlank(rsImpianti!idSoggetto)
            ListItem.ListSubItems("codSoggetto").Text = nullToBlank(rsImpianti!codSoggetto)
            ListItem.ListSubItems("idAutorizzazione").Text = nullToBlank(rsImpianti!idAutorizzazione)
            ListItem.ListSubItems("ragioneSociale").Text = nullToBlank(rsImpianti!RagioneSociale)
            ListItem.ListSubItems("CodFiscale").Text = nullToBlank(rsImpianti!CodFiscale)
            ListItem.ListSubItems("PartitaIva").Text = nullToBlank(rsImpianti!PartitaIva)
            ListItem.ListSubItems("cap").Text = nullToBlank(rsImpianti!Cap)
            ListItem.ListSubItems("CodComune").Text = nullToBlank(rsImpianti!codComune)
            ListItem.ListSubItems("Indirizzo").Text = nullToBlank(rsImpianti!Indirizzo)
            ListItem.ListSubItems("NumCivico").Text = nullToBlank(rsImpianti!NumCivico)
            ListItem.ListSubItems("Subalterno").Text = nullToBlank(rsImpianti!Subalterno)
            ListItem.ListSubItems("Interno").Text = nullToBlank(rsImpianti!Interno)
            ListItem.ListSubItems("NumeroTelefono").Text = nullToBlank(rsImpianti!PrefissoTel) & nullToBlank(rsImpianti!NumeroTelefono)
            ListItem.ListSubItems("NumeroFax").Text = nullToBlank(rsImpianti!PrefissoFax) & nullToBlank(rsImpianti!NumeroFax)
            ListItem.ListSubItems("CodIstat").Text = nullToBlank(rsImpianti!CodiceIstat)
            ListItem.ListSubItems("NumREA").Text = nullToBlank(rsImpianti!NumRea)
            ListItem.ListSubItems("TotAddetti").Text = nullToBlank(rsImpianti!TotAddetti)
            
        Else
            ListItem.ListSubItems.Add , "idImpianto", nullToBlank(rsImpianti!idAutorizzazione)
            ListItem.ListSubItems.Add , "idSoggetto", nullToBlank(rsImpianti!idSoggetto)
            ListItem.ListSubItems.Add , "codSoggetto", nullToBlank(rsImpianti!codSoggetto)
            ListItem.ListSubItems.Add , "idAutorizzazione", nullToBlank(rsImpianti!idAutorizzazione)
            ListItem.ListSubItems.Add , "ragioneSociale", nullToBlank(rsImpianti!RagioneSociale)
            ListItem.ListSubItems.Add , "CodFiscale", nullToBlank(rsImpianti!CodFiscale)
            ListItem.ListSubItems.Add , "PartitaIva", nullToBlank(rsImpianti!PartitaIva)
            ListItem.ListSubItems.Add , "cap", nullToBlank(rsImpianti!Cap)
            ListItem.ListSubItems.Add , "CodComune", nullToBlank(rsImpianti!codComune)
            ListItem.ListSubItems.Add , "Indirizzo", nullToBlank(rsImpianti!Indirizzo)
            ListItem.ListSubItems.Add , "NumCivico", nullToBlank(rsImpianti!NumCivico)
            ListItem.ListSubItems.Add , "Subalterno", nullToBlank(rsImpianti!Subalterno)
            ListItem.ListSubItems.Add , "Interno", nullToBlank(rsImpianti!Interno)
            ListItem.ListSubItems.Add , "NumeroTelefono", nullToBlank(rsImpianti!PrefissoTel) & nullToBlank(rsImpianti!NumeroTelefono)
            ListItem.ListSubItems.Add , "NumeroFax", nullToBlank(rsImpianti!PrefissoFax) & nullToBlank(rsImpianti!NumeroFax)
            ListItem.ListSubItems.Add , "CodIstat", nullToBlank(rsImpianti!CodiceIstat)
            ListItem.ListSubItems.Add , "NumREA", nullToBlank(rsImpianti!NumRea)
            ListItem.ListSubItems.Add , "TotAddetti", nullToBlank(rsImpianti!TotAddetti)
            
        End If
        
End Sub


Private Sub cmdImpianti_click(Index As Integer)
    
    Select Case Index
        Case 3
            rsImpianti.AddNew

            rsImpianti!codSoggetto = rsSocieta!codSoggetto
            rsImpianti!idSoggetto = rsSocieta!idSocieta
            rsImpianti!idAutorizzazione = getAutorizzazioneDiDefault
            rsImpianti!codComune = rsSocieta!codComune

            FillImpianti
            
            Set lvwImpianti.SelectedItem = lvwImpianti.ListItems.Add(, , "")
            
            UpdateNewListImpianti lvwImpianti.SelectedItem, False
            
            lvwImpianti.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwImpianti.SelectedItem Is Nothing Then Exit Sub
            
            lvwImpianti.ListItems.Remove (lvwImpianti.SelectedItem.Index)
                            
            rsImpianti.Delete
            
            SalvaImpianti
                            
            lvwImpianti_ItemClick lvwImpianti.SelectedItem

            
        Case 1
            If lvwImpianti.SelectedItem Is Nothing Or Not ValidaImpianti Then Exit Sub
            
            On Error GoTo Errore:
            rsImpianti!codSoggetto = txtImpCodSoggetto
            rsImpianti!idSoggetto = txtImpIdSoggetto
            rsImpianti!idAutorizzazione = txtImpIdAutorizzazione
            rsImpianti!RagioneSociale = txtImpRagioneSociale
            rsImpianti!CodFiscale = txtImpCodFiscale
            rsImpianti!PartitaIva = IIf(Val(txtImpPartitaIva.Text) = 0, Null, txtImpPartitaIva.Text)
            rsImpianti!Cap = txtImpCap
            rsImpianti!codComune = txtImpCodComune
            rsImpianti!Indirizzo = txtImpIndirizzo
            rsImpianti!NumCivico = toZero(txtImpNumeroCivico)
            rsImpianti!Subalterno = txtImpSubAlterno
            rsImpianti!Interno = txtImpInterno
            rsImpianti!NumeroTelefono = txtImpTelefono
            rsImpianti!NumeroFax = txtImpFax
            rsImpianti!NumRea = txtImpNrRea
            rsImpianti!CodiceIstat = txtImpIstatAttivita
            rsImpianti!TotAddetti = Val(txtImpNumAddetti)
            rsImpianti!PrefissoFax = txtImpPrefissoFax
            rsImpianti!PrefissoTel = txtImpPrefissoTel
            
            SalvaImpianti
            
            lvwImpianti.SelectedItem.Tag = "MOD"
            
            UpdateNewListImpianti lvwImpianti.SelectedItem, True
            
            flgModRecord = False
    End Select
    
    Exit Sub
Errore:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Salvataggio"
    
    
End Sub

Private Sub lvwImpianti_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwImpianti, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwImpianti_ItemClick(ByVal item As MSComctlLib.ListItem)
    rsImpianti.MoveFirst
    On Error Resume Next
    rsImpianti.Move item.Index - 1
    
    FillImpianti
End Sub



Private Sub caricaImpianti(idSocieta As Long, codSogg As String)
    
    Set rsImpianti = CreateObject("ADODB.Recordset")
    rsImpianti.LockType = adLockBatchOptimistic
    rsImpianti.CursorLocation = adUseClient
    rsImpianti.CursorType = adOpenStatic
    
    Set rsImpianti.ActiveConnection = gDBC
    
    
    rsImpianti.Source = "select * from Impianti where CodSoggetto='" & codSogg & "' and idSoggetto=" & idSocieta
    rsImpianti.Open
    
    If rsImpianti.EOF Then Exit Sub
    
    Set rsImpianti.ActiveConnection = Nothing
    
    FillListImpianti
    
End Sub

Private Sub FillImpianti()
    txtImpIdImpianto.Text = nullToBlank(rsImpianti!idImpianto)
    txtImpIdAutorizzazione.Text = nullToBlank(rsImpianti!idAutorizzazione)
    txtImpCodSoggetto.Text = nullToBlank(rsImpianti!codSoggetto)
    txtImpIdSoggetto.Text = nullToBlank(rsImpianti!idSoggetto)
    
    '
    txtImpCap.Text = nullToBlank(rsImpianti!Cap)
    txtImpCodComune.Text = nullToBlank(rsImpianti!codComune)
    txtImpCodFiscale.Text = nullToBlank(rsImpianti!CodFiscale)
    txtImpFax.Text = nullToBlank(rsImpianti!NumeroFax)
    txtImpIndirizzo.Text = nullToBlank(rsImpianti!Indirizzo)
    txtImpInterno.Text = nullToBlank(rsImpianti!Interno)
    txtImpNumeroCivico.Text = nullToBlank(rsImpianti!NumCivico)
    txtImpPartitaIva.Text = nullToBlank(rsImpianti!PartitaIva)
    txtImpRagioneSociale.Text = nullToBlank(rsImpianti!RagioneSociale)
    txtImpSubAlterno.Text = nullToBlank(rsImpianti!Subalterno)
    txtImpTelefono.Text = nullToBlank(rsImpianti!NumeroTelefono)
    txtImpIstatAttivita.Text = nullToBlank(rsImpianti!CodiceIstat)
    txtImpNumAddetti.Text = nullToBlank(rsImpianti!TotAddetti)
    txtImpPrefissoFax.Text = nullToBlank(rsImpianti!PrefissoFax)
    txtImpPrefissoTel.Text = nullToBlank(rsImpianti!PrefissoTel)
    txtImpNrRea.Text = nullToBlank(rsImpianti!NumRea)
    
    ctlImpComune.BoundText = nullToBlank(rsImpianti!codComune)
    ctlImpAutorizzazioni.BoundText = nullToBlank(rsImpianti!idAutorizzazione)
End Sub


Private Sub PicSocieta_Resize(Index As Integer)
    Dim i As Integer
    Dim stdPic As StdPicture
    Dim xCoeff As Double
       Select Case Index
        Case 0
            
            On Error Resume Next
            Set stdPic = LoadPicture(App.Path & "\Loghi\" & txtLogo.Text)
            
            ImageLogo.width = PicSocieta(0).ScaleWidth - ImageLogo.left - 150
            ImageLogo.Picture = stdPic
            
            ImageLogo.left = Frame1.left + Frame1.width + 150
    

        Case 1
            
            lvwAutorizzazioni.width = Me.ScaleWidth - 200
            Shape1.width = Me.ScaleWidth - 200
            
            AllargaLeColonneLVW lvwAutorizzazioni, False
        
            
            lvwAutorizzazioni.ColumnHeaders(1).width = 28 * 15
            lvwAutorizzazioni.ColumnHeaders(2).width = 0
            lvwAutorizzazioni.ColumnHeaders(3).width = 0
            lvwAutorizzazioni.ColumnHeaders(4).width = 0
            
            
        Case 2
            
            lvwImpianti.width = Me.ScaleWidth - 200
            shpImp.width = Me.ScaleWidth - 200
                    
            AllargaLeColonneLVW lvwImpianti, False
            
            lvwImpianti.ColumnHeaders(1).width = 28 * 15
            lvwImpianti.ColumnHeaders(2).width = 0
            lvwImpianti.ColumnHeaders(3).width = 0
            lvwImpianti.ColumnHeaders(4).width = 0
    End Select
        
End Sub

Private Sub FillListImpianti()
    Dim ListItem As ListItem
    
    lvwImpianti.ListItems.Clear
    
    If rsImpianti.RecordCount = 0 Then Exit Sub
    
    rsImpianti.MoveFirst
    
    Do While Not rsImpianti.EOF
        Set ListItem = lvwImpianti.ListItems.Add
        
        UpdateNewListImpianti ListItem, False
        
        rsImpianti.MoveNext
        
    Loop
    

End Sub




