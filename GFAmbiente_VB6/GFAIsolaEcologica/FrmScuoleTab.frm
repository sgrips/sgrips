VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmScuoleTab 
   Caption         =   "Gestione Società"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   16605
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2910
      Index           =   0
      Left            =   150
      ScaleHeight     =   2910
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   8865
      Begin VB.PictureBox picIstitutiDett 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   2880
         ScaleHeight     =   2055
         ScaleWidth      =   6255
         TabIndex        =   69
         Top             =   120
         Width           =   6255
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
            Left            =   45
            MaxLength       =   16
            TabIndex        =   80
            Top             =   1635
            Width           =   2820
         End
         Begin VB.TextBox txtImpIdIstituto 
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   225
            Width           =   1230
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
            Left            =   2955
            MaxLength       =   11
            TabIndex        =   78
            Top             =   1635
            Width           =   2745
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
            Left            =   45
            MaxLength       =   50
            TabIndex        =   77
            Top             =   930
            Width           =   5625
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   2340
            Width           =   990
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
            Left            =   45
            MaxLength       =   100
            TabIndex        =   75
            Top             =   3075
            Width           =   5460
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
            Left            =   930
            MaxLength       =   4
            TabIndex        =   74
            Top             =   3675
            Width           =   495
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   73
            Top             =   3675
            Width           =   495
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
            Left            =   30
            MaxLength       =   6
            TabIndex        =   72
            Top             =   3675
            Width           =   660
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
            Left            =   45
            MaxLength       =   15
            TabIndex        =   71
            Top             =   4410
            Width           =   5490
         End
         Begin VB.TextBox txtImpFax 
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
            Left            =   30
            MaxLength       =   15
            TabIndex        =   70
            Top             =   5085
            Width           =   5490
         End
         Begin MSDataListLib.DataCombo ctlComune 
            Bindings        =   "FrmScuoleTab.frx":0000
            DataField       =   "CODISTAT"
            Height          =   375
            Left            =   1230
            TabIndex        =   81
            Top             =   2340
            Width           =   4425
            _ExtentX        =   7805
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
            Left            =   45
            TabIndex        =   90
            Top             =   1395
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   20
            Left            =   45
            TabIndex        =   89
            Top             =   675
            Width           =   1410
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
            Left            =   2955
            TabIndex        =   88
            Top             =   1395
            Width           =   855
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
            Left            =   0
            TabIndex        =   87
            Top             =   2115
            Width           =   795
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
            Left            =   45
            TabIndex        =   86
            Top             =   2850
            Width           =   1170
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
            Left            =   45
            TabIndex        =   85
            Top             =   3450
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
            Index           =   27
            Left            =   45
            TabIndex        =   84
            Top             =   4125
            Width           =   780
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Istituto"
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
            Left            =   30
            TabIndex        =   83
            Top             =   0
            Width           =   585
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   0
            Left            =   45
            TabIndex        =   82
            Top             =   4800
            Width           =   390
         End
      End
      Begin MSComctlLib.ListView lvwIstituti 
         Height          =   1590
         Left            =   75
         TabIndex        =   1
         Top             =   75
         Width           =   2670
         _ExtentX        =   4710
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Istituto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   2
      Left            =   8160
      ScaleHeight     =   3255
      ScaleWidth      =   6345
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   6345
      Begin VB.PictureBox picClassiDett 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   3000
         ScaleHeight     =   3255
         ScaleWidth      =   4215
         TabIndex        =   33
         Top             =   120
         Width           =   4215
         Begin VB.TextBox txtClasseIdIstituto 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   870
            Width           =   990
         End
         Begin VB.TextBox txtClasseIdClasse 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   225
            Width           =   1230
         End
         Begin VB.TextBox txtClasseRagioneSociale 
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
            Left            =   15
            MaxLength       =   50
            TabIndex        =   35
            Top             =   2370
            Width           =   5625
         End
         Begin VB.TextBox txtClasseIdScuola 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1590
            Width           =   990
         End
         Begin MSDataListLib.DataCombo ctlClasseIstituto 
            Bindings        =   "FrmScuoleTab.frx":0024
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1275
            TabIndex        =   38
            Top             =   870
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
         Begin MSDataListLib.DataCombo ctlClasseScuola 
            Bindings        =   "FrmScuoleTab.frx":0059
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1275
            TabIndex        =   39
            Top             =   1590
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
            Caption         =   "Istituto"
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
            Left            =   0
            TabIndex        =   43
            Top             =   645
            Width           =   585
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   13
            Left            =   15
            TabIndex        =   42
            Top             =   2115
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Classe"
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
            Index           =   26
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Scuola"
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
            Left            =   45
            TabIndex        =   40
            Top             =   1365
            Width           =   600
         End
      End
      Begin MSComctlLib.ListView lvwClasse 
         Height          =   1590
         Left            =   75
         TabIndex        =   16
         Top             =   75
         Width           =   2790
         _ExtentX        =   4921
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Classe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Scuola"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Istituto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Index           =   3
      Left            =   8160
      ScaleHeight     =   3735
      ScaleWidth      =   7305
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   7305
      Begin VB.PictureBox picSezioniDett 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   3960
         ScaleHeight     =   4095
         ScaleWidth      =   3135
         TabIndex        =   19
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtSezioniIdScuola 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1590
            Width           =   990
         End
         Begin VB.TextBox txtSezioniRagioneSociale 
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
            Left            =   15
            MaxLength       =   50
            TabIndex        =   23
            Top             =   3090
            Width           =   5625
         End
         Begin VB.TextBox txtSezioniIdSezione 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   225
            Width           =   1230
         End
         Begin VB.TextBox txtSezioniidIstituto 
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
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   870
            Width           =   990
         End
         Begin VB.TextBox txtSezioniIdClasse 
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2310
            Width           =   990
         End
         Begin MSDataListLib.DataCombo ctlSezioniIstituto 
            Bindings        =   "FrmScuoleTab.frx":008E
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1275
            TabIndex        =   25
            Top             =   870
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
         Begin MSDataListLib.DataCombo ctlSezioniScuola 
            Bindings        =   "FrmScuoleTab.frx":00C3
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1275
            TabIndex        =   26
            Top             =   1590
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
         Begin MSDataListLib.DataCombo ctlSezioniClasse 
            Bindings        =   "FrmScuoleTab.frx":00F8
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1305
            TabIndex        =   27
            Top             =   2310
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
            Caption         =   "Scuola"
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
            Left            =   45
            TabIndex        =   32
            Top             =   1365
            Width           =   600
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Sezione"
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
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   705
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Descrizione"
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
            Left            =   15
            TabIndex        =   30
            Top             =   2835
            Width           =   1020
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Istituto"
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
            Left            =   0
            TabIndex        =   29
            Top             =   645
            Width           =   585
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Classe"
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
            Left            =   75
            TabIndex        =   28
            Top             =   2085
            Width           =   600
         End
      End
      Begin MSComctlLib.ListView lvwSezioni 
         Height          =   2070
         Left            =   75
         TabIndex        =   18
         Top             =   75
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   3651
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sezione"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Classe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Scuola"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Istituto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   7665
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   7665
      Begin VB.PictureBox picScuoleDett 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   3480
         ScaleHeight     =   6495
         ScaleWidth      =   3735
         TabIndex        =   44
         Top             =   120
         Width           =   3735
         Begin VB.TextBox txtScuolaFax 
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
            Left            =   45
            MaxLength       =   10
            TabIndex        =   67
            Top             =   5805
            Width           =   5490
         End
         Begin VB.TextBox txtScuolaTelefono 
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
            Left            =   45
            MaxLength       =   15
            TabIndex        =   55
            Top             =   5010
            Width           =   5490
         End
         Begin VB.TextBox txtScuolaNumeroCivico 
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
            Left            =   30
            MaxLength       =   6
            TabIndex        =   54
            Top             =   4275
            Width           =   660
         End
         Begin VB.TextBox txtScuolaSubAlterno 
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   53
            Top             =   4275
            Width           =   495
         End
         Begin VB.TextBox txtScuolaInterno 
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
            Left            =   930
            MaxLength       =   4
            TabIndex        =   52
            Top             =   4275
            Width           =   495
         End
         Begin VB.TextBox txtScuolaIndirizzo 
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
            Left            =   45
            MaxLength       =   100
            TabIndex        =   51
            Top             =   3675
            Width           =   5460
         End
         Begin VB.TextBox txtScuolaCodComune 
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   2940
            Width           =   990
         End
         Begin VB.TextBox txtScuolaRagioneSociale 
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
            Left            =   45
            MaxLength       =   50
            TabIndex        =   49
            Top             =   1530
            Width           =   5625
         End
         Begin VB.TextBox txtScuolaPartitaIva 
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
            Left            =   2955
            MaxLength       =   11
            TabIndex        =   48
            Top             =   2235
            Width           =   2745
         End
         Begin VB.TextBox txtScuolaID 
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   1230
         End
         Begin VB.TextBox txtScuolaCodFiscale 
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
            Left            =   45
            MaxLength       =   16
            TabIndex        =   46
            Top             =   2235
            Width           =   2820
         End
         Begin VB.TextBox txtScuolaidIstituto 
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
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   870
            Width           =   990
         End
         Begin MSDataListLib.DataCombo ctlScuolaComune 
            Bindings        =   "FrmScuoleTab.frx":012D
            CausesValidation=   0   'False
            DataField       =   "CODISTAT"
            Height          =   375
            Left            =   1230
            TabIndex        =   56
            Top             =   2925
            Width           =   4425
            _ExtentX        =   7805
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
         Begin MSDataListLib.DataCombo ctlScuolaIstituto 
            Bindings        =   "FrmScuoleTab.frx":0151
            CausesValidation=   0   'False
            DataField       =   "CodUtente"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1305
            TabIndex        =   57
            Top             =   870
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   661
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "RagioneSociale"
            BoundColumn     =   "CodUtente"
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
            Index           =   1
            Left            =   45
            TabIndex        =   68
            Top             =   5520
            Width           =   390
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Scuola"
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
            Left            =   30
            TabIndex        =   66
            Top             =   0
            Width           =   600
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
            Index           =   3
            Left            =   45
            TabIndex        =   65
            Top             =   4725
            Width           =   780
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
            Index           =   4
            Left            =   45
            TabIndex        =   64
            Top             =   4050
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
            Index           =   5
            Left            =   45
            TabIndex        =   63
            Top             =   3450
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
            Index           =   6
            Left            =   0
            TabIndex        =   62
            Top             =   2715
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
            Index           =   7
            Left            =   2955
            TabIndex        =   61
            Top             =   1995
            Width           =   855
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   8
            Left            =   45
            TabIndex        =   60
            Top             =   1275
            Width           =   1410
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
            Index           =   9
            Left            =   45
            TabIndex        =   59
            Top             =   1995
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Istituto"
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
            Left            =   30
            TabIndex        =   58
            Top             =   645
            Width           =   585
         End
      End
      Begin MSComctlLib.ListView lvwScuole 
         Height          =   1590
         Left            =   75
         TabIndex        =   14
         Top             =   75
         Width           =   3390
         _ExtentX        =   5980
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Scuola"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Istituto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picRight 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   11520
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   12
      Top             =   7560
      Width           =   4815
      Begin VB.CommandButton cmd 
         Caption         =   "Nuovo"
         DisabledPicture =   "FrmScuoleTab.frx":0186
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
         Picture         =   "FrmScuoleTab.frx":69D8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Inserisce un nuovo elemento"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Cancella"
         DisabledPicture =   "FrmScuoleTab.frx":D22A
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
         Picture         =   "FrmScuoleTab.frx":13A7C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "FrmScuoleTab.frx":1A2CE
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "FrmScuoleTab.frx":20B20
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Chiude"
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1560
      ScaleHeight     =   1215
      ScaleWidth      =   6420
      TabIndex        =   3
      Top             =   7680
      Width           =   6420
      Begin VB.CommandButton cmd 
         Caption         =   "Classi"
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
         Left            =   2400
         Picture         =   "FrmScuoleTab.frx":27372
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Alunni"
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
         Left            =   4800
         Picture         =   "FrmScuoleTab.frx":2D74C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Scuole"
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
         Index           =   7
         Left            =   1200
         Picture         =   "FrmScuoleTab.frx":33F9E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Istituti"
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
         Index           =   8
         Left            =   0
         Picture         =   "FrmScuoleTab.frx":3A378
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Sezioni"
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
         Left            =   3600
         Picture         =   "FrmScuoleTab.frx":40752
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9945
      Top             =   7905
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":46B2C
            Key             =   "Upd"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":4D38E
            Key             =   "Alunno"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":53BF0
            Key             =   "Classe"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":59FDA
            Key             =   "Scuola"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":603C4
            Key             =   "Sezione"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":667AE
            Key             =   "Istituto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":6CB98
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmScuoleTab.frx":733FA
            Key             =   "New"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmScuoleTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHdrIcons As New cLVHeaderSortIcons

Dim rsIstituti As Recordset
Dim rsComuni1 As Recordset

'Scuole
Dim rsScuole As Recordset
'Classe
Dim rsClasse As Recordset
'Sezioni
Dim rsSezioni As Recordset


Dim flgModRecord As Boolean
Dim LastIndexPage As Integer

Public Sub ShowMe()
    Load Me
    
    pic(0).Visible = True
    
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        Me.Show
    Else
        Me.Show
    End If
    
    cmd_Click 8
   
End Sub




'[10.35.10] Riccardo Sammito Ufficio : 1024x600
Private Sub cmd_Click(Index As Integer)
    
    If flgModRecord And ((Index >= 4 And Index <= 8) Or Index = 0) Then
        If MsgBox("Attenzione Esistono delle modifiche non salvate, continuando andranno perse,si vuole continuare?", vbQuestion + vbYesNo) = vbYes Then
            On Error Resume Next
            rsIstituti.CancelUpdate
            flgModRecord = False
        End If
    End If
    
    Select Case Index
        Case 0
            Unload Me
        Case 1, 2, 3
        
            cmd(3).enabled = Index <> 3
            cmd(1).enabled = True
            cmd(2).enabled = True

            If pic(0).Visible Then
                cmdIstituto_Click Index
                EnableDisableControl pic(0), Not rsIstituti.EOF
            ElseIf pic(1).Visible = True Then
                cmdScuola_Click Index
                EnableDisableControl pic(1), Not rsScuole.EOF
            ElseIf pic(2).Visible = True Then
                cmdClasse_Click Index
                EnableDisableControl pic(2), Not rsClasse.EOF
            ElseIf pic(3).Visible = True Then
                cmdSezioni_Click Index
                EnableDisableControl pic(3), Not rsSezioni.EOF
            End If
        Case 4, 5, 6, 7, 8
            pic(LastIndexPage).Visible = False
            pic(8 - Index).Visible = True
            
            Select Case Index
                Case 8
                    cmd(1).enabled = Not rsIstituti.EOF And Not rsIstituti.BOF
                    cmd(2).enabled = cmd(1).enabled
                Case 7
                    cmd(1).enabled = Not rsScuole.EOF And Not rsScuole.BOF
                    cmd(2).enabled = cmd(1).enabled
                Case 6
                    cmd(1).enabled = Not rsClasse.EOF And Not rsClasse.BOF
                    cmd(2).enabled = cmd(1).enabled
                Case 5
                    cmd(1).enabled = Not rsSezioni.EOF And Not rsSezioni.BOF
                    cmd(2).enabled = cmd(1).enabled
                Case 4

            End Select
            
            LastIndexPage = 8 - Index
    End Select
End Sub







Private Sub Form_Load()
    lvwIstituti.Icons = ImageList1
    lvwIstituti.SmallIcons = ImageList1
    
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
    
    Set ctlScuolaComune.DataSource = rsComuni1
    Set ctlScuolaComune.RowSource = rsComuni1
    
    
    ImageList1.ListImages.Add , "IstitutoNew", ImageList1.Overlay("Istituto", "New")
    ImageList1.ListImages.Add , "IstitutoUpd", ImageList1.Overlay("Istituto", "Upd")
    ImageList1.ListImages.Add , "IstitutoDel", ImageList1.Overlay("Istituto", "Del")
    
    'scuole
    lvwScuole.Icons = ImageList1
    lvwScuole.SmallIcons = ImageList1
    
    
    ImageList1.ListImages.Add , "ScuolaNew", ImageList1.Overlay("Scuola", "New")
    ImageList1.ListImages.Add , "ScuolaUpd", ImageList1.Overlay("Scuola", "Upd")
    ImageList1.ListImages.Add , "ScuolaDel", ImageList1.Overlay("Scuola", "Del")

    Set ctlScuolaComune.DataSource = Nothing
    Set ctlScuolaComune.RowSource = Nothing

    
    Set ctlScuolaComune.DataSource = rsComuni1.Clone
    Set ctlScuolaComune.RowSource = ctlScuolaComune.DataSource

    'classe
    lvwClasse.Icons = ImageList1
    lvwClasse.SmallIcons = ImageList1
    
    
    ImageList1.ListImages.Add , "ClasseNew", ImageList1.Overlay("Classe", "New")
    ImageList1.ListImages.Add , "ClasseUpd", ImageList1.Overlay("Classe", "Upd")
    ImageList1.ListImages.Add , "ClasseDel", ImageList1.Overlay("Classe", "Del")

    'Sezioni
    lvwSezioni.Icons = ImageList1
    lvwSezioni.SmallIcons = ImageList1
    
    
    ImageList1.ListImages.Add , "SezioneNew", ImageList1.Overlay("Sezione", "New")
    ImageList1.ListImages.Add , "SezioneUpd", ImageList1.Overlay("Sezione", "Upd")
    ImageList1.ListImages.Add , "SezioneDel", ImageList1.Overlay("Sezione", "Del")
    
    '
    caricaSezioni
    caricaClasse
    caricaScuole
    caricaIstituti
    
    cmd(8).Visible = gPuntiAssegnatiPer > 0
    cmd(7).Visible = gPuntiAssegnatiPer > 1
    cmd(6).Visible = gPuntiAssegnatiPer > 2
    cmd(5).Visible = gPuntiAssegnatiPer > 3
    cmd(4).Visible = gPuntiAssegnatiPer > 4
    
    cmd(7).enabled = False
    cmd(6).enabled = False
    cmd(5).enabled = False
    cmd(4).enabled = False

End Sub



Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    On Error Resume Next
    picLeft.Move 0, Me.ScaleHeight - picLeft.Height - 150
    picRight.Move Me.ScaleWidth - picRight.width - 150, picLeft.top
    For i = pic.LBound To pic.UBound
        pic(i).Move Me.ScaleLeft + 150, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight - picLeft.Height - 150
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


    rsIstituti.Close
    Set rsIstituti = Nothing

    rsComuni1.Close
    Set rsComuni1 = Nothing

    rsScuole.Close
    Set rsScuole = Nothing
    
    rsClasse.Close
    Set rsClasse = Nothing
    
    rsSezioni.Close
    Set rsSezioni = Nothing
    
    Set pfrmScuole = Nothing
End Sub


Private Sub pic_Resize(Index As Integer)
    Dim i As Integer
    Dim stdPic As StdPicture
    Dim xCoeff As Double
    Dim width As Long
    
       Select Case Index
        Case 0
            
            width = txtImpRagioneSociale.width + txtImpRagioneSociale.left + 150
            picIstitutiDett.Move pic(Index).ScaleWidth - width - 150, 0, width, pic(0).ScaleHeight - 150
            lvwIstituti.Move 0, 0, picIstitutiDett.left - 150, pic(0).ScaleHeight - 150
            
            AllargaLeColonneLVW lvwIstituti, False
            
            lvwIstituti.ColumnHeaders(1).width = 28 * 15
        
        Case 1
            
            width = txtScuolaRagioneSociale.width + txtScuolaRagioneSociale.left + 150
            picScuoleDett.Move pic(Index).ScaleWidth - width - 150, 0, width, pic(Index).ScaleHeight - 150
            lvwScuole.Move 0, 0, picScuoleDett.left - 150, pic(Index).ScaleHeight - 150
            
            lvwScuole.width = Me.ScaleWidth - 200
                    
            AllargaLeColonneLVW lvwScuole, False
            
            lvwScuole.ColumnHeaders(1).width = 28 * 15
        Case 2
            
            width = txtScuolaRagioneSociale.width + txtScuolaRagioneSociale.left + 150
            picClassiDett.Move pic(Index).ScaleWidth - width - 150, 0, width, pic(Index).ScaleHeight - 150
            lvwClasse.Move 0, 0, picClassiDett.left - 150, pic(Index).ScaleHeight - 150
                    
            AllargaLeColonneLVW lvwClasse, False
            
            lvwClasse.ColumnHeaders(1).width = 28 * 15
        Case 3
            
            width = txtSezioniRagioneSociale.width + txtSezioniRagioneSociale.left + 150
            picSezioniDett.Move pic(Index).ScaleWidth - width - 150, 0, width, pic(Index).ScaleHeight - 150
            lvwSezioni.Move 0, 0, picSezioniDett.left - 150, pic(Index).ScaleHeight - 150
                    
            AllargaLeColonneLVW lvwSezioni, False
            
            lvwSezioni.ColumnHeaders(1).width = 28 * 15
    
    End Select
        
End Sub

Private Function getParentDes(codFiglio As Long) As String
    Dim rs As ADODB.Recordset
    
    getParentDes = ""
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    
    
    rs.Source = "select RagioneSociale from Utenti where codUtente=" & codFiglio
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    
    
    If rs.EOF Then Exit Function
    
       
    getParentDes = rs!RagioneSociale
End Function

Private Function getParent(codFiglio As Long) As Long
    Dim rs As ADODB.Recordset
    
    getParent = 0
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    
    
    rs.Source = "select codPadre from Utenti where codUtente=" & codFiglio
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    
    
    If rs.EOF Then Exit Function
    
       
    getParent = rs!codPadre
End Function


'----------------------------------
'-- ISTITUTI
'----------------------------------

Private Sub ctlComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtImpCodComune.SetFocus
        txtImpCodComune.Text = ctlComune.BoundText
    End If
End Sub

Private Sub txtImpFax_KeyPress(KeyAscii As Integer)
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

Private Sub txtImpTelefono_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

End Sub





Private Sub UpdateNewListIstituti(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsIstituti!codUtente) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "IstitutoNew"
            ListItem.SmallIcon = "IstitutoNew"
        ElseIf (Not IsNull(rsIstituti!codUtente) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Istituto"
            ListItem.SmallIcon = "Istituto"
        Else
            ListItem.Icon = "IstitutoDel"
            ListItem.SmallIcon = "IstitutoDel"
        End If
        
        ListItem.Text = nullToBlank(rsIstituti!RagioneSociale)
        
        If Upd Then
            ListItem.ListSubItems("Istituto").Text = nullToBlank(rsIstituti!RagioneSociale)
            
        Else
            ListItem.ListSubItems.Add , "Istituto", nullToBlank(rsIstituti!RagioneSociale)
        End If
        
End Sub


Private Sub CancellaScuole(idPadre As Long)
    Dim deleted As Boolean
    

    rsScuole.Filter = ""
    rsScuole.Filter = "codPadre=" & idPadre
    
    Set rsScuole.ActiveConnection = gDBC
    Do While Not rsScuole.EOF
        CancellaClassi rsScuole!codUtente
        rsScuole.Delete
        rsScuole.MoveNext
        deleted = True
    Loop
    
    If deleted Then
        rsScuole.UpdateBatch adAffectAllChapters
        lvwScuole.ListItems.Clear
        ButtonDisableEnable
    End If
    
    Set rsScuole.ActiveConnection = Nothing
    rsScuole.Filter = ""
End Sub

Private Sub CancellaClassi(idPadre As Long)
    Dim deleted As Boolean
    
    rsClasse.Filter = ""
    rsClasse.Filter = "codPadre=" & idPadre
    
    Set rsClasse.ActiveConnection = gDBC
    Do While Not rsClasse.EOF
        CancellaSezioni rsClasse!codUtente
        rsClasse.Delete
        rsClasse.MoveNext
        deleted = True
    Loop
    
    
    If deleted Then
        lvwClasse.ListItems.Clear
        rsClasse.UpdateBatch adAffectAllChapters
        ButtonDisableEnable
    End If
    Set rsClasse.ActiveConnection = Nothing
    rsClasse.Filter = ""
End Sub

Private Sub CancellaSezioni(idPadre As Long)
    Dim deleted As Boolean
    
    rsSezioni.Filter = ""
    rsSezioni.Filter = "codPadre=" & idPadre
    Set rsSezioni.ActiveConnection = gDBC
    Do While Not rsSezioni.EOF
        rsSezioni.Delete
        rsSezioni.MoveNext
        deleted = True
    Loop
    
    
    If deleted Then
        lvwSezioni.ListItems.Clear
        rsSezioni.UpdateBatch adAffectAllChapters
        ButtonDisableEnable
    End If
    Set rsSezioni.ActiveConnection = Nothing
    rsSezioni.Filter = ""
    
End Sub



Private Sub cmdIstituto_Click(Index As Integer)
    
    Select Case Index
        Case 3
            rsIstituti.AddNew
            rsIstituti!TipoAnagrafe = "I"
            rsIstituti!codComuneResidenza.Value = GetCurrentComune!codistat

            FillIstituti
            
            Set lvwIstituti.SelectedItem = lvwIstituti.ListItems.Add(, , "")
            UpdateNewListIstituti lvwIstituti.SelectedItem, False
            lvwIstituti.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwIstituti.SelectedItem Is Nothing Then Exit Sub
            
            'cancella tutti i figli
            CancellaScuole rsIstituti!codUtente
             
            lvwIstituti.ListItems.Remove (lvwIstituti.SelectedItem.Index)
            rsIstituti.Delete
            
            SalvaIstituti
            
            If lvwIstituti.ListItems.Count > 0 Then lvwIstituti_ItemClick lvwIstituti.SelectedItem
        Case 1
            If lvwIstituti.SelectedItem Is Nothing Or Not ValidaIstituti Then Exit Sub
            
            rsIstituti!RagioneSociale = txtImpRagioneSociale
            rsIstituti!CodFiscale = txtImpCodFiscale
            rsIstituti!PartitaIva = IIf(Val(txtImpPartitaIva.Text) = 0, Null, txtImpPartitaIva.Text)
            rsIstituti!codComuneResidenza = txtImpCodComune
            rsIstituti!IndirizzoResidenza = txtImpIndirizzo
            rsIstituti!NumCivicoResidenza = toZero(txtImpNumeroCivico)
            rsIstituti!SubalternoResidenza = txtImpSubAlterno
            rsIstituti!InternoResidenza = txtImpInterno
            rsIstituti!NumeroTelefono = txtImpTelefono
            rsIstituti!NumeroFax = txtImpFax
            
            SalvaIstituti
            
            lvwIstituti.SelectedItem.Tag = "MOD"
            
            UpdateNewListIstituti lvwIstituti.SelectedItem, True
            
            FillIstituti
            
            flgModRecord = False
    End Select
    
End Sub

Private Sub lvwIstituti_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwIstituti, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwIstituti_ItemClick(ByVal Item As MSComctlLib.ListItem)
        
    rsIstituti.MoveFirst
    rsIstituti.Move Item.Index - 1
    
    FillIstituti
    
End Sub

Private Sub caricaIstituti()
    
    Set rsIstituti = CreateObject("ADODB.Recordset")
    rsIstituti.LockType = adLockBatchOptimistic
    rsIstituti.CursorLocation = adUseClient
    rsIstituti.CursorType = adOpenStatic
    
    Set rsIstituti.ActiveConnection = gDBC
    
    
    rsIstituti.Source = "select * from Utenti where TipoAnagrafe='I'"
    rsIstituti.Open
    Set rsIstituti.ActiveConnection = Nothing
    
    
    EnableDisableControl pic(0), Not rsIstituti.EOF
    
    If rsIstituti.EOF Then Exit Sub
    
    FillListIstituti
    
   
    Set ctlScuolaIstituto.DataSource = rsIstituti.Clone
    Set ctlScuolaIstituto.RowSource = ctlScuolaIstituto.DataSource
End Sub


Private Sub EnableDisableControl(pic As PictureBox, enable As Boolean)
    Dim c As Control
    Dim b As Long
    Dim b1 As Long
    b = &HE0E0E0
    For Each c In Me.Controls
        
        If (TypeOf c Is TextBox Or TypeOf c Is ComboBox Or TypeOf c Is DataCombo) Then
            On Error Resume Next
            b1 = c.BackColor
            If c.Container Is pic And b1 <> b Then
                c.enabled = enable
            End If
        End If
    Next
    
     
End Sub
    



Private Sub FillListIstituti()
    Dim ListItem As ListItem
    
    lvwIstituti.ListItems.Clear
    
    If rsIstituti.RecordCount = 0 Then Exit Sub
    
    Do While Not rsIstituti.EOF
        Set ListItem = lvwIstituti.ListItems.Add
        
        UpdateNewListIstituti ListItem, False
        
        rsIstituti.MoveNext
        
    Loop
    
    
    If lvwIstituti.ListItems.Count > 0 Then
        lvwIstituti.SelectedItem = lvwIstituti.ListItems(1)
        lvwIstituti_ItemClick lvwIstituti.SelectedItem
    Else
        FillIstituti
    End If

End Sub

Private Sub FillIstituti()
    txtImpIdIstituto.Text = vbNullString
    txtImpCodComune.Text = vbNullString
    txtImpFax.Text = vbNullString
    txtImpIndirizzo.Text = vbNullString
    txtImpInterno.Text = vbNullString
    txtImpNumeroCivico.Text = vbNullString
    txtImpPartitaIva.Text = vbNullString
    txtImpCodFiscale.Text = vbNullString
    txtImpRagioneSociale.Text = vbNullString
    txtImpSubAlterno.Text = vbNullString
    txtImpTelefono.Text = vbNullString

    
    If Not rsIstituti.EOF And Not rsIstituti.BOF Then
        txtImpIdIstituto.Text = nullToBlank(rsIstituti!codUtente)
        '
        txtImpCodComune.Text = nullToBlank(rsIstituti!codComuneResidenza)
        txtImpFax.Text = nullToBlank(rsIstituti!NumeroFax)
        txtImpIndirizzo.Text = nullToBlank(rsIstituti!IndirizzoResidenza)
        txtImpInterno.Text = nullToBlank(rsIstituti!InternoResidenza)
        txtImpNumeroCivico.Text = nullToBlank(rsIstituti!NumCivicoResidenza)
        txtImpPartitaIva.Text = nullToBlank(rsIstituti!PartitaIva)
        txtImpCodFiscale.Text = nullToBlank(rsIstituti!CodFiscale)
        txtImpRagioneSociale.Text = nullToBlank(rsIstituti!RagioneSociale)
        txtImpSubAlterno.Text = nullToBlank(rsIstituti!SubalternoResidenza)
        txtImpTelefono.Text = nullToBlank(rsIstituti!NumeroTelefono)
        
        ctlComune.BoundText = nullToBlank(rsIstituti!codComuneResidenza)
        'carica le scuole relative.....
    End If
    
    FillListScuole toZero(txtImpIdIstituto.Text)
    
    ButtonDisableEnable
End Sub

Private Sub ButtonDisableEnable()
    cmd(7).enabled = rsIstituti.RecordCount > 0
    cmd(6).enabled = rsScuole.RecordCount > 0 And rsIstituti.RecordCount > 0
    cmd(5).enabled = rsClasse.RecordCount > 0 And rsScuole.RecordCount > 0
    cmd(4).enabled = rsSezioni.RecordCount > 0 And rsClasse.RecordCount > 0 And rsScuole.RecordCount > 0
End Sub


Private Function ValidaIstituti() As Boolean
    
    ctlComune.BackColor = &H80000005
    ctlComune.ToolTipText = vbNullString
    
    txtImpIndirizzo.BackColor = &H80000005
    txtImpIndirizzo.ToolTipText = ""
    txtImpNumeroCivico.BackColor = &H80000005
    txtImpNumeroCivico.ToolTipText = ""
    txtImpTelefono.BackColor = &H80000005
    txtImpTelefono.ToolTipText = ""
    txtImpCodFiscale.BackColor = &H80000005
    txtImpCodFiscale.ToolTipText = ""
    txtImpPartitaIva.BackColor = &H80000005
    txtImpPartitaIva.ToolTipText = ""
    
    ValidaIstituti = True
    If txtImpCodFiscale.Text = vbNullString And txtImpPartitaIva.Text = vbNullString Then
        txtImpCodFiscale.BackColor = &H8080FF
        txtImpCodFiscale.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
        ValidaIstituti = False
        txtImpPartitaIva.BackColor = &H8080FF
        txtImpPartitaIva.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
    End If
    
    If txtImpCodComune = vbNullString Then
        ctlComune.BackColor = &H8080FF
        ctlComune.ToolTipText = "Il Campo è Obbligatorio"
        ValidaIstituti = False
    End If
    
    If txtImpIndirizzo = vbNullString Then
        txtImpIndirizzo.BackColor = &H8080FF
        txtImpIndirizzo.ToolTipText = "Il Campo è Obbligatorio"
        ValidaIstituti = False
    End If
    
    
    If txtImpNumeroCivico = vbNullString Then
        txtImpNumeroCivico.BackColor = &H8080FF
        txtImpNumeroCivico.ToolTipText = "Il Campo è Obbligatorio"
        ValidaIstituti = False
    End If
    
    If txtImpTelefono = vbNullString Then
        txtImpTelefono.BackColor = &H8080FF
        txtImpTelefono.ToolTipText = "Il Campo è Obbligatorio"
        ValidaIstituti = False
    End If
        
    
End Function


Private Sub SalvaIstituti()
    Dim position As Long
    
    On Error Resume Next
    position = rsIstituti.AbsolutePosition
    
    Set rsIstituti.ActiveConnection = gDBC
    rsIstituti.MoveFirst
    
    Do While Not rsIstituti.EOF
        If IsNull(rsIstituti!codUtente) Then rsIstituti!codUtente = getNumeratore(UTENTI, True)
        rsIstituti.MoveNext
    Loop
    
    
    rsIstituti.UpdateBatch adAffectAllChapters
    
    Set rsIstituti.ActiveConnection = Nothing

    If position > 0 Then rsIstituti.AbsolutePosition = position

End Sub


'----------------------------------
'-- SCUOLE
'----------------------------------

Private Sub ctlScuolaComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtScuolaCodComune.SetFocus
        txtScuolaCodComune.Text = ctlScuolaComune.BoundText
    End If
End Sub

Private Sub ctlScuolaIstituto_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtScuolaidIstituto.SetFocus
        txtScuolaidIstituto.Text = ctlScuolaIstituto.BoundText
    End If
End Sub

Private Sub txtScuolaFax_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtScuolaNumeroCivico_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtScuolaPartitaIva_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtScuolaTelefono_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If

End Sub



Private Sub UpdateNewListScuole(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsScuole!codUtente) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "ScuolaNew"
            ListItem.SmallIcon = "ScuolaNew"
        ElseIf (Not IsNull(rsScuole!codUtente) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Scuola"
            ListItem.SmallIcon = "Scuola"
        Else
            ListItem.Icon = "ScuolaDel"
            ListItem.SmallIcon = "ScuolaDel"
        End If
        
        ListItem.Text = nullToBlank(rsScuole!RagioneSociale)
        
        If Upd Then
            ListItem.ListSubItems("Scuola").Text = nullToBlank(rsScuole!codUtente)
            ListItem.ListSubItems("Istituto").Text = getParentDes(rsScuole!codPadre)
        Else
            ListItem.ListSubItems.Add , "Scuola", nullToBlank(rsScuole!codUtente)
            ListItem.ListSubItems.Add , "Istituto", getParentDes(rsScuole!codPadre)
        End If
        
End Sub


Private Sub cmdScuola_Click(Index As Integer)
    
    Select Case Index
        Case 3
            rsScuole.AddNew

            rsScuole!TipoAnagrafe = "S"
            rsScuole!codPadre = txtImpIdIstituto.Text
            rsScuole!codComuneResidenza.Value = txtImpCodComune.Text
            FillScuole
            
            Set lvwScuole.SelectedItem = lvwScuole.ListItems.Add(, , "")
            
            UpdateNewListScuole lvwScuole.SelectedItem, False
            
            lvwScuole.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwScuole.SelectedItem Is Nothing Then Exit Sub
            
            lvwScuole.ListItems.Remove (lvwScuole.SelectedItem.Index)
                            
            
            'cancella tutti i figli
            CancellaClassi rsScuole!codUtente
            
            rsScuole.Delete
            
            SalvaScuole
                          
            If lvwScuole.ListItems.Count > 0 Then lvwScuole_ItemClick lvwScuole.SelectedItem

            
        Case 1
            If lvwScuole.SelectedItem Is Nothing Or Not ValidaScuole Then Exit Sub
            
            rsScuole!RagioneSociale = txtScuolaRagioneSociale
            rsScuole!CodFiscale = txtScuolaCodFiscale
            rsScuole!PartitaIva = IIf(Val(txtScuolaPartitaIva.Text) = 0, Null, txtScuolaPartitaIva.Text)
            rsScuole!codComuneResidenza = txtScuolaCodComune
            rsScuole!codPadre = txtScuolaidIstituto
            rsScuole!IndirizzoResidenza = txtScuolaIndirizzo
            rsScuole!NumCivicoResidenza = toZero(txtScuolaNumeroCivico)
            rsScuole!SubalternoResidenza = txtScuolaSubAlterno
            rsScuole!InternoResidenza = txtScuolaInterno
            rsScuole!NumeroTelefono = txtScuolaTelefono
            rsScuole!NumeroFax = txtScuolaFax
            
            SalvaScuole
            
            lvwScuole.SelectedItem.Tag = "MOD"
            
            UpdateNewListScuole lvwScuole.SelectedItem, True
            
            FillScuole
            
            flgModRecord = False
    End Select
    
End Sub

Private Sub lvwScuole_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwScuole, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwScuole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsScuole.MoveFirst
    rsScuole.Move Item.Index - 1
    
    FillScuole
End Sub



Private Sub caricaScuole()
    
    Set rsScuole = CreateObject("ADODB.Recordset")
    rsScuole.LockType = adLockBatchOptimistic
    rsScuole.CursorLocation = adUseClient
    rsScuole.CursorType = adOpenStatic
    
    Set rsScuole.ActiveConnection = gDBC
    
    
    rsScuole.Source = "select * from Utenti where TipoAnagrafe='S'"
    rsScuole.Open
    Set rsScuole.ActiveConnection = Nothing
    

    
End Sub

Private Sub FillListScuole(idIstituto As Long)
    Dim ListItem As ListItem
    
    lvwScuole.ListItems.Clear
    
    If Not rsScuole Is Nothing Then
        rsScuole.Filter = ""
        rsScuole.Filter = "codPadre=" & idIstituto
    
        Do While Not rsScuole.EOF
            Set ListItem = lvwScuole.ListItems.Add
            
            UpdateNewListScuole ListItem, False
            
            rsScuole.MoveNext
        Loop
    End If
    
    Set ctlScuolaIstituto.DataSource = rsIstituti.Clone
    Set ctlScuolaIstituto.RowSource = ctlScuolaIstituto.DataSource

    If lvwScuole.ListItems.Count > 0 Then
        lvwScuole.SelectedItem = lvwScuole.ListItems(1)
        lvwScuole_ItemClick lvwScuole.SelectedItem
    Else
        FillScuole
    End If
        
    
    EnableDisableControl pic(1), Not rsScuole.EOF
End Sub


Private Sub FillScuole()


    txtScuolaID.Text = vbNullString
    txtScuolaidIstituto.Text = vbNullString
    '
    txtScuolaCodComune.Text = vbNullString
    txtScuolaFax.Text = vbNullString
    txtScuolaIndirizzo.Text = vbNullString
    txtScuolaInterno.Text = vbNullString
    txtScuolaNumeroCivico.Text = vbNullString
    txtScuolaPartitaIva.Text = vbNullString
    txtScuolaCodFiscale.Text = vbNullString
    txtScuolaRagioneSociale.Text = vbNullString
    txtScuolaSubAlterno.Text = vbNullString
    txtScuolaTelefono.Text = vbNullString
    
    If Not rsScuole.EOF And Not rsScuole.BOF Then
        txtScuolaID.Text = nullToBlank(rsScuole!codUtente)
        txtScuolaidIstituto.Text = nullToBlank(rsScuole!codPadre)
        '
        txtScuolaCodComune.Text = nullToBlank(rsScuole!codComuneResidenza)
        txtScuolaFax.Text = nullToBlank(rsScuole!NumeroFax)
        txtScuolaIndirizzo.Text = nullToBlank(rsScuole!IndirizzoResidenza)
        txtScuolaInterno.Text = nullToBlank(rsScuole!InternoResidenza)
        txtScuolaNumeroCivico.Text = nullToBlank(rsScuole!NumCivicoResidenza)
        txtScuolaPartitaIva.Text = nullToBlank(rsScuole!PartitaIva)
        txtScuolaCodFiscale.Text = nullToBlank(rsScuole!CodFiscale)
        txtScuolaRagioneSociale.Text = nullToBlank(rsScuole!RagioneSociale)
        txtScuolaSubAlterno.Text = nullToBlank(rsScuole!SubalternoResidenza)
        txtScuolaTelefono.Text = nullToBlank(rsScuole!NumeroTelefono)
        
        ctlScuolaComune.BoundText = nullToBlank(rsScuole!codComuneResidenza)
        ctlScuolaIstituto.BoundText = nullToBlank(rsScuole!codPadre)
    End If

    FillListClasse toZero(txtScuolaID.Text)

    ButtonDisableEnable
End Sub


Private Function ValidaScuole() As Boolean
    
    ctlScuolaComune.BackColor = &H80000005
    ctlScuolaComune.ToolTipText = vbNullString
    
    ctlScuolaIstituto.BackColor = &H80000005
    ctlScuolaIstituto.ToolTipText = vbNullString
    
    txtScuolaIndirizzo.BackColor = &H80000005
    txtScuolaIndirizzo.ToolTipText = ""
    txtScuolaNumeroCivico.BackColor = &H80000005
    txtScuolaNumeroCivico.ToolTipText = ""
    txtScuolaTelefono.BackColor = &H80000005
    txtScuolaTelefono.ToolTipText = ""
    txtScuolaCodFiscale.BackColor = &H80000005
    txtScuolaCodFiscale.ToolTipText = ""
    txtScuolaPartitaIva.BackColor = &H80000005
    txtScuolaPartitaIva.ToolTipText = ""
    
    ValidaScuole = True
    If txtScuolaCodFiscale.Text = vbNullString And txtScuolaPartitaIva.Text = vbNullString Then
        txtScuolaCodFiscale.BackColor = &H8080FF
        txtScuolaCodFiscale.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
        ValidaScuole = False
        txtScuolaPartitaIva.BackColor = &H8080FF
        txtScuolaPartitaIva.ToolTipText = "Inserire il Codice Fiscale o la P.Iva"
    End If
    
    If txtScuolaCodComune = vbNullString Then
        ctlScuolaComune.BackColor = &H8080FF
        ctlScuolaComune.ToolTipText = "Il Campo è Obbligatorio"
        ValidaScuole = False
    End If
    
    If txtScuolaidIstituto = vbNullString Then
        ctlScuolaIstituto.BackColor = &H8080FF
        ctlScuolaIstituto.ToolTipText = "Il Campo è Obbligatorio"
        ValidaScuole = False
    End If
    
    If txtScuolaIndirizzo = vbNullString Then
        txtScuolaIndirizzo.BackColor = &H8080FF
        txtScuolaIndirizzo.ToolTipText = "Il Campo è Obbligatorio"
        ValidaScuole = False
    End If
    
    
    If txtScuolaNumeroCivico = vbNullString Then
        txtScuolaNumeroCivico.BackColor = &H8080FF
        txtScuolaNumeroCivico.ToolTipText = "Il Campo è Obbligatorio"
        ValidaScuole = False
    End If
    
    If txtScuolaTelefono = vbNullString Then
        txtScuolaTelefono.BackColor = &H8080FF
        txtScuolaTelefono.ToolTipText = "Il Campo è Obbligatorio"
        ValidaScuole = False
    End If
        
    
End Function


Private Sub SalvaScuole()
    Dim position As Long
    
    On Error Resume Next
    position = rsScuole.AbsolutePosition
    
    Set rsScuole.ActiveConnection = gDBC
    rsScuole.MoveFirst
    
    Do While Not rsScuole.EOF
        If IsNull(rsScuole!codUtente) Then rsScuole!codUtente = getNumeratore(UTENTI, True)
        rsScuole.MoveNext
    Loop
    
    
    rsScuole.UpdateBatch adAffectAllChapters
    
    Set rsScuole.ActiveConnection = Nothing

    rsScuole.AbsolutePosition = position

End Sub

'----------------------------------
'-- CLASSE
'----------------------------------

Private Sub ctlClasseScuole_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtClasseIdScuola.SetFocus
        txtClasseIdScuola.Text = ctlClasseScuola.BoundText
    End If
End Sub

Private Sub ctlClasseIstituto_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtClasseIdIstituto.SetFocus
        txtClasseIdIstituto.Text = ctlClasseIstituto.BoundText
    End If
End Sub

Private Sub ctlClasseScuola_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtClasseIdScuola.SetFocus
        txtClasseIdScuola.Text = ctlClasseScuola.BoundText
    End If
End Sub

Private Sub UpdateNewListClasse(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsClasse!codUtente) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "ClasseNew"
            ListItem.SmallIcon = "ClasseNew"
        ElseIf (Not IsNull(rsClasse!codUtente) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Classe"
            ListItem.SmallIcon = "Classe"
        Else
            ListItem.Icon = "ClasseDel"
            ListItem.SmallIcon = "ClasseDel"
        End If
        
        ListItem.Text = nullToBlank(rsClasse!RagioneSociale)
        
        If Upd Then
            ListItem.ListSubItems("Classe").Text = nullToBlank(rsClasse!RagioneSociale)
            ListItem.ListSubItems("Scuola").Text = getParentDes(rsClasse!codPadre)
            ListItem.ListSubItems("Istituto").Text = getParentDes(getParent(rsClasse!codPadre))
        
        Else
            ListItem.ListSubItems.Add , "Classe", nullToBlank(rsClasse!RagioneSociale)
            ListItem.ListSubItems.Add , "Scuola", getParentDes(rsClasse!codPadre)
            ListItem.ListSubItems.Add , "Istituto", getParentDes(getParent(rsClasse!codPadre))
        End If
        
End Sub


Private Sub cmdClasse_Click(Index As Integer)
    
    Select Case Index
        Case 3
            rsClasse.AddNew

            rsClasse!TipoAnagrafe = "C"
            rsClasse!codPadre = txtScuolaID.Text

            FillClasse
            
            Set lvwClasse.SelectedItem = lvwClasse.ListItems.Add(, , "")
            
            UpdateNewListClasse lvwClasse.SelectedItem, False
            
            lvwClasse.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwClasse.SelectedItem Is Nothing Then Exit Sub
            
            lvwClasse.ListItems.Remove (lvwClasse.SelectedItem.Index)
                            
            'cancella tutti i figli
            CancellaSezioni rsClasse!codUtente
                            
            rsClasse.Delete
            
            SalvaClasse
                            
                             
            If lvwClasse.ListItems.Count > 0 Then lvwClasse_ItemClick lvwClasse.SelectedItem

            
        Case 1
            If lvwClasse.SelectedItem Is Nothing Or Not ValidaClasse Then Exit Sub
            
            rsClasse!RagioneSociale = txtClasseRagioneSociale
            
            SalvaClasse
            
            lvwClasse.SelectedItem.Tag = "MOD"
            
            UpdateNewListClasse lvwClasse.SelectedItem, True
                        
            FillClasse
            
            flgModRecord = False
    End Select
    
End Sub

Private Sub lvwClasse_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwClasse, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwClasse_ItemClick(ByVal Item As MSComctlLib.ListItem)
       
    rsClasse.MoveFirst
    rsClasse.Move Item.Index - 1
    
    FillClasse
End Sub



Private Sub caricaClasse()
    
    Set rsClasse = CreateObject("ADODB.Recordset")
    rsClasse.LockType = adLockBatchOptimistic
    rsClasse.CursorLocation = adUseClient
    rsClasse.CursorType = adOpenStatic
    
    Set rsClasse.ActiveConnection = gDBC
    
    
    rsClasse.Source = "select * from Utenti where TipoAnagrafe='C'"
    rsClasse.Open
    Set rsClasse.ActiveConnection = Nothing
    
    
    
    
    
End Sub

Private Sub FillListClasse(idScuole As Long)
    Dim ListItem As ListItem
    
    lvwClasse.ListItems.Clear
    
    rsClasse.Filter = ""
    rsClasse.Filter = "codPadre=" & idScuole
    
    Do While Not rsClasse.EOF
        Set ListItem = lvwClasse.ListItems.Add
        
        UpdateNewListClasse ListItem, False
        
        rsClasse.MoveNext
        
    Loop
    
    Set ctlClasseIstituto.DataSource = rsIstituti.Clone
    Set ctlClasseIstituto.RowSource = ctlScuolaIstituto.DataSource
    
    Set ctlClasseScuola.DataSource = rsScuole.Clone
    Set ctlClasseScuola.RowSource = ctlClasseScuola.DataSource
    

    If lvwClasse.ListItems.Count > 0 Then
        lvwClasse.SelectedItem = lvwClasse.ListItems(1)
        lvwClasse_ItemClick lvwClasse.SelectedItem
    Else
        FillClasse
    End If
    
    EnableDisableControl pic(2), Not rsClasse.EOF
End Sub


Private Sub FillClasse()

    txtClasseIdClasse.Text = vbNullString
    txtClasseIdIstituto.Text = vbNullString
    txtClasseIdScuola.Text = vbNullString
   
    txtClasseRagioneSociale.Text = vbNullString
    
    If Not rsClasse.EOF And Not rsClasse.BOF Then
        txtClasseIdClasse.Text = nullToBlank(rsClasse!codUtente)
        txtClasseIdIstituto.Text = getParent(toZero(rsClasse!codPadre))
        txtClasseIdScuola.Text = nullToBlank(rsClasse!codPadre)
        
        txtClasseRagioneSociale.Text = nullToBlank(rsClasse!RagioneSociale)
        ctlClasseIstituto.BoundText = nullToBlank(txtClasseIdIstituto.Text)
        ctlClasseScuola.BoundText = nullToBlank(txtClasseIdScuola.Text)
    End If
    
    FillListSezioni toZero(txtClasseIdClasse.Text)
    
    ButtonDisableEnable
End Sub


Private Function ValidaClasse() As Boolean
    
    ctlClasseScuola.BackColor = &H80000005
    ctlClasseScuola.ToolTipText = vbNullString
    
    ctlClasseIstituto.BackColor = &H80000005
    ctlClasseIstituto.ToolTipText = vbNullString
    
    txtClasseRagioneSociale.BackColor = &H80000005
    txtClasseRagioneSociale.ToolTipText = ""
    
    ValidaClasse = True
    
    If txtClasseRagioneSociale = vbNullString Then
        txtClasseRagioneSociale.BackColor = &H8080FF
        txtClasseRagioneSociale.ToolTipText = "Il Campo è Obbligatorio"
        ValidaClasse = False
    End If
    
    If txtClasseIdIstituto = vbNullString Then
        ctlScuolaIstituto.BackColor = &H8080FF
        ctlScuolaIstituto.ToolTipText = "Il Campo è Obbligatorio"
        ValidaClasse = False
    End If
    
    If txtClasseIdScuola = vbNullString Then
        ctlClasseScuola.BackColor = &H8080FF
        ctlClasseScuola.ToolTipText = "Il Campo è Obbligatorio"
        ValidaClasse = False
    End If
    
End Function


Private Sub SalvaClasse()
    Dim position As Long
    
    On Error Resume Next
    position = rsClasse.AbsolutePosition
    
    Set rsClasse.ActiveConnection = gDBC
    rsClasse.MoveFirst
    
    Do While Not rsClasse.EOF
        If IsNull(rsClasse!codUtente) Then rsClasse!codUtente = getNumeratore(UTENTI, True)
        rsClasse.MoveNext
    Loop
    
    
    rsClasse.UpdateBatch adAffectAllChapters
    
    Set rsClasse.ActiveConnection = Nothing

    rsClasse.AbsolutePosition = position

End Sub


'----------------------------------
'-- SEZIONI
'----------------------------------

Private Sub ctlSezioniScuole_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtSezioniIdScuola.SetFocus
        txtSezioniIdScuola.Text = ctlSezioniScuola.BoundText
    End If
End Sub

Private Sub ctlSezioniIstituto_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtSezioniidIstituto.SetFocus
        txtSezioniidIstituto.Text = ctlSezioniIstituto.BoundText
    End If
End Sub

Private Sub ctlSezioniClasse_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtSezioniIdClasse.SetFocus
        txtSezioniIdClasse.Text = ctlSezioniClasse.BoundText
    End If
End Sub

Private Sub UpdateNewListSezioni(ListItem As ListItem, Upd As Boolean)

        If IsNull(rsSezioni!codUtente) And ListItem.Tag <> "DEL" Then
            ListItem.Icon = "SezioneNew"
            ListItem.SmallIcon = "SezioneNew"
        ElseIf (Not IsNull(rsSezioni!codUtente) And ListItem.Tag <> "DEL") Or ListItem.Tag <> "MOD" Then
            ListItem.Icon = "Sezione"
            ListItem.SmallIcon = "Sezione"
        Else
            ListItem.Icon = "SezioneDel"
            ListItem.SmallIcon = "SezioneDel"
        End If
        
        ListItem.Text = nullToBlank(rsSezioni!RagioneSociale)
        Dim P As Long
        If Upd Then
            ListItem.ListSubItems("Sezione").Text = nullToBlank(rsSezioni!RagioneSociale)
            ListItem.ListSubItems("Classe").Text = getParentDes(rsSezioni!codPadre)
            ListItem.ListSubItems("Scuola").Text = getParentDes(getParent(rsSezioni!codPadre))
            
            P = getParent(rsSezioni!codPadre)
            ListItem.ListSubItems("Istituto").Text = getParentDes(getParent(P))
        Else
            ListItem.ListSubItems.Add , "Sezione", nullToBlank(rsSezioni!RagioneSociale)
            ListItem.ListSubItems.Add , "Classe", getParentDes(rsSezioni!codPadre)
            P = getParent(rsSezioni!codPadre)
            ListItem.ListSubItems.Add , "Scuola", getParentDes(P)
            ListItem.ListSubItems.Add , "Istituto", getParentDes(getParent(P))
        End If
        
End Sub


Private Sub cmdSezioni_Click(Index As Integer)
    
    Select Case Index
        Case 3
            rsSezioni.AddNew

            rsSezioni!TipoAnagrafe = "Z"
            rsSezioni!codPadre = txtClasseIdClasse.Text
            

            FillSezioni
            
            Set lvwSezioni.SelectedItem = lvwSezioni.ListItems.Add(, , "")
            
            UpdateNewListSezioni lvwSezioni.SelectedItem, False
            
            lvwSezioni.SelectedItem.Tag = "INS"
            
            flgModRecord = True
        Case 2
            If lvwSezioni.SelectedItem Is Nothing Then Exit Sub
            
            lvwSezioni.ListItems.Remove (lvwSezioni.SelectedItem.Index)
                            
            rsSezioni.Delete
            
            SalvaSezioni
                            
            If lvwSezioni.ListItems.Count > 0 Then lvwSezioni_ItemClick lvwSezioni.SelectedItem

            
        Case 1
            If lvwSezioni.SelectedItem Is Nothing Or Not ValidaSezioni Then Exit Sub
            
            rsSezioni!RagioneSociale = txtSezioniRagioneSociale
            
            SalvaSezioni
            
            lvwSezioni.SelectedItem.Tag = "MOD"
            
            UpdateNewListSezioni lvwSezioni.SelectedItem, True
            
            FillSezioni
            
            flgModRecord = False
    End Select
    
End Sub

Private Sub lvwSezioni_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwSezioni, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwSezioni_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsSezioni.MoveFirst
    rsSezioni.Move Item.Index - 1
    
    FillSezioni
End Sub



Private Sub caricaSezioni()
    
    Set rsSezioni = CreateObject("ADODB.Recordset")
    rsSezioni.LockType = adLockBatchOptimistic
    rsSezioni.CursorLocation = adUseClient
    rsSezioni.CursorType = adOpenStatic
    
    Set rsSezioni.ActiveConnection = gDBC
    
    
    rsSezioni.Source = "select * from Utenti where TipoAnagrafe='Z'"
    rsSezioni.Open
    Set rsSezioni.ActiveConnection = Nothing
    

    

    
End Sub

Private Sub FillListSezioni(idClasse As Long)
    Dim ListItem As ListItem
    
    lvwSezioni.ListItems.Clear
    
    rsSezioni.Filter = ""
    rsSezioni.Filter = "codPadre=" & idClasse
    
    Do While Not rsSezioni.EOF
        Set ListItem = lvwSezioni.ListItems.Add
        
        UpdateNewListSezioni ListItem, False
        
        rsSezioni.MoveNext
        
    Loop
    
    Set ctlSezioniIstituto.DataSource = rsIstituti.Clone
    Set ctlSezioniIstituto.RowSource = ctlScuolaIstituto.DataSource
    
    Set ctlSezioniScuola.DataSource = rsScuole.Clone
    Set ctlSezioniScuola.RowSource = ctlSezioniScuola.DataSource
    
    Set ctlSezioniClasse.DataSource = rsClasse.Clone
    Set ctlSezioniClasse.RowSource = ctlSezioniClasse.DataSource


    If lvwSezioni.ListItems.Count > 0 Then
        lvwSezioni.SelectedItem = lvwSezioni.ListItems(1)
        lvwSezioni_ItemClick lvwSezioni.SelectedItem
    Else
        FillSezioni
    End If


    EnableDisableControl pic(3), Not rsSezioni.EOF
End Sub


Private Sub FillSezioni()

    txtSezioniIdSezione.Text = vbNullString
    txtSezioniIdClasse.Text = vbNullString
    txtSezioniIdScuola.Text = vbNullString
    txtSezioniidIstituto.Text = vbNullString
    txtSezioniRagioneSociale.Text = vbNullString


    If Not rsSezioni.EOF And Not rsSezioni.BOF Then
        txtSezioniIdSezione.Text = nullToBlank(rsSezioni!codUtente)
        txtSezioniIdClasse.Text = nullToBlank(rsSezioni!codPadre)
        txtSezioniIdScuola.Text = getParent(toZero(txtSezioniIdClasse.Text))
        txtSezioniidIstituto.Text = getParent(toZero(txtSezioniIdScuola.Text))
        
        txtSezioniRagioneSociale.Text = nullToBlank(rsSezioni!RagioneSociale)
    End If
    
    
    ctlSezioniIstituto.BoundText = nullToBlank(txtSezioniidIstituto.Text)
    ctlSezioniScuola.BoundText = nullToBlank(txtSezioniIdScuola.Text)
    ctlSezioniClasse.BoundText = nullToBlank(txtSezioniIdClasse.Text)
End Sub


Private Function ValidaSezioni() As Boolean
    
    ctlSezioniScuola.BackColor = &H80000005
    ctlSezioniScuola.ToolTipText = vbNullString
    
    ctlSezioniIstituto.BackColor = &H80000005
    ctlSezioniIstituto.ToolTipText = vbNullString
    
    ctlSezioniClasse.BackColor = &H80000005
    ctlSezioniClasse.ToolTipText = vbNullString
    
    txtSezioniRagioneSociale.BackColor = &H80000005
    txtSezioniRagioneSociale.ToolTipText = ""
    
    ValidaSezioni = True
    
    If txtSezioniRagioneSociale = vbNullString Then
        txtSezioniRagioneSociale.BackColor = &H8080FF
        txtSezioniRagioneSociale.ToolTipText = "Il Campo è Obbligatorio"
        ValidaSezioni = False
    End If
    
    If txtSezioniidIstituto = vbNullString Then
        ctlScuolaIstituto.BackColor = &H8080FF
        ctlScuolaIstituto.ToolTipText = "Il Campo è Obbligatorio"
        ValidaSezioni = False
    End If
    
    If txtSezioniIdScuola = vbNullString Then
        ctlSezioniScuola.BackColor = &H8080FF
        ctlSezioniScuola.ToolTipText = "Il Campo è Obbligatorio"
        ValidaSezioni = False
    End If
    
    If txtSezioniIdClasse = vbNullString Then
        ctlSezioniClasse.BackColor = &H8080FF
        ctlSezioniClasse.ToolTipText = "Il Campo è Obbligatorio"
        ValidaSezioni = False
    End If
End Function


Private Sub SalvaSezioni()
    Dim position As Long
    
    On Error Resume Next
    position = rsSezioni.AbsolutePosition
    
    Set rsSezioni.ActiveConnection = gDBC
    rsSezioni.MoveFirst
    
    Do While Not rsSezioni.EOF
        If IsNull(rsSezioni!codUtente) Then rsSezioni!codUtente = getNumeratore(UTENTI, True)
        rsSezioni.MoveNext
    Loop
    
    
    rsSezioni.UpdateBatch adAffectAllChapters
    
    Set rsSezioni.ActiveConnection = Nothing

    rsSezioni.AbsolutePosition = position

End Sub



