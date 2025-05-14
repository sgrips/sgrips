VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImmobili 
   Caption         =   "Immobili dell'Utente"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   1
      Left            =   2040
      ScaleHeight     =   6795
      ScaleWidth      =   9285
      TabIndex        =   30
      Top             =   120
      Width           =   9285
      Begin VB.ComboBox Combo1 
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
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   3330
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
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
         Index           =   0
         Left            =   2520
         TabIndex        =   18
         Top             =   5115
         Width           =   1815
      End
      Begin VB.CheckBox chkFlgUtenzaDomesticaResidente 
         BackColor       =   &H00E0E0E0&
         DataField       =   "FlgUtenzaDomesticaResidente"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Height          =   285
         Left            =   2115
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6015
         Width           =   330
      End
      Begin VB.TextBox txtTipoUtenza 
         BackColor       =   &H00E0E0E0&
         DataField       =   "TipoUtenza"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3330
         Width           =   300
      End
      Begin VB.TextBox txtCodTariffa 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodTariffa"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2835
         Width           =   825
      End
      Begin VB.TextBox txtNucleoFamiliare 
         BackColor       =   &H00E0E0E0&
         DataField       =   "NucleoFamiliare"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2115
         TabIndex        =   9
         Top             =   3765
         Width           =   810
      End
      Begin VB.TextBox txtSuperficie 
         DataField       =   "Superficie"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2115
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2310
         Width           =   855
      End
      Begin VB.TextBox txtPiano 
         DataField       =   "Piano"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2940
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1860
         Width           =   660
      End
      Begin VB.TextBox txtScala 
         DataField       =   "Scala"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1860
         Width           =   660
      End
      Begin VB.TextBox txtNumInterno 
         DataField       =   "NumInterno"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1380
         Width           =   660
      End
      Begin VB.TextBox txtSubalterno 
         DataField       =   "Subalterno"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2940
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1380
         Width           =   660
      End
      Begin VB.TextBox txtNumCivico 
         DataField       =   "NumCivico"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2115
         TabIndex        =   3
         Top             =   1380
         Width           =   660
      End
      Begin VB.TextBox txtIndirizzo 
         DataField       =   "Indirizzo"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Height          =   510
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   2
         Top             =   750
         Width           =   6015
      End
      Begin VB.TextBox txtCodVia 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodVia"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4215
         Width           =   1020
      End
      Begin VB.TextBox txtCodComune 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodComune"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4665
         Width           =   990
      End
      Begin VB.TextBox txtCodImmobilePrincipale 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodImmobilePrincipale"
         DataMember      =   "cmdImmobiliPerUtente"
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
         Left            =   2115
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5610
         Width           =   855
      End
      Begin VB.TextBox txtTipo 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Tipo"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5115
         Width           =   300
      End
      Begin VB.TextBox txtCodUtente 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodUtente"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   3915
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   660
      End
      Begin VB.TextBox txtCodImmobile 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodImmobile"
         DataMember      =   "cmdImmobiliPerUtente"
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2115
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   660
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Bindings        =   "frmImmobili.frx":0000
         DataField       =   "DataInizioValidita"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2115
         TabIndex        =   21
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59703297
         CurrentDate     =   39613
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Bindings        =   "frmImmobili.frx":0049
         DataField       =   "DataFineValidita"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   22
         Top             =   6360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   59703297
         CurrentDate     =   39613
      End
      Begin MSDataListLib.DataCombo ctlComune 
         Bindings        =   "frmImmobili.frx":0090
         DataField       =   "CODISTAT"
         Height          =   360
         Left            =   3240
         TabIndex        =   16
         Top             =   4665
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   14737632
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
      Begin MSDataListLib.DataCombo ctlTariffa 
         Bindings        =   "frmImmobili.frx":00B4
         DataField       =   "codTariffa"
         Height          =   360
         Left            =   3000
         TabIndex        =   11
         Top             =   2835
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "Descrizione"
         BoundColumn     =   "codTariffa"
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
         Caption         =   "Nr.Civico/Sub/Int:"
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
         Left            =   570
         TabIndex        =   46
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Scala/Piano:"
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
         Left            =   960
         TabIndex        =   45
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Fine Validita:"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   6420
         Width           =   1605
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Inizio Validita:"
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
         Left            =   405
         TabIndex        =   43
         Top             =   6420
         Width           =   1680
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UD Residente:"
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
         Left            =   825
         TabIndex        =   42
         Top             =   6030
         Width           =   1260
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Utenza:"
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
         Left            =   990
         TabIndex        =   41
         Top             =   3375
         Width           =   1095
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipologia Immobile"
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
         Left            =   480
         TabIndex        =   40
         Top             =   2880
         Width           =   1605
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nucleo Familiare:"
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
         Left            =   570
         TabIndex        =   39
         Top             =   3810
         Width           =   1515
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Superficie:"
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
         Left            =   1170
         TabIndex        =   38
         Top             =   2355
         Width           =   915
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
         Index           =   6
         Left            =   1290
         TabIndex        =   37
         Top             =   795
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Via:"
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
         Left            =   1335
         TabIndex        =   36
         Top             =   4260
         Width           =   750
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
         Left            =   1305
         TabIndex        =   35
         Top             =   4710
         Width           =   780
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Imm.Principale:"
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
         Left            =   735
         TabIndex        =   34
         Top             =   5655
         Width           =   1350
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   1665
         TabIndex        =   33
         Top             =   5160
         Width           =   420
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Utente:"
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
         Left            =   3255
         TabIndex        =   32
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Immobile:"
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
         Left            =   1245
         TabIndex        =   31
         Top             =   315
         Width           =   840
      End
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
      Index           =   1
      Left            =   3960
      TabIndex        =   25
      Top             =   7680
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
      Index           =   2
      Left            =   2520
      TabIndex        =   24
      Top             =   7680
      Width           =   1215
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
      Index           =   3
      Left            =   1200
      TabIndex        =   23
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   360
      ScaleHeight     =   2535
      ScaleWidth      =   3495
      TabIndex        =   28
      Top             =   120
      Width           =   3495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "CodImmobile"
            Caption         =   "Cod Immobile"
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
            DataField       =   "CodUtente"
            Caption         =   "Cod Utente"
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
         BeginProperty Column03 
            DataField       =   "CodImmobilePrincipale"
            Caption         =   "CodImmobilePrincipale"
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
         BeginProperty Column05 
            DataField       =   "Subalterno"
            Caption         =   "Subalterno"
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
            DataField       =   "NumInterno"
            Caption         =   "Nr Interno"
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
            DataField       =   "Scala"
            Caption         =   "Scala"
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
            DataField       =   "Piano"
            Caption         =   "Piano"
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
            DataField       =   "CodComune"
            Caption         =   "Cod Comune"
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
            DataField       =   "CodVia"
            Caption         =   "Cod Via"
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
            DataField       =   "Tipo"
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
         BeginProperty Column12 
            DataField       =   "Superficie"
            Caption         =   "Superficie"
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
            DataField       =   "NucleoFamiliare"
            Caption         =   "NucleoFamiliare"
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
            DataField       =   "CodTariffa"
            Caption         =   "Cod Tariffa"
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
            DataField       =   "TipoUtenza"
            Caption         =   "Tipo Utenza"
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
            DataField       =   "FlgUtenzaDomesticaResidente"
            Caption         =   "FlgUtenzaDomesticaResidente"
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
            DataField       =   "DataInizioValidita"
            Caption         =   "Data Inizio Validita"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d MMM yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "DataFineValidita"
            Caption         =   "Data Fine Validita"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d MMM yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2700,284
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   2264,882
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
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
      Left            =   5280
      TabIndex        =   27
      Top             =   7680
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   360
      TabIndex        =   26
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Key             =   "LISTA"
            Object.ToolTipText     =   "Elenco degli Immobili"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dettaglio"
            Key             =   "DETTAGLIO"
            Object.ToolTipText     =   "Modifica Immobile"
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
Attribute VB_Name = "frmImmobili"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private sourceElenco As String

Private mcodUtente As Long

Private rsComuni As Recordset
Private rsTariffe As Recordset

Private pNoEvent As Boolean
Private pErrore As Boolean


Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    Dim rs As ADODB.Recordset
    Dim codComune As String
    
    
    Select Case Index
        Case 0
            cancelCurrentOperation
            Unload Me
        Case 1
                    
            pErrore = DevEnvMDB.rscmdImmobiliPerUtente.EditMode <> adEditNone
                    
            If DevEnvMDB.rscmdImmobiliPerUtente.EditMode <> adEditNone And validaDati Then
                DevEnvMDB.rscmdImmobiliPerUtente.UpdateBatch adAffectAllChapters
                DataGrid1.Refresh
                
                pErrore = False
                If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
            End If
            
        Case 2
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdImmobiliPerUtente.AddNew
            
            txtCodUtente.Text = mcodUtente
            txtCodImmobile.Text = getNumeratore(IMMOBILI, True, NUMERATORE_IMMOBILE_MIN)
            DevEnvMDB.rscmdImmobiliPerUtente("CodUtente").Value = mcodUtente
            DevEnvMDB.rscmdImmobiliPerUtente("CodImmobile").Value = txtCodImmobile.Text
            DevEnvMDB.rscmdImmobiliPerUtente("FlgUtenzaDomesticaResidente").Value = False
            
            
            Dim rsComune As ADODB.Recordset
            Set rsComune = GetCurrentComune
            If rsComune Is Nothing Then Exit Sub

            codComune = rsComune!codistat
               
            DevEnvMDB.rscmdImmobiliPerUtente("codComune").Value = codComune
            ctlComune.BoundText = codComune
            DTPicker1(0).Value = Format$(DateAdd("d", -1, Now), "yyyy\-mm\-dd")
            
            cmd(1).enabled = True
            cmd(3).enabled = True
            
            Combo1(0).ListIndex = 0
            txtNucleoFamiliare.Text = 1
            
            disableAllControl False
            
        Case 3
            If (MsgBox("Attenzione, se sicuro di voler cancellare l'immobile:" & DevEnvMDB.rscmdImmobiliPerUtente("Indirizzo"), vbQuestion + vbYesNo) = vbYes) Then
                DevEnvMDB.rscmdImmobiliPerUtente.Delete
                DevEnvMDB.rscmdImmobiliPerUtente.UpdateBatch adAffectCurrent
            End If
    End Select
End Sub

Private Sub disableAllControl(bBoolean As Boolean)
    
    txtCodImmobilePrincipale.Locked = True
    txtCodUtente.Locked = True
    txtCodComune.Locked = True
    txtCodTariffa.Locked = True
    txtCodVia.Locked = True
    
    txtCodImmobilePrincipale.TabStop = False
    txtCodUtente.TabStop = False
    txtCodComune.TabStop = False
    txtCodTariffa.TabStop = False
    txtCodVia.TabStop = False
    
    txtIndirizzo.Locked = bBoolean
    txtNumCivico.Locked = bBoolean
    txtNumInterno.Locked = bBoolean
    txtPiano.Locked = bBoolean
    txtScala.Locked = bBoolean
    txtSubalterno.Locked = bBoolean
    txtSuperficie.Locked = bBoolean
    
    Combo1(0).Locked = True
    Combo1(0).TabStop = False
    ctlComune.Locked = True
    ctlComune.TabStop = False
    txtNucleoFamiliare.Locked = True
    txtNucleoFamiliare.TabStop = False
    
    ctlTariffa.Locked = bBoolean
    Combo1(1).Locked = bBoolean
    
    chkFlgUtenzaDomesticaResidente.enabled = False
    chkFlgUtenzaDomesticaResidente.TabStop = False
    
    DTPicker1(0).enabled = Not bBoolean
    DTPicker1(0).TabStop = Not bBoolean
    DTPicker1(1).enabled = Not bBoolean
    DTPicker1(1).TabStop = Not bBoolean
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


Private Sub Combo1_Click(Index As Integer)
    Select Case Index
        Case 0
            txtTipo.Text = left$(Combo1(0).Text, 1)
        Case 1
            txtTipoUtenza.Text = left$(Combo1(1).Text, 1)
            If txtTipoUtenza.Text = "N" Then
                DevEnvMDB.rscmdImmobiliPerUtente("NucleoFamiliare").Value = 0
            Else
                DevEnvMDB.rscmdImmobiliPerUtente("NucleoFamiliare").Value = 1
            End If
    End Select
    
End Sub

Private Sub ctlComune_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        DoEvents
        txtCodComune.Text = ctlComune.BoundText
        txtCodComune.SetFocus
    End If
End Sub


Private Sub ctlTariffa_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        DoEvents
        txtCodTariffa.Text = ctlTariffa.BoundText
        txtCodTariffa.SetFocus
    End If
End Sub


Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs(2).Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then
        ctlComune.Text = getComune(nullToBlank(DevEnvMDB.rscmdImmobiliPerUtente.Fields("codcomune")))
        ctlTariffa.Text = getTariffa(nullToBlank(DevEnvMDB.rscmdImmobiliPerUtente.Fields("codtariffa")))
        
        Combo1(0).Text = nullToBlank(DevEnvMDB.rscmdImmobiliPerUtente.Fields("Tipo"))
        Combo1(1).Text = nullToBlank(DevEnvMDB.rscmdImmobiliPerUtente.Fields("TipoUtenza"))
        
        
        cmd(3).enabled = TabStrip1.SelectedItem.Index = 1 And IIf(DevEnvMDB.rscmdImmobiliPerUtente.Fields("CodImmobile").Value > NUMERATORE_IMMOBILE_MIN, True, False)
        cmd(2).enabled = True
        
    End If
End Sub

Function getComune(codistat As String) As String
    Dim rs As ADODB.Recordset
    
    getComune = vbNullString
    If codistat <> vbNullString Then
        Set rs = gDBC.Execute("select comune from comuni where codistat='" & codistat & "'")
        getComune = rs(0)
    End If
    

End Function

Function getTariffa(codTariffa As String) As String
    Dim rs As ADODB.Recordset
    
    getTariffa = vbNullString
    If codTariffa <> vbNullString Then
        Set rs = gDBC.Execute("select Descrizione from tariffe where codTariffa='" & codTariffa & "'")
        getTariffa = rs(0)
    End If
    
End Function


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
    
    sourceElenco = DevEnvMDB.rscmdImmobiliPerUtente.Source
        
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



    Set ctlTariffa.DataSource = Nothing
    Set ctlTariffa.RowSource = Nothing
    
    Set rsTariffe = CreateObject("ADODB.Recordset")
    rsTariffe.LockType = adLockBatchOptimistic
    rsTariffe.CursorLocation = adUseClient
    rsTariffe.CursorType = adOpenStatic
    
    Set rsTariffe.ActiveConnection = gDBC
    
    rsTariffe.Source = "Select * from Tariffe order by codTariffa"
    rsTariffe.Open
    
    Set rsTariffe.ActiveConnection = Nothing
    
    ctlTariffa.CausesValidation = False
    
    Set ctlTariffa.DataSource = rsTariffe
    Set ctlTariffa.RowSource = rsTariffe.DataSource

    Combo1(0).AddItem "P"
    Combo1(0).AddItem "Z"
    
    Combo1(1).AddItem "D Domestico"
    Combo1(1).AddItem "N Non Domestico"
    
    
    DataGrid1_RowColChange -1, -1
    

    Exit Sub
LoadERR:
    MsgBox "frmImmobili Load:" & Err.Description, vbCritical, App.Title
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
    On Error Resume Next
    DevEnvMDB.rscmdImmobiliPerUtente.Filter = vbNullString
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
    
    If (Not DevEnvMDB.rscmdImmobiliPerUtente.EOF And Not DevEnvMDB.rscmdImmobiliPerUtente.BOF) Then
        
        
        If DevEnvMDB.rscmdImmobiliPerUtente.EditMode = adEditAdd Then
            
            DevEnvMDB.rscmdImmobiliPerUtente.CancelBatch adAffectAllChapters
            If (Not DevEnvMDB.rscmdImmobiliPerUtente.EOF And Not DevEnvMDB.rscmdImmobiliPerUtente.BOF) Then
                If (DevEnvMDB.rscmdImmobiliPerUtente.state = adStateOpen) Then
                    DevEnvMDB.rscmdImmobiliPerUtente.Close
                End If
                
                DevEnvMDB.cmdImmobiliPerUtente mcodUtente
                DevEnvMDB.rscmdImmobiliPerUtente.MoveFirst
            End If
        ElseIf DevEnvMDB.rscmdImmobiliPerUtente.EditMode = adEditInProgress Then
            DevEnvMDB.rscmdImmobiliPerUtente.CancelUpdate
        End If
    End If
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If (Not DevEnvMDB.rscmdImmobiliPerUtente.EOF And Not DevEnvMDB.rscmdImmobiliPerUtente.BOF) Then

        If DevEnvMDB.rscmdImmobiliPerUtente.EditMode <> adEditNone Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click 2
                
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
    
    If Not (DevEnvMDB.rscmdImmobiliPerUtente.EOF And DevEnvMDB.rscmdImmobiliPerUtente.BOF) Then
    
        cmd(1).enabled = TabStrip1.SelectedItem.Index = 2 And IIf(DevEnvMDB.rscmdImmobiliPerUtente.Fields("CodImmobile").Value > NUMERATORE_IMMOBILE_MIN, True, False)
        cmd(1).default = TabStrip1.SelectedItem.Index = 2
        cmd(0).default = TabStrip1.SelectedItem.Index = 1
        cmd(3).enabled = TabStrip1.SelectedItem.Index = 2 And IIf(DevEnvMDB.rscmdImmobiliPerUtente.Fields("CodImmobile").Value > NUMERATORE_IMMOBILE_MIN, True, False)
    
        disableAllControl DevEnvMDB.rscmdImmobiliPerUtente.Fields("CodImmobile").Value < NUMERATORE_IMMOBILE_MIN
    
    End If
    
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1

    If (TabStrip1.SelectedItem.Index = 2) Then
        txtIndirizzo.SetFocus
    End If


End Sub

Private Sub ctlComune_GotFocus()
    cmd(1).default = False
End Sub

Private Sub ctlComune_LostFocus()
    cmd(1).default = True
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (5)
End Sub


Public Sub showMe(codUtente As Long, Des As String)

    mcodUtente = codUtente
    
    If (DevEnvMDB.rscmdImmobiliPerUtente.state = adStateOpen) Then
        DevEnvMDB.rscmdImmobiliPerUtente.Close
    End If
    

    DevEnvMDB.cmdImmobiliPerUtente codUtente

    If Not DevEnvMDB.rscmdImmobiliPerUtente.EOF Then
        Set Me.DataGrid1.DataSource = DevEnvMDB.rscmdImmobiliPerUtente
        Set DTPicker1(0).DataSource = DevEnvMDB.rscmdImmobiliPerUtente
        Set DTPicker1(1).DataSource = DevEnvMDB.rscmdImmobiliPerUtente
    End If
 

    Me.Caption = "Immobili dell'Utente " & Des
    
    Me.Show vbModal
End Sub


Private Sub Combo1_change(Index As Integer)
    'Combo1_Click Index
End Sub

Private Function validaDati() As Boolean
    Dim campoInErrore As Boolean
    
    validaDati = False
    campoInErrore = False
    
    txtIndirizzo.ToolTipText = ""
    txtIndirizzo.BackColor = -2147483643
    
    txtCodTariffa.ToolTipText = ""
    txtCodTariffa.BackColor = &HE0E0E0

    txtCodComune.ToolTipText = ""
    txtCodComune.BackColor = &HE0E0E0

    txtCodVia.ToolTipText = ""
    txtCodVia.BackColor = &HE0E0E0
    
    txtTipo.ToolTipText = ""
    txtTipo.BackColor = &HE0E0E0
    
    txtTipoUtenza.ToolTipText = ""
    txtTipoUtenza.BackColor = &HE0E0E0
    
    If txtCodComune.Text = "" Then
        txtCodComune.ToolTipText = "Attenzione Comune è obbligatorio"
        txtCodComune.BackColor = vbRed
        txtCodComune.SetFocus
        
        campoInErrore = True
    End If
    
    If txtTipo.Text = "" Then
        txtTipo.ToolTipText = "Attenzione Tipo è obbligatorio"
        txtTipo.BackColor = vbRed
        txtTipo.SetFocus
        
        campoInErrore = True
    End If
    
    If txtTipoUtenza.Text = "" Then
        txtTipoUtenza.ToolTipText = "Attenzione Tipo Utenza è obbligatorio"
        txtTipoUtenza.BackColor = vbRed
        txtTipoUtenza.SetFocus
        
        campoInErrore = True
    End If
    
    
    If txtIndirizzo.Text = "" Then
        txtIndirizzo.ToolTipText = "Attenzione l'indirizzo è obbligatorio"
        txtIndirizzo.BackColor = vbRed
        txtIndirizzo.SetFocus
        
        campoInErrore = True
    End If
    
    'If txtCodVia.Text = "" Then
    '    txtCodVia.ToolTipText = "Attenzione Cod Via è obbligatorio"
    '    txtCodVia.BackColor = vbRed
    '    txtCodVia.SetFocus
    '
    '    campoInErrore = True
    'End If
    
    If txtCodTariffa.Text = "" Then
        txtCodTariffa.ToolTipText = "Attenzione Cod Tariffa è obbligatorio"
        txtCodTariffa.BackColor = vbRed
        txtCodTariffa.SetFocus
        
        campoInErrore = True
    End If
    
    
    validaDati = Not campoInErrore

End Function

Private Sub txtNucleoFamiliare_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub

Private Sub txtSuperficie_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub

Private Sub txtNumCivico_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub
