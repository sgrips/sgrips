VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpostazioni 
   Caption         =   "Impostazioni Programma"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   14250
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   6
      Left            =   6720
      ScaleHeight     =   6015
      ScaleWidth      =   7695
      TabIndex        =   137
      Top             =   4920
      Width           =   7695
      Begin VB.Frame Frame1 
         Caption         =   "Accessi Isola"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5415
         Index           =   11
         Left            =   3480
         TabIndex        =   138
         Top             =   0
         Width           =   5895
         Begin VB.TextBox txtSecRefreshAccessMonitor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            MaxLength       =   2
            TabIndex        =   147
            Text            =   "0"
            Top             =   3480
            Width           =   2895
         End
         Begin VB.TextBox txtMaxMinInIsola 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            MaxLength       =   3
            TabIndex        =   145
            Text            =   "0"
            ToolTipText     =   "Indicare il tempo medio di permanenza nell'isola oltre al quale l'utente viene marcato come uscito."
            Top             =   2760
            Width           =   2895
         End
         Begin VB.ComboBox cmb 
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
            Index           =   7
            ItemData        =   "frmImpostazioni.frx":0000
            Left            =   120
            List            =   "frmImpostazioni.frx":0019
            TabIndex        =   143
            Text            =   "COM1"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.CheckBox chkActiveAccess 
            Caption         =   "Attiva Gestione Accessi"
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
            Left            =   120
            TabIndex        =   142
            ToolTipText     =   "Se Attiva, vengono gestiti gli accessi attraverso sbarra automatica"
            Top             =   360
            Width           =   4170
         End
         Begin VB.CheckBox chkIdentifyUserByCodFiscalePiva 
            Caption         =   "Identificazione  utente da Cod.Fiscale/PIVA"
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
            Left            =   120
            TabIndex        =   140
            ToolTipText     =   "Se attivo l'utente viene riconoscuto attraverso tessera con codice fiscale o Partita IVA"
            Top             =   720
            Width           =   4170
         End
         Begin VB.TextBox txtMaxUsersIsland 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            MaxLength       =   3
            TabIndex        =   139
            Text            =   "0"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Rinfresco Monitor Accessi (SEC)"
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
            Left            =   120
            TabIndex        =   148
            ToolTipText     =   "Una volta raggiunto il massimo consentito il semaforo diventa rosso"
            Top             =   3240
            Width           =   2865
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Medio Permanenza (MIN) "
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
            Left            =   120
            TabIndex        =   146
            ToolTipText     =   "Una volta raggiunto il massimo consentito il semaforo diventa rosso"
            Top             =   2520
            Width           =   2910
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Porta Seriale - Controllo Accessi"
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
            Index           =   36
            Left            =   120
            TabIndex        =   144
            Top             =   1920
            Width           =   2835
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Max Utenti per (semaforo rosso)"
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
            Left            =   120
            TabIndex        =   141
            ToolTipText     =   "Una volta raggiunto il massimo consentito il semaforo diventa rosso"
            Top             =   1200
            Width           =   2790
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   0
      Left            =   120
      ScaleHeight     =   6015
      ScaleWidth      =   7695
      TabIndex        =   3
      Top             =   1920
      Width           =   7695
      Begin VB.Frame Frame1 
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Index           =   5
         Left            =   6120
         TabIndex        =   35
         Top             =   0
         Width           =   3135
         Begin VB.PictureBox picRegistro 
            BorderStyle     =   0  'None
            Height          =   5055
            Left            =   120
            ScaleHeight     =   5055
            ScaleWidth      =   6855
            TabIndex        =   114
            Top             =   240
            Width           =   6855
            Begin VB.TextBox txtGGMovCarico 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               TabIndex        =   117
               Text            =   "1"
               Top             =   840
               Width           =   495
            End
            Begin VB.OptionButton optRegCaricoRegistro 
               Caption         =   $"frmImpostazioni.frx":0047
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Index           =   0
               Left            =   240
               MaskColor       =   &H8000000F&
               TabIndex        =   116
               Top             =   240
               Value           =   -1  'True
               Width           =   3615
            End
            Begin VB.OptionButton optRegCaricoRegistro 
               Caption         =   $"frmImpostazioni.frx":00D0
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Index           =   1
               Left            =   240
               TabIndex        =   115
               Top             =   1320
               Visible         =   0   'False
               Width           =   3735
            End
            Begin VB.Label Label1 
               Caption         =   "Ogni"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   119
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "giorno/i"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   118
               Top             =   840
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Allarmi Scadenze"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   5895
         Begin VB.OptionButton optQtaMaxPeriodo 
            Caption         =   "Qta Max Settimanale"
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
            Left            =   3600
            TabIndex        =   150
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton optQtaMaxPeriodo 
            Caption         =   "Qta Max Giornaliera"
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
            Index           =   0
            Left            =   3600
            TabIndex        =   149
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   2520
            TabIndex        =   12
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   2520
            TabIndex        =   11
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkVisualizzaAllarmeGiac 
            Caption         =   "Visualizza Allarme Giacenze"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Giacenza Normativa (GG)"
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
            TabIndex        =   14
            Top             =   360
            Width           =   2220
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Giacenza Gestionale (GG)"
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
            TabIndex        =   13
            Top             =   720
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "EcoPunto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Index           =   6
         Left            =   6000
         TabIndex        =   77
         Top             =   1920
         Width           =   4215
         Begin VB.CheckBox chkEcoPuntoDataMovimentoIns 
            Caption         =   "Usare la Data Importazione come Data Movimento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   82
            Top             =   1200
            Width           =   4455
         End
         Begin VB.CheckBox chkEcoPuntoCodPesa 
            Caption         =   "Utilizzare i primi 2 caratteri del Cod.Cer come Codice Pesa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   4455
         End
         Begin VB.CheckBox chkEcoPunto 
            Caption         =   "Le pesate provenienti  dal terminale ECOPUNTO generano  Giacenza per l'Isola Ecologica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   4455
         End
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
            Index           =   4
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cod.Causale +"
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
            Left            =   120
            TabIndex        =   80
            Top             =   1800
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Isola Ecologica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3855
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   1920
         Width           =   5895
         Begin VB.CheckBox chkFlgOrderByImmobiliUD 
            Caption         =   "OrderBy Immobili per UD"
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
            Left            =   2880
            TabIndex        =   136
            ToolTipText     =   "se impostato l'ordinamento immobili in ""PESATA"" è per utenza domestica altrimenti per utenza attività"
            Top             =   2400
            Width           =   2775
         End
         Begin VB.CheckBox chkFatturazioneConKD 
            Caption         =   "Fatturazione con KD"
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
            Left            =   120
            TabIndex        =   135
            ToolTipText     =   "se impostato effettua il calcolo della quantità max per generazione record fattura usando i KD legati alle superfici"
            Top             =   2760
            Width           =   2130
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   6
            ItemData        =   "frmImpostazioni.frx":0192
            Left            =   3240
            List            =   "frmImpostazioni.frx":01A8
            TabIndex        =   120
            Text            =   "Nessuno"
            Top             =   3360
            Width           =   2535
         End
         Begin VB.TextBox txtLunghezzaBadge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            MaxLength       =   2
            TabIndex        =   113
            Text            =   "0"
            Top             =   3360
            Width           =   2895
         End
         Begin VB.CheckBox chkDisabilitaNotifiche 
            Caption         =   "Disabilita tutte le notifiche"
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
            Left            =   2880
            TabIndex        =   111
            ToolTipText     =   "se abilitato non vengono mostrate le notifiche delle operazioni fatte"
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CheckBox chkImmobili 
            Caption         =   "Attiva Importazione Immobili"
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
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "se impostato abilita l'importazione degli Immobile e Utenti, in caso contrario vegono generati in automatico"
            Top             =   2400
            Width           =   2730
         End
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
            Index           =   5
            Left            =   120
            TabIndex        =   76
            Text            =   "Combo1"
            Top             =   1800
            Width           =   5535
         End
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
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   600
            Width           =   5535
         End
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
            Left            =   120
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   1200
            Width           =   5535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Maturazione Punti Per"
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
            Left            =   3240
            TabIndex        =   121
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Lunghezza Barcode"
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
            Left            =   120
            TabIndex        =   112
            Top             =   3120
            Width           =   1725
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Codice Pesa"
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
            Left            =   120
            TabIndex        =   75
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Società"
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
            TabIndex        =   8
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lbl 
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
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   1
      Left            =   120
      ScaleHeight     =   6495
      ScaleWidth      =   10215
      TabIndex        =   15
      Top             =   120
      Width           =   10215
      Begin VB.Frame Frame1 
         Caption         =   "Stampa Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5055
         Begin VB.TextBox txtMargDwnRegistro 
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
            Left            =   3765
            MaxLength       =   4
            TabIndex        =   42
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargUpRegistro 
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
            Left            =   3765
            MaxLength       =   4
            TabIndex        =   40
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMargdxRegistro 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   38
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtMargSxRegistro 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   36
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtNrCopieRegistro 
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
            Left            =   1485
            MaxLength       =   3
            TabIndex        =   21
            Top             =   1560
            Width           =   855
         End
         Begin VB.CheckBox chkAnteprimaRegistro 
            Alignment       =   1  'Right Justify
            Caption         =   "Anteprima"
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
            Left            =   2520
            TabIndex        =   20
            Top             =   1560
            Width           =   1455
         End
         Begin VB.ComboBox cmbStampanteRegistro 
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
            Left            =   1500
            TabIndex        =   19
            Text            =   "Combo2"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Inferiore"
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
            Left            =   2400
            TabIndex        =   43
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Superiore"
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
            Left            =   2400
            TabIndex        =   41
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Destro"
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
            TabIndex        =   39
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Sinistro"
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
            TabIndex        =   37
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Numero Copie"
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
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Stampante"
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
            TabIndex        =   17
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stampa Scontrino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Index           =   7
         Left            =   5145
         TabIndex        =   94
         Top             =   2280
         Width           =   5055
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   50
            Left            =   1065
            ScaleHeight     =   45
            ScaleWidth      =   3855
            TabIndex        =   133
            Top             =   2640
            Width           =   3855
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   50
            Left            =   1185
            ScaleHeight     =   45
            ScaleWidth      =   3735
            TabIndex        =   131
            Top             =   1965
            Width           =   3735
         End
         Begin VB.TextBox txtDichiarazioneScontrino 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   110
            Text            =   "frmImpostazioni.frx":01E3
            Top             =   3000
            Width           =   4815
         End
         Begin VB.ComboBox cmbStampanteScontrino 
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
            Left            =   1500
            TabIndex        =   101
            Text            =   "Combo2"
            Top             =   360
            Width           =   3375
         End
         Begin VB.CheckBox chkAnteprimaScontrino 
            Alignment       =   1  'Right Justify
            Caption         =   "Anteprima"
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
            Left            =   2520
            TabIndex        =   100
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtNrCopieScontrino 
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
            Left            =   1485
            MaxLength       =   3
            TabIndex        =   99
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtMargSxScontrino 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   98
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMargDxScontrino 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   97
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargUpScontrino 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   96
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMargDwnScontrino 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   95
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chkFlgInsRagSocScontrino 
            Alignment       =   1  'Right Justify
            Caption         =   "Rag. Soc. Comune"
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
            Left            =   120
            TabIndex        =   128
            Top             =   2040
            Width           =   2295
         End
         Begin VB.CheckBox chkFlgInsImpiantoScontrino 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   2520
            TabIndex        =   129
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox chkFlgInsComuneImpiantoScontrino 
            Alignment       =   1  'Right Justify
            Caption         =   "Comune Impianto"
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
            Left            =   120
            TabIndex        =   130
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CheckBox chkFlgInsDichiarazioneScontrino 
            Alignment       =   1  'Right Justify
            Caption         =   "Inserire Dichirazione di respondabilità"
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
            Left            =   120
            TabIndex        =   109
            Top             =   2760
            Width           =   4815
         End
         Begin VB.Label Label6 
            Caption         =   "Pie Pagina"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   2535
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Intestazione"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   1863
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Stampante"
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
            Index           =   35
            Left            =   120
            TabIndex        =   107
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Numero Copie"
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
            TabIndex        =   106
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Sinistro"
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
            Left            =   120
            TabIndex        =   105
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Destro"
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
            Left            =   120
            TabIndex        =   104
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Superiore"
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
            Left            =   2520
            TabIndex        =   103
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Inferiore"
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
            Left            =   2520
            TabIndex        =   102
            Top             =   1200
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stampa Documento Ingresso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Index           =   4
         Left            =   0
         TabIndex        =   28
         Top             =   2280
         Width           =   5055
         Begin VB.TextBox txtMargDwnPesa 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   50
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargUpPesa 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMargDxPesa 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   46
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargSxPesa 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   44
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtNrCopiePesa 
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
            Left            =   1485
            MaxLength       =   3
            TabIndex        =   31
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox chkAnteprimaPesa 
            Alignment       =   1  'Right Justify
            Caption         =   "Anteprima"
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
            Left            =   2520
            TabIndex        =   30
            Top             =   1560
            Width           =   1575
         End
         Begin VB.ComboBox cmbStampantePesa 
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
            Left            =   1500
            TabIndex        =   29
            Text            =   "Combo2"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Inferiore"
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
            Left            =   2520
            TabIndex        =   51
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Superiore"
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
            Left            =   2520
            TabIndex        =   49
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Destro"
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
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Sinistro"
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
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Numero Copie"
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
            TabIndex        =   33
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Stampante"
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
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stampa Formulario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Index           =   3
         Left            =   5160
         TabIndex        =   22
         Top             =   0
         Width           =   5055
         Begin VB.TextBox txtMargDwnFormulario 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   58
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargUpFormulario 
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
            Left            =   3885
            MaxLength       =   4
            TabIndex        =   56
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMargDxFormulario 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   54
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMargSxFormulario 
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
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   52
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtNrCopieFormulario 
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
            Left            =   1485
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox chkAnteprimaFormulario 
            Alignment       =   1  'Right Justify
            Caption         =   "Anteprima"
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
            Left            =   2520
            TabIndex        =   24
            Top             =   1560
            Width           =   1575
         End
         Begin VB.ComboBox cmbStampanteFormulario 
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
            Left            =   1500
            TabIndex        =   23
            Text            =   "Combo2"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Inferiore"
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
            Left            =   2520
            TabIndex        =   59
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Superiore"
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
            Left            =   2520
            TabIndex        =   57
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Destro"
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
            Left            =   120
            TabIndex        =   55
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Marg.Sinistro"
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
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Numero Copie"
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
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Stampante"
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
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   5
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   8895
      TabIndex        =   122
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame4 
         Caption         =   "Abilitazione Bonus €/ Premi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Width           =   9615
         Begin MSComctlLib.ListView lvwBonus 
            Height          =   1335
            Left            =   120
            TabIndex        =   125
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lvwPremi 
            Height          =   1335
            Left            =   4080
            TabIndex        =   126
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblPremi 
            Caption         =   "Premi"
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
            Left            =   4080
            TabIndex        =   127
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Bonus €"
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
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   4
      Left            =   240
      ScaleHeight     =   3015
      ScaleWidth      =   8895
      TabIndex        =   84
      Top             =   6480
      Width           =   8895
      Begin VB.Frame Frame3 
         Caption         =   "Backup / Storicizzazione"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   9615
         Begin VB.TextBox txtPercorsoStoricizzazione 
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
            Left            =   240
            TabIndex        =   91
            Top             =   1680
            Width           =   8055
         End
         Begin VB.CommandButton cmdPathStorico 
            Caption         =   "..."
            Height          =   495
            Left            =   8400
            TabIndex        =   90
            Top             =   1635
            Width           =   495
         End
         Begin VB.CommandButton cmdPathBackup 
            Caption         =   "..."
            Height          =   495
            Left            =   8400
            TabIndex        =   89
            Top             =   555
            Width           =   495
         End
         Begin VB.CheckBox chkStoricizzazioneAuto 
            Caption         =   "Visuallizza Allarme Storicizzazione"
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
            Left            =   240
            TabIndex        =   88
            Top             =   2160
            Width           =   3375
         End
         Begin VB.CheckBox chkRichiedereAUscitaProgramma 
            Caption         =   "Richiedere Backup all'Uscita programma"
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
            Left            =   240
            TabIndex        =   87
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtPercorsoBackup 
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
            Left            =   240
            TabIndex        =   86
            Top             =   600
            Width           =   8055
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Percorso Storicizzazione"
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
            Index           =   29
            Left            =   240
            TabIndex        =   93
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Percorso Backup"
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
            Left            =   240
            TabIndex        =   92
            Top             =   360
            Width           =   1500
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   3
      Left            =   1920
      ScaleHeight     =   3495
      ScaleWidth      =   7575
      TabIndex        =   60
      Top             =   120
      Width           =   7575
      Begin VB.Frame Frame2 
         Caption         =   "Movimento per Carico da Uscita Merce"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3255
         Left            =   120
         TabIndex        =   61
         Top             =   0
         Width           =   6855
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
            Index           =   3
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   600
            Width           =   3615
         End
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
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   600
            Width           =   2655
         End
         Begin VB.CommandButton cmbRicerca 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   1560
            Picture         =   "frmImpostazioni.frx":02AC
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   240
            TabIndex        =   65
            ToolTipText     =   "Codice dell'utente rilevato"
            Top             =   1320
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwImmobili 
            Height          =   495
            Left            =   240
            TabIndex        =   68
            Top             =   2160
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   873
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cod.User"
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
            Left            =   3000
            TabIndex        =   74
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   72
            Top             =   1290
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ragione Sociale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   2160
            TabIndex        =   71
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comune"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2160
            TabIndex        =   70
            Top             =   1815
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indirizzo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   5
            Left            =   2160
            TabIndex        =   69
            Top             =   1605
            Width           =   660
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Immobile"
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
            Left            =   240
            TabIndex        =   64
            Top             =   1920
            Width           =   1245
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Utente"
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
            Left            =   240
            TabIndex        =   63
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cod.Causale +"
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
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1290
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   2
      Left            =   4560
      ScaleHeight     =   1575
      ScaleWidth      =   5655
      TabIndex        =   34
      Top             =   4200
      Width           =   5655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmImpostazioni.frx":0A16
         Height          =   735
         Left            =   480
         TabIndex        =   108
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1296
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
         DataMember      =   "cmdNumeratori"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Entita"
            Caption         =   "Entita"
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
            DataField       =   "Numeratore"
            Caption         =   "Numeratore"
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
            DataField       =   "Anno"
            Caption         =   "Anno"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   2025,071
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   2145,26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7920
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   9120
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1455
      Left            =   7920
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2566
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Generale"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampe"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Numeratori"
            Key             =   "Numeratori"
            Object.ToolTipText     =   "Gestione Numeratori"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimenti"
            Key             =   "MOVCOR"
            Object.ToolTipText     =   "Configurazione dei movimenti per sistemare le quantità"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Percorsi"
            Key             =   "Percorsi"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bonus / Premi"
            Key             =   "bonuspremi"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Accessi"
            Key             =   "Accessi"
            Object.ToolTipText     =   "Impostazione gestione accessi"
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
Attribute VB_Name = "frmImpostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TXT_GGREMINDERGIACNORMATIVA = 0
Private Const TXT_GGREMINDERGIACGESTIONALE = 1
Private Const TXT_COD_UTENTE = 2
Private mCodImmobile As String



Private Sub cmbRicerca_Click(Index As Integer)
    frmRicercaUtente.schowMe txt(TXT_COD_UTENTE), "U"
    If txt(TXT_COD_UTENTE).Text <> vbNullString Then caricaDatiUtente txt(TXT_COD_UTENTE).Text
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If salva Then Unload Me
    End Select
         
End Sub

Private Function salva() As Boolean
    Dim f As New FileSystemObject
    On Error GoTo salvaErr
    salva = False
    
    
   
   
    If f.GetParentFolderName(txtPercorsoBackup.Text) = vbNullString Then
        If Not f.DriveExists(txtPercorsoBackup.Text) Then
            MsgBox "Attenzione disco Backup " & txtPercorsoBackup.Text & " non trovato", vbCritical + vbOKOnly, App.Title
            Exit Function
        End If
    ElseIf Not f.FolderExists(txtPercorsoBackup.Text) Then
        MsgBox "Attenzione percorso Backup " & txtPercorsoBackup.Text & " non trovato", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
   
    If f.GetParentFolderName(txtPercorsoStoricizzazione.Text) = vbNullString Then
        If Not f.DriveExists(txtPercorsoStoricizzazione.Text) Then
            MsgBox "Attenzione disco Backup " & txtPercorsoStoricizzazione.Text & " non trovato", vbCritical + vbOKOnly, App.Title
            Exit Function
        End If
    ElseIf Not f.FolderExists(txtPercorsoStoricizzazione.Text) Then
        MsgBox "Attenzione percorso Storicizzazione " & txtPercorsoStoricizzazione.Text & " non trovato", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    
    gPuntiAssegnatiPer = Combo1(6).ItemData(Combo1(6).ListIndex)
    
    MDIMain.mnuScuole.enabled = gPuntiAssegnatiPer > 0
    
    If gPuntiAssegnatiPer > 0 Then
        'verifico se la gestione impianti e utenti è attiva l'importazione
        If chkImmobili.value = 1 Then
            '
            Dim rs  As ADODB.Recordset
            Set rs = gDBC.Execute("select * from Contatori where Entita='UTENTI'")
            If rs.EOF Then
                getNumeratore UTENTI, True, 1000000
            End If
        End If
    End If
    
    MDIMain.MnuImmobiliImp.enabled = chkImmobili.value = 1
    
    gActiveAccess = chkActiveAccess.value
    gIdentifyUserByCodFiscalePiva = chkIdentifyUserByCodFiscalePiva.value
    gMaxUsersIsland = txtMaxUsersIsland.Text
    gRecycleTimeUsersIsland = txtMaxMinInIsola.Text
    gAccessDeviceCOMName = cmb(7).Text
    gSecRefreshAccessMonitor = txtSecRefreshAccessMonitor.Text
    
    
    gPercorsoBackup = txtPercorsoBackup.Text
    gPercorsoStoricizzazione = txtPercorsoStoricizzazione.Text
    gRichiedereAUscitaProgramma = chkRichiedereAUscitaProgramma.value
    gStoricizzazioneAutomatica = chkStoricizzazioneAuto.value
    gRegCaricoSuRegistroOgni = txtGGMovCarico.Text
    gRegCaricoSuRegistroAllUscita = optRegCaricoRegistro(1).value
    gDisabilitaNotifiche = chkDisabilitaNotifiche.value
    gQtaMaxGiornaliera = optQtaMaxPeriodo(0).value


    gGGAllarmeGiacNormativa = txt(TXT_GGREMINDERGIACNORMATIVA).Text
    gGGAllarmeGiacGestionale = txt(TXT_GGREMINDERGIACGESTIONALE).Text
    
    If Combo1(1).ListCount > 0 Then gIdImpiantoProduttore = Combo1(1).ItemData(Combo1(1).ListIndex)
    If Combo1(0).ListCount > 0 Then gIdSocietaProduttore = Combo1(0).ItemData(Combo1(0).ListIndex)
    
    gStampaPesataAnteprima = chkAnteprimaPesa.value
    gStampaRegistroAnteprima = chkAnteprimaRegistro.value
    gStampaFormularioAnteprima = chkAnteprimaFormulario.value
    gStampaScontrinoAnteprima = chkAnteprimaScontrino.value
    gStampaScontrinoInsDichiaraDiResponsabilita = chkFlgInsDichiarazioneScontrino.value
    
    gStampaScontrinoInsComuneImpiantoScontrino = chkFlgInsComuneImpiantoScontrino
    gStampaScontrinoInsImpiantoScontrino = chkFlgInsImpiantoScontrino
    gStampaScontrinoInsRagSocScontrino = chkFlgInsRagSocScontrino
    
    gStampaScontrinoDichiarazione = txtDichiarazioneScontrino.Text
    
    
    gStampaPesataStampante = cmbStampantePesa.Text
    gStampaRegistroStampante = cmbStampanteRegistro.Text
    gStampaFormularioStampante = cmbStampanteFormulario.Text
    gStampaScontrinoStampante = cmbStampanteScontrino.Text
    
    gStampaPesataNrCopie = txtNrCopiePesa
    gStampaRegistroNrCopie = txtNrCopieRegistro
    gStampaFormularioNrCopie = txtNrCopieFormulario
    gStampaScontrinoNrCopie = txtNrCopieScontrino
    
    
    gStampaFormularioMargSx = txtMargSxFormulario.Text
    gStampaPesataMargSx = txtMargSxPesa.Text
    gStampaRegistroMargSx = txtMargSxRegistro.Text
    gStampaScontrinoMargSx = txtMargSxScontrino.Text
        
    gStampaFormularioMargDx = txtMargDxFormulario.Text
    gStampaPesataMargDx = txtMargDxPesa.Text
    gStampaRegistroMargDx = txtMargdxRegistro.Text
    gStampaScontrinoMargDx = txtMargDxScontrino.Text
        
    gStampaFormularioMargUp = txtMargUpFormulario.Text
    gStampaPesataMargUp = txtMargUpPesa.Text
    gStampaRegistroMargUp = txtMargUpRegistro.Text
    gStampaScontrinoMargUp = txtMargUpScontrino.Text
        
    gStampaFormularioMargDwn = txtMargDwnFormulario.Text
    gStampaPesataMargDwn = txtMargDwnPesa.Text
    gStampaRegistroMargDwn = txtMargDwnRegistro.Text
    gStampaScontrinoMargDwn = txtMargDwnScontrino.Text
    gLunghezzaBadge = txtLunghezzaBadge.Text
    
    
    gCodCausaleMovManuale = left$(Combo1(2).Text, 2)
    gCodUserMovManuale = left$(Combo1(3).Text, InStr(1, Combo1(3).Text, " - ") - 1)
    
    If mCodImmobile = "" Then
        MsgBox "Attenzione Immobile non impostato"
    Else
        gCodUtenteMovManuale = txt(TXT_COD_UTENTE).Text
        gCodImmobileMovManuale = Mid$(mCodImmobile, 2)
    End If
    
    gflgMovECOPuntoGeneraGiacenza = chkEcoPunto.value
    gflgOrderByImmobiliUD = chkFlgOrderByImmobiliUD.value
    gflgCodPesaECOPunto = chkEcoPuntoCodPesa.value
    gflgEcoPuntoDataMovimentoIns = chkEcoPuntoDataMovimentoIns.value
    gflgImmobili = chkImmobili.value
    gflgFatturazioneConKD = chkFatturazioneConKD.value
    
    
    gCodCausaleMovEcoPuntoIN = left$(Combo1(4).Text, 2)
    gPesaIsolaCod = left$(Combo1(5).Text, 2)
    
    SaveSetting App.Title, "IMPOSTAZIONI", "FATTURAZIONECONKD", gflgFatturazioneConKD
    SaveSetting App.Title, "IMPOSTAZIONI", "GESTIONEIMMOBILI", gflgImmobili
    SaveSetting App.Title, "IMPOSTAZIONI", "LUNGHEZZABADGE", gLunghezzaBadge
    
    SaveSetting App.Title, "IMPOSTAZIONI", "ACTIVEACCESS", gActiveAccess
    SaveSetting App.Title, "IMPOSTAZIONI", "IDENTIFYUSERBYCODFISCALEPIVA", gIdentifyUserByCodFiscalePiva
    SaveSetting App.Title, "IMPOSTAZIONI", "MAXUSERSISLAND", gMaxUsersIsland
    SaveSetting App.Title, "IMPOSTAZIONI", "RECYCLETIMEUSERSISLAND", gRecycleTimeUsersIsland
    SaveSetting App.Title, "IMPOSTAZIONI", "ACCESSDEVICECOMNAME", gAccessDeviceCOMName
    SaveSetting App.Title, "IMPOSTAZIONI", "SECREFRESHACCESSMONITOR", gSecRefreshAccessMonitor
          
    
    SaveSetting App.Title, "IMPOSTAZIONI", "CODPESAISOLA", gPesaIsolaCod
    SaveSetting App.Title, "IMPOSTAZIONI", "RECARICOSUREGISTROOGNI", gRegCaricoSuRegistroOgni
    SaveSetting App.Title, "IMPOSTAZIONI", "RECARICOSUREGISTROALLUSCITA", gRegCaricoSuRegistroAllUscita
    SaveSetting App.Title, "IMPOSTAZIONI", "QTAMASSIMAGIORNALIERA", gQtaMaxGiornaliera
    
    
    SaveSetting App.Title, "IMPOSTAZIONI", "GGREMINDERGIACNORMATIVA", gGGAllarmeGiacNormativa
    SaveSetting App.Title, "IMPOSTAZIONI", "GGREMINDERGIACGESTIONALE", gGGAllarmeGiacGestionale
    SaveSetting App.Title, "IMPOSTAZIONI", "VISUALIZZAALLARMEGIACENZA", chkVisualizzaAllarmeGiac.value
    
    
    SaveSetting App.Title, "IMPOSTAZIONI", "SOCIETA", gIdSocietaProduttore
    SaveSetting App.Title, "IMPOSTAZIONI", "IMPIANTO", gIdImpiantoProduttore
    SaveSetting App.Title, "IMPOSTAZIONI", "NOTIFICHE", gDisabilitaNotifiche
    SaveSetting App.Title, "IMPOSTAZIONI", "PUNTIASSEGNATIPER", gPuntiAssegnatiPer
    
    
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATAANTEPRIMA", gStampaPesataAnteprima
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATASTAMPANTE", gStampaPesataStampante
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATANRCOPIE", gStampaPesataNrCopie
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGSX", gStampaPesataMargSx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGDX", gStampaPesataMargDx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGUP", gStampaPesataMargUp
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGDWN", gStampaPesataMargDwn

    
    'scontrino
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOANTEPRIMA", gStampaScontrinoAnteprima
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOSTAMPANTE", gStampaScontrinoStampante
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINONRCOPIE", gStampaScontrinoNrCopie
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGSX", gStampaScontrinoMargSx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGDX", gStampaScontrinoMargDx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGUP", gStampaScontrinoMargUp
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGDWN", gStampaScontrinoMargDwn
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGDICHIARAZIONE", gStampaScontrinoInsDichiaraDiResponsabilita
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGCOMUNEIMPIANTO", gStampaScontrinoInsComuneImpiantoScontrino
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGIMPIANTO", gStampaScontrinoInsImpiantoScontrino
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGRAGSOC", gStampaScontrinoInsRagSocScontrino
    
    
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPASCONTRINODICHIARAZIONE", gStampaScontrinoDichiarazione

    

    'registro
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROANTEPRIMA", gStampaRegistroAnteprima
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROSTAMPANTE", gStampaRegistroStampante
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTRONRCOPIE", gStampaRegistroNrCopie
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGSX", gStampaRegistroMargSx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGDX", gStampaRegistroMargDx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGUP", gStampaRegistroMargUp
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGDWN", gStampaRegistroMargDwn

    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOANTEPRIMA", gStampaFormularioAnteprima
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOSTAMPANTE", gStampaFormularioStampante
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIONRCOPIE", gStampaFormularioNrCopie
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGSX", gStampaFormularioMargSx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGDX", gStampaFormularioMargDx
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGUP", gStampaFormularioMargUp
    SaveSetting App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGDWN", gStampaFormularioMargDwn

    SaveSetting App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVDAUSCITAPIU", gCodCausaleMovManuale
    SaveSetting App.Title, "IMPOSTAZIONI", "CODUTENTEMOVDAUSCITAPIU", gCodUtenteMovManuale
    SaveSetting App.Title, "IMPOSTAZIONI", "CODIMMOBILEMOVDAUSCITAPIU", gCodImmobileMovManuale
    SaveSetting App.Title, "IMPOSTAZIONI", "CODUSERMOVDAUSCITAPIU", gCodUserMovManuale
    SaveSetting App.Title, "IMPOSTAZIONI", "FLGORDERBYIMMOBILIUD", gflgOrderByImmobiliUD
    
    'ecopunto
    SaveSetting App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVECOPUNTOIN", gCodCausaleMovEcoPuntoIN
    SaveSetting App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVECOPUNTOOUT", gCodCausaleMovEcoPuntoOUT
    SaveSetting App.Title, "IMPOSTAZIONI", "FLGMOVECOPUNTOGENMOV", gflgMovECOPuntoGeneraGiacenza
    SaveSetting App.Title, "IMPOSTAZIONI", "FLGCODPESAECOPUNTO", gflgCodPesaECOPunto
    SaveSetting App.Title, "IMPOSTAZIONI", "FLGECOPUNTODATAMOVIMENTOINS", gflgEcoPuntoDataMovimentoIns
    
    
    'backup
    SaveSetting App.Title, "DATABASE", "PERCORSOBACKUP", gPercorsoBackup
    SaveSetting App.Title, "DATABASE", "RICHIESTABACKUPUSCITA", gRichiedereAUscitaProgramma

    'storico
    SaveSetting App.Title, "DATABASE", "PERCORSOSTORICO", gPercorsoStoricizzazione
    SaveSetting App.Title, "DATABASE", "STORICIZZAZIONEAUTO", gStoricizzazioneAutomatica
    
    SalvaBonusPremi
    
    letturaStrutturaMessaggioPesa
    
    salva = True
    Exit Function
salvaErr:
    MsgBox "Attenzione verificare i dati inseriti:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Sub cmdPathBackup_Click()
    Dim sFile As String
    DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso per Backup", Me.hWnd, "IsolaEcologica.BAK", , App.Path, "*.bak"
    If sFile <> vbNullString Then
        txtPercorsoBackup.Text = sFile
    End If
End Sub

Private Sub cmdPathStorico_Click()

    Dim sFile As String
    DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso per lo Storico", Me.hWnd, "Storico*.*", , App.Path, "*.mdb"
    If sFile <> vbNullString Then
        txtPercorsoStoricizzazione.Text = sFile
    End If


End Sub

Private Sub Combo1_Click(Index As Integer)
   If Index = 0 Then
        caricaImpianti Combo1(0).ItemData(Combo1(0).ListIndex)
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    
    Me.width = 15360
    Me.Height = 8500
    
    applicaRestrizioniUtente
    
    caricaNomiStampanti
    txt(TXT_GGREMINDERGIACNORMATIVA).Text = gGGAllarmeGiacNormativa
    txt(TXT_GGREMINDERGIACGESTIONALE).Text = gGGAllarmeGiacGestionale
    
    chkVisualizzaAllarmeGiac.value = IIf(CBool(GetSetting(App.Title, "IMPOSTAZIONI", "VISUALIZZAALLARMEGIACENZA", True)), 1, 0)
    
    chkAnteprimaPesa.value = gStampaPesataAnteprima
    chkAnteprimaRegistro.value = gStampaRegistroAnteprima
    chkAnteprimaFormulario.value = gStampaFormularioAnteprima
    chkAnteprimaScontrino.value = gStampaScontrinoAnteprima
    chkImmobili.value = gflgImmobili
    chkFatturazioneConKD.value = gflgFatturazioneConKD
    
    chkDisabilitaNotifiche.value = gDisabilitaNotifiche
    
    
    
    If DevEnvMDB.rscmdNumeratori.state = adStateOpen Then
        DevEnvMDB.rscmdNumeratori.Close
        DevEnvMDB.cmdNumeratori
    Else
        DevEnvMDB.cmdNumeratori
    End If
    
    DataGrid1.DataMember = "cmdNumeratori"
    Set DataGrid1.DataSource = DevEnvMDB
    
    
    
    If gStampaPesataStampante <> "" Then cmbStampantePesa.Text = gStampaPesataStampante
    If gStampaRegistroStampante <> "" Then cmbStampanteRegistro.Text = gStampaRegistroStampante
    If gStampaFormularioStampante <> "" Then cmbStampanteFormulario.Text = gStampaFormularioStampante
    If gStampaScontrinoStampante <> "" Then cmbStampanteScontrino.Text = gStampaScontrinoStampante
    
    chkFlgInsDichiarazioneScontrino.value = gStampaScontrinoInsDichiaraDiResponsabilita
    chkFlgInsComuneImpiantoScontrino.value = gStampaScontrinoInsComuneImpiantoScontrino
    chkFlgInsImpiantoScontrino.value = gStampaScontrinoInsImpiantoScontrino
    chkFlgInsRagSocScontrino.value = gStampaScontrinoInsRagSocScontrino
    
    
    txtDichiarazioneScontrino = gStampaScontrinoDichiarazione
    
    
    txtNrCopiePesa.Text = gStampaPesataNrCopie
    txtNrCopieRegistro.Text = gStampaRegistroNrCopie
    txtNrCopieFormulario = gStampaFormularioNrCopie
    txtNrCopieScontrino = gStampaScontrinoNrCopie
        
    txtMargSxFormulario.Text = gStampaFormularioMargSx
    txtMargSxPesa.Text = gStampaPesataMargSx
    txtMargSxRegistro.Text = gStampaRegistroMargSx
    txtMargSxScontrino.Text = gStampaScontrinoMargSx
        
    txtMargDxFormulario.Text = gStampaFormularioMargDx
    txtMargDxPesa.Text = gStampaPesataMargDx
    txtMargdxRegistro.Text = gStampaRegistroMargDx
    txtMargDxScontrino.Text = gStampaScontrinoMargDx
        
    txtMargUpFormulario.Text = gStampaFormularioMargUp
    txtMargUpPesa.Text = gStampaPesataMargUp
    txtMargUpRegistro.Text = gStampaRegistroMargUp
    txtMargUpScontrino.Text = gStampaScontrinoMargUp
        
    txtMargDwnFormulario.Text = gStampaFormularioMargDwn
    txtMargDwnPesa.Text = gStampaPesataMargDwn
    txtMargDwnRegistro.Text = gStampaRegistroMargDwn
    txtMargDwnScontrino.Text = gStampaScontrinoMargDwn
    
    txtLunghezzaBadge.Text = gLunghezzaBadge
        
        
    txtGGMovCarico.Text = gRegCaricoSuRegistroOgni
    optRegCaricoRegistro(1).value = gRegCaricoSuRegistroAllUscita
    optRegCaricoRegistro(0).value = Not optRegCaricoRegistro(1)
    optQtaMaxPeriodo(0).value = gQtaMaxGiornaliera
    optQtaMaxPeriodo(1).value = Not gQtaMaxGiornaliera
        
    chkEcoPunto.value = gflgMovECOPuntoGeneraGiacenza
    chkEcoPuntoCodPesa.value = gflgCodPesaECOPunto
    chkEcoPuntoDataMovimentoIns = gflgEcoPuntoDataMovimentoIns
    chkFlgOrderByImmobiliUD.value = gflgOrderByImmobiliUD
    
    chkActiveAccess.value = gActiveAccess
    chkIdentifyUserByCodFiscalePiva.value = gIdentifyUserByCodFiscalePiva
    txtMaxUsersIsland.Text = gMaxUsersIsland
    txtMaxMinInIsola.Text = gRecycleTimeUsersIsland
    cmb(7).Text = gAccessDeviceCOMName
    txtSecRefreshAccessMonitor.Text = gSecRefreshAccessMonitor
    
    Picture1(0).Visible = True
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    Picture1(3).Visible = False
    Picture1(4).Visible = False
    
    CaricaSocieta
    
    lvwImmobili.ColumnHeaders.Add , "cod", "Cod.Imm."
    lvwImmobili.ColumnHeaders.Add , , "Tipo"
    lvwImmobili.ColumnHeaders(lvwImmobili.ColumnHeaders.Count).width = 0
    lvwImmobili.ColumnHeaders.Add , , "Comune"
    lvwImmobili.ColumnHeaders.Add , "TipoUtenza", "Tipo Utenza"
    lvwImmobili.ColumnHeaders.Add , , "Indirizzo"
    lvwImmobili.ColumnHeaders.Add , , "SubAlterno"
    lvwImmobili.ColumnHeaders.Add , , "Num.Interno"
    lvwImmobili.ColumnHeaders.Add , , "Cod.Tariffa"
    lvwImmobili.ColumnHeaders.Add , , "Tariffa"
    
    Label1(4).Caption = ""
    Label1(6).Caption = ""
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from causali"
    rs.Open
    Do While Not rs.EOF
        Combo1(2).AddItem rs!codCausale & " - " & rs!Descrizione
        Combo1(4).AddItem rs!codCausale & " - " & rs!Descrizione
        If (rs!codCausale = gCodCausaleMovManuale) Then Combo1(2).ListIndex = Combo1(2).ListCount - 1
        If (rs!codCausale = gCodCausaleMovEcoPuntoIN) Then Combo1(4).ListIndex = Combo1(4).ListCount - 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from users"
    rs.Open
    Do While Not rs.EOF
        Combo1(3).AddItem rs!codUser & " - " & rs!Nome & " " & rs!Cognome
        If (rs!codUser = gCodUserMovManuale) Then
            Combo1(3).ListIndex = Combo1(3).ListCount - 1
        End If
        rs.MoveNext
    Loop
    
    If Combo1(3).ListIndex < 0 Then Combo1(3).ListIndex = 0
    
    rs.Close
    Set rs.ActiveConnection = Nothing
    txt(TXT_COD_UTENTE).Text = gCodUtenteMovManuale
    
    
    
    
    g_rsPese.MoveFirst
    g_rsPese.Sort = "ComID"
    
    Do While Not g_rsPese.EOF
        If nullToBlank(g_rsPese!PCNAME) = vbNullString And g_rsPese!Active Then
            Combo1(5).AddItem g_rsPese!codPesa & " - " & g_rsPese!Descrizione
        End If
        
        If g_rsPese!codPesa = gPesaIsolaCod Then
             Combo1(5).ListIndex = Combo1(5).ListCount - 1
        End If
        
        g_rsPese.MoveNext
    Loop
    
    If Combo1(5).ListIndex < 0 And Combo1(5).ListCount > 0 Then Combo1(5).ListIndex = 0
    Combo1(6).ListIndex = gPuntiAssegnatiPer
    
    txtPercorsoBackup.Text = gPercorsoBackup
    chkRichiedereAUscitaProgramma.value = IIf(gRichiedereAUscitaProgramma, 1, 0)
    chkStoricizzazioneAuto.value = IIf(gStoricizzazioneAutomatica, 1, 0)
    txtPercorsoStoricizzazione.Text = gPercorsoStoricizzazione
    
    caricaImmobili gCodUtenteMovManuale
    
    Dim i  As Long
    For i = Picture1.LBound + 1 To Picture1.ubound
        Picture1(i).Visible = False
    Next
    
    Combo1(6).enabled = Not EsisteMovimentiPerLivelloBonus
    
    caricaBonusPremi
End Sub

Private Function EsisteMovimentiPerLivelloBonus() As Boolean
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("SELECT 1 FROM TipiAnagrafe INNER JOIN (Movimenti INNER JOIN Utenti ON Movimenti.CodUtente = Utenti.CodUtente) " & _
              "ON TipiAnagrafe.TipoAnagrafe = Utenti.TipoAnagrafe " & _
              "WHERE TipiAnagrafe.TipoAnagrafe In ('C','E','I','S')")
              
    EsisteMovimentiPerLivelloBonus = Not rs.EOF
End Function


Private Sub caricaNomiStampanti()
    Dim P As Printer
    
    cmbStampanteFormulario.Clear
    cmbStampantePesa.Clear
    cmbStampanteRegistro.Clear
    cmbStampanteScontrino.Clear
    
    For Each P In Printers
        cmbStampanteFormulario.AddItem P.DeviceName
        cmbStampantePesa.AddItem P.DeviceName
        cmbStampanteRegistro.AddItem P.DeviceName
        cmbStampanteScontrino.AddItem P.DeviceName
    Next
End Sub

Private Sub CaricaSocieta()
    If DevEnvMDB.rscmdSocieta.state = adStateClosed Then
        DevEnvMDB.cmdSocieta
    Else
        DevEnvMDB.rscmdSocieta.MoveFirst
    End If
    
    Combo1(0).Clear
    
    Do While Not DevEnvMDB.rscmdSocieta.EOF
        Combo1(0).AddItem DevEnvMDB.rscmdSocieta!ragioneSociale
        Combo1(0).ItemData(Combo1(0).ListCount - 1) = DevEnvMDB.rscmdSocieta!idSocieta
        If (DevEnvMDB.rscmdSocieta!idSocieta = gIdSocietaProduttore) Then
            Combo1(0).ListIndex = Combo1(0).ListCount - 1
        End If
        DevEnvMDB.rscmdSocieta.MoveNext
    Loop
    
    If Combo1(0).ListCount > 0 And Combo1(0).ListIndex < 0 Then
        Combo1(0).ListIndex = 0
        caricaImpianti Combo1(0).ItemData(0)
    End If


End Sub

Private Sub caricaImpianti(idSocieta As Long)
    Combo1(1).Clear
    
    If DevEnvMDB.rscmdImpianti.state = adStateOpen Then
        DevEnvMDB.rscmdImpianti.Close
    End If
    
    DevEnvMDB.cmdImpianti
    DevEnvMDB.rscmdImpianti.Filter = " idSoggetto=" & idSocieta & " and codSoggetto='S'"
    
    Do While Not DevEnvMDB.rscmdImpianti.EOF
        Combo1(1).AddItem nullToBlank(DevEnvMDB.rscmdImpianti!ragioneSociale)
        Combo1(1).ItemData(Combo1(1).ListCount - 1) = DevEnvMDB.rscmdImpianti!idImpianto
        If (DevEnvMDB.rscmdImpianti!idImpianto = gIdImpiantoProduttore) Then
            Combo1(1).ListIndex = Combo1(1).ListCount - 1
        End If
        DevEnvMDB.rscmdImpianti.MoveNext
    Loop
    If Combo1(1).ListCount > 0 And Combo1(1).ListIndex < 0 Then
        Combo1(1).ListIndex = 0
    End If

End Sub


Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - 50
    posLeft = Me.ScaleWidth - cmd(i).width - 50
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - 50 - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    
    TabStrip1.Move 50, 50, Me.ScaleWidth - 150, cmd(0).top - 150

    For i = 0 To Picture1.ubound
        Picture1(i).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    Next
    
    
    
    DataGrid1.Move 50, 50, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
End Sub

Private Sub Picture1_Resize(Index As Integer)
    Select Case Index
        Case 0
            
            Frame1(5).width = Picture1(Index).ScaleWidth - Frame1(5).left - 150
            picRegistro.Move 150, 200, Frame1(5).width - 250, Frame1(5).Height - 250
            Frame1(6).width = Picture1(Index).ScaleWidth - Frame1(6).left - 150
            
            Frame1(0).Height = Frame1(5).Height
            Frame1(1).top = Frame1(5).Height + 150
            Frame1(6).top = Frame1(5).Height + 150
            
            Frame3.width = Picture1(Index).ScaleWidth - 150
            
            optRegCaricoRegistro(0).width = Frame1(5).width - optRegCaricoRegistro(0).left - 150
            optRegCaricoRegistro(1).width = Frame1(5).width - optRegCaricoRegistro(0).left - 150
            chkEcoPunto.width = Frame1(6).width - chkEcoPunto.left - 150
            chkEcoPuntoCodPesa.width = Frame1(6).width - chkEcoPuntoCodPesa.left - 150
            chkEcoPuntoDataMovimentoIns.width = Frame1(6).width - chkEcoPuntoDataMovimentoIns.left - 150
        Case 3
            Frame2.width = Picture1(Index).ScaleWidth - 300
            Frame2.Height = Picture1(Index).ScaleHeight - 150
            lvwImmobili.width = Frame2.width - 300
            lvwImmobili.Height = Frame2.Height - lvwImmobili.top - 150
            
        Case 5
            Frame4.width = Picture1(Index).ScaleWidth - 150
            Frame4.Height = Picture1(Index).ScaleHeight - 150
            
            lvwBonus.width = Picture1(Index).ScaleWidth / 2 - lvwBonus.left - 150
            lvwBonus.Height = Picture1(Index).ScaleHeight - lvwBonus.top - 300
            
            
            lvwPremi.Move lvwBonus.width + lvwBonus.left + 150, lvwBonus.top, Picture1(Index).ScaleWidth - lvwPremi.left - 300, Picture1(Index).ScaleHeight - lvwBonus.top - 150
            lblPremi.left = lvwPremi.left
    End Select
End Sub

Private Sub TabStrip1_Click()
    Dim i As Integer
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    For i = Picture1.LBound To Picture1.ubound
        Picture1(i).Visible = False
    Next
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
End Sub

Public Sub applicaRestrizioniUtente()
    cmd(1).enabled = frmLogin.Livello <> LVL_USER
    DataGrid1.AllowUpdate = frmLogin.Livello <> LVL_USER
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case TXT_COD_UTENTE
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    End Select
End Sub


Private Sub txt_LostFocus(Index As Integer)
    On Error Resume Next
    If (Index = TXT_COD_UTENTE) Then caricaDatiUtente Val(txt(TXT_COD_UTENTE))
End Sub

Private Sub txtNrCopieFormulario_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub

Private Sub txtNrCopiePesa_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub

Private Sub txtNrCopieRegistro_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
     End If
End Sub


Private Sub caricaDatiUtente(codUtente As Long)
    Dim rs As Recordset
    
    
    Label1(4).Caption = ""
    Label1(6).Caption = ""
    
    If codUtente = 0 Then Exit Sub
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Source = "select utenti.* from Utenti  LEFT JOIN Comuni ON Utenti.CodComuneResidenza = Comuni.CODISTAT  where utenti.codUtente=" & codUtente
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
            
    If rs.EOF Then
        MsgBox "Attenzione CodUtente:" & codUtente & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
    Else
        
        Label1(4).Caption = IIf(nullToBlank(rs!ragioneSociale) = vbNullString, nullToBlank(rs!Nome) & " " & nullToBlank(rs!Cognome), nullToBlank(rs!ragioneSociale))
        Label1(6).Caption = nullToBlank(rs!IndirizzoResidenza) & ", " & nullToBlank(rs!NumCivicoResidenza)
        DoEvents
        
    End If
    
    rs.Close
    
    caricaImmobili codUtente
    
End Sub


Private Sub caricaImmobili(codUtente As Long)
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim noCheck As Boolean
    
    lvwImmobili.ListItems.Clear
    
    If codUtente = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    Set rs = gDBC.Execute("SELECT Comuni.comune, Tariffe.Descrizione AS Tariffa, Immobili.* " & _
            "FROM (immobili INNER JOIN Comuni ON immobili.CodComune = Comuni.CODISTAT) INNER JOIN Tariffe ON immobili.CodTariffa = Tariffe.CodTariffa " & _
            "WHERE  DATE() between DataInizioValidita and DataFineValidita and   immobili.Tipo<>'Z' AND immobili.[codUtente]=" & codUtente & " " & _
            "ORDER BY immobili.FlgUtenzaDomesticaResidente")

    noCheck = False
   lvwImmobili.Tag = ""
    On Error Resume Next
    Do While Not rs.EOF
        Set item = lvwImmobili.ListItems.Add(, "_" & rs!codImmobile, rs!codImmobile)
        item.ListSubItems.Add , rs!tipo.Name, rs!tipo.value
        item.ListSubItems.Add , rs!Comune.Name, rs!Comune.value
        item.ListSubItems.Add , rs!TipoUtenza.Name, rs!TipoUtenza.value
        item.ListSubItems.Add , rs!Indirizzo.Name, rs!Indirizzo.value & "," & rs!NumCivico.value
        item.ListSubItems.Add , rs!Subalterno.Name, nullToBlank(rs!Subalterno.value)
        item.ListSubItems.Add , rs!NumInterno.Name, nullToBlank(rs!NumInterno.value)
        item.ListSubItems.Add , rs!codTariffa.Name, nullToBlank(rs!codTariffa.value)
        item.ListSubItems.Add , rs!Tariffa.Name, nullToBlank(rs!Tariffa.value)
        
        If ((rs!flgUtenzaDomesticaResidente Or _
            lvwImmobili.ListItems.Count = 1 Or _
             gCodImmobileMovManuale = rs!codImmobile) And Not noCheck) Then
            
 
            lvwImmobili_ItemCheck item
            item.Checked = True
            
            noCheck = gCodCausaleMovManuale = item.key
        End If
        
        rs.MoveNext
    Loop
    
    
    AllargaLeColonneLVW lvwImmobili, False
    Screen.MousePointer = vbDefault
    lvwImmobili.ColumnHeaders(2).width = 0
    
    
    rs.Close
    Set rs = Nothing
End Sub



Private Sub lvwImmobili_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If lvwImmobili.Tag <> vbNullString And lvwImmobili.ListItems.Count > 1 Then lvwImmobili.ListItems(lvwImmobili.Tag).Checked = False
    
    lvwImmobili.Tag = item.key
    
    item.Selected = True
    
    mCodImmobile = item.key
End Sub



Private Sub caricaBonusPremi()
    Dim item As ListItem
    
    lvwBonus.ListItems.Clear
    lvwPremi.ListItems.Clear
    
    Dim rs As ADODB.Recordset
    Set rs = gDBC.Execute("select * from TipiAnagrafe")
    
    Do While Not rs.EOF
        Set item = lvwBonus.ListItems.Add(, rs!TipoAnagrafe, rs!Descrizione)
        item.Checked = IIf(rs!AttivoBonus, True, False)
    
        Set item = lvwPremi.ListItems.Add(, rs!TipoAnagrafe, rs!Descrizione)
        item.Checked = IIf(rs!AttivoPremio, True, False)
    
        rs.MoveNext
    Loop
    rs.Close
End Sub


Private Sub SalvaBonusPremi()
    Dim item As ListItem
    Dim chkPremio As Long
    Dim chkBonus As Long
   
    
    For Each item In lvwBonus.ListItems
          
        
        chkPremio = IIf(lvwPremi.ListItems(item.Index).Checked, -1, 0)
        chkBonus = IIf(item.Checked, -1, 0)
        gDBC.Execute "update TipiAnagrafe set AttivoBonus= " & chkBonus & ", AttivoPremio=" & chkPremio & " where TipoAnagrafe='" & item.key & "'"
        
    Next
End Sub
