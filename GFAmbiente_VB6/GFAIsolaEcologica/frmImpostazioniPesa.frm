VERSION 5.00
Begin VB.Form frmImpostazioniPesa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impostazioni Pesa"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7020
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      Top             =   0
      Width           =   9975
      Begin VB.Frame Frame2 
         Caption         =   "Struttura del Messaggio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   3480
         TabIndex        =   45
         Top             =   240
         Width           =   6135
         Begin VB.ComboBox cmbSetting 
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
            Index           =   9
            ItemData        =   "frmImpostazioniPesa.frx":0000
            Left            =   4800
            List            =   "frmImpostazioniPesa.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetting 
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
            Index           =   8
            ItemData        =   "frmImpostazioniPesa.frx":0016
            Left            =   3240
            List            =   "frmImpostazioniPesa.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetting 
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
            ItemData        =   "frmImpostazioniPesa.frx":0042
            Left            =   1680
            List            =   "frmImpostazioniPesa.frx":0052
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetting 
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
            ItemData        =   "frmImpostazioniPesa.frx":006E
            Left            =   120
            List            =   "frmImpostazioniPesa.frx":007B
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   56
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "CR"
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
            Left            =   4800
            TabIndex        =   54
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   53
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caratt. Iniziale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "STX"
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
            Left            =   1680
            TabIndex        =   48
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "ETX"
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
            Left            =   3240
            TabIndex        =   47
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label3 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   46
            Top             =   480
            Width           =   135
         End
      End
      Begin VB.TextBox Text1 
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
         Index           =   11
         Left            =   3480
         TabIndex        =   44
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkDoppiaPesata 
         Alignment       =   1  'Right Justify
         Caption         =   "Utilizzo di 2 pese come se fosse 1"
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
         Left            =   6600
         TabIndex        =   43
         Top             =   2520
         Width           =   3015
      End
      Begin VB.CheckBox chkEcopunto 
         Alignment       =   1  'Right Justify
         Caption         =   "Pesa per ECO Punto"
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
         Left            =   6600
         TabIndex        =   42
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox chkRilevazionePesoAuto 
         Alignment       =   1  'Right Justify
         Caption         =   "Rilevazione Peso Automatica"
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
         Left            =   6600
         TabIndex        =   41
         ToolTipText     =   "Se impostato il sistema in automatico rileverà ogni variazione di peso"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.CheckBox chkBadgeDaPesata 
         Alignment       =   1  'Right Justify
         Caption         =   "Badge fornito con la Pesata"
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
         Left            =   6600
         TabIndex        =   40
         ToolTipText     =   "se impostato indica che nei dati della pesata vi è anche l'informazione del Badge"
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   2280
         TabIndex        =   38
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   9
         Left            =   2280
         TabIndex        =   37
         Top             =   5520
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   120
         TabIndex        =   35
         Top             =   6240
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   120
         TabIndex        =   33
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CheckBox chkattiva 
         Alignment       =   1  'Right Justify
         Caption         =   "Attiva"
         DownPicture     =   "frmImpostazioniPesa.frx":008E
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   6720
         Picture         =   "frmImpostazioniPesa.frx":04C1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4560
         Width           =   3015
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Chiudi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   8280
         TabIndex        =   27
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Salva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   6720
         TabIndex        =   26
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkCheckSum 
         Alignment       =   1  'Right Justify
         Caption         =   "Controlla CheckSum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3480
         TabIndex        =   12
         Top             =   6360
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         ItemData        =   "frmImpostazioniPesa.frx":0911
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":0933
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         ItemData        =   "frmImpostazioniPesa.frx":0975
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":0982
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         ItemData        =   "frmImpostazioniPesa.frx":09A0
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":09B6
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         ItemData        =   "frmImpostazioniPesa.frx":09E2
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":09F5
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4200
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         ItemData        =   "frmImpostazioniPesa.frx":0A08
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":0A15
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5640
         Width           =   2535
      End
      Begin VB.ComboBox cmbSetting 
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
         Index           =   6
         ItemData        =   "frmImpostazioniPesa.frx":0A24
         Left            =   3480
         List            =   "frmImpostazioniPesa.frx":0A31
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4920
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Verifica Cavo Seriale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6600
         TabIndex        =   11
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Peso Stabile"
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
         Left            =   3480
         TabIndex        =   39
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CMD Tara"
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
         TabIndex        =   36
         Top             =   6000
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CMD Scrittura AlibiM / Ret.Code"
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
         TabIndex        =   34
         Top             =   5280
         Width           =   2820
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CMD Lettura Peso"
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
         TabIndex        =   32
         Top             =   4560
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CMD Scelta Bilancia / Ret. Code"
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
         TabIndex        =   30
         Top             =   3840
         Width           =   2850
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome PC"
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
         TabIndex        =   25
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo IP "
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
         TabIndex        =   24
         Top             =   2400
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modello"
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
         TabIndex        =   23
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porta Seriale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   3480
         TabIndex        =   19
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parità"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   3480
         TabIndex        =   18
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bit per Secondo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   17
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bit di Dati"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   3480
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bit di Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   3480
         TabIndex        =   15
         Top             =   5400
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controllo Flusso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   3480
         TabIndex        =   14
         Top             =   4680
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmImpostazioniPesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmd_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
            Screen.MousePointer = vbHourglass
            DoEvents
            UpdateRow
        
            InitPese False, Nothing
            
            If Not gDisabilitaNotifiche Then MsgBox "Operazione Riuscita", vbExclamation + vbOKOnly, App.Title
            
            Screen.MousePointer = vbDefault
            DoEvents

    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    g_rsPese.MoveFirst
    
    For i = cmbSetting.LBound To cmbSetting.ubound
        cmbSetting(i).ListIndex = 0
    Next
    
    Do While Not g_rsPese.EOF
        List1.AddItem (g_rsPese!codPesa + " - " + g_rsPese!Descrizione)
        g_rsPese.MoveNext
    Loop
    
    applicaRestrizioniUtente
    
    Text1(0).Locked = True
    Text1(1).Locked = True
    Text1(2).Locked = True
    Text1(3).Locked = True
    Text1(4).Locked = True
End Sub


Public Sub applicaRestrizioniUtente()
    cmd(1).enabled = frmLogin.Livello <> LVL_USER
End Sub


Private Sub List1_Click()
            
    g_rsPese.MoveFirst
    g_rsPese.Find "CodPesa=" & left$(List1.Text, 2), , adSearchForward
            
    Text1(0).Text = g_rsPese!codPesa
    Text1(1).Text = g_rsPese!Descrizione
    Text1(2).Text = g_rsPese!Modello
    Text1(3).Text = nullToBlank(g_rsPese!IPAddress)
    Text1(4).Text = nullToBlank(g_rsPese!PCNAME)
    
    chkEcoPunto.Value = IIf(g_rsPese!flgEcopunto, 1, 0)
    chkDoppiaPesata.Value = IIf(g_rsPese!MSGPerDoppiaPesa, 1, 0)
    chkCheckSum.Value = IIf(g_rsPese!CheckSum, 1, 0)
    Check1.Value = IIf(g_rsPese!CheckPort, 1, 0)
    chkattiva.Value = IIf(g_rsPese!Active, 1, 0)
    
    cmbSetting(0).Text = IIf(g_rsPese!StartChar = "N", "Nessuno", g_rsPese!StartChar)
    cmbSetting(1).Text = "Com" & g_rsPese!ComID
    
    cmbSetting(7).Text = IIf(g_rsPese!STX = "N", "Nessuno", g_rsPese!STX)
    cmbSetting(8).Text = IIf(g_rsPese!ETX = "N", "Nessuno", g_rsPese!ETX)
    cmbSetting(9).Text = IIf(g_rsPese!CR = True, "SI", "NO")
    
    chkBadgeDaPesata.Value = IIf(g_rsPese!BadgeInStringaPesata, 1, 0)
    chkRilevazionePesoAuto.Value = IIf(g_rsPese!ModalitaLetturaPesoAutomatica, 1, 0)
    
    Text1(5).Text = nullToBlank(g_rsPese!CmdSceltaPesa)
    Text1(6).Text = nullToBlank(g_rsPese!cmdLetturaPeso)
    Text1(7).Text = nullToBlank(g_rsPese!CmdScritturaAlibiM)
    Text1(8).Text = nullToBlank(g_rsPese!CmdTara)
    Text1(9).Text = nullToBlank(g_rsPese!CmdScritturaAlibiMRetERRCode)
    Text1(10).Text = nullToBlank(g_rsPese!CmdSceltaPesaRetOKCode)
    Text1(11).Text = nullToBlank(g_rsPese!cmdPesoStabileCode)
    
    
    If g_rsPese!ParityBit = "E" Then
        cmbSetting(2).Text = "Parità"
    ElseIf g_rsPese!ParityBit = "O" Then
        cmbSetting(2).Text = "Dispari"
    ElseIf g_rsPese!ParityBit = "N" Then
        cmbSetting(2).Text = "Nessuna"
    End If
    
    cmbSetting(3).Text = g_rsPese!BaudeRate
    cmbSetting(4).Text = g_rsPese!DataBits
    cmbSetting(5).Text = g_rsPese!StopBit
    cmbSetting(6).ListIndex = g_rsPese!XonXoff
    
    
    
    

End Sub

Private Sub UpdateRow()
    If List1.ListIndex = -1 Then Exit Sub
    
    g_rsPese.MoveFirst
    g_rsPese.Find "CodPesa=" & left$(List1.Text, 2), , adSearchForward
            
    g_rsPese!codPesa = Text1(0).Text
    g_rsPese!Descrizione = Text1(1).Text
    g_rsPese!Modello = Text1(2).Text
    g_rsPese!IPAddress = IIf(Text1(3).Text = "", Null, Text1(3).Text)
    g_rsPese!PCNAME = IIf(Text1(4).Text = "", Null, Text1(4).Text)
    
    g_rsPese!flgEcopunto = IIf(chkEcoPunto.Value = 1, True, False)
    g_rsPese!MSGPerDoppiaPesa = IIf(chkDoppiaPesata.Value = 1, True, False)
    g_rsPese!CheckSum = IIf(chkCheckSum.Value = 1, True, False)
    g_rsPese!CheckPort = IIf(Check1.Value = 1, True, False)
    g_rsPese!Active = IIf(chkattiva.Value = 1, True, False)
    
    g_rsPese!StartChar = left$(cmbSetting(0).Text, 1)
    g_rsPese!STX = IIf(cmbSetting(7).Text = "Nessuno", "N", cmbSetting(7).Text)
    g_rsPese!ETX = IIf(cmbSetting(8).Text = "Nessuno", "N", cmbSetting(8).Text)
    g_rsPese!CR = IIf(cmbSetting(9).Text = "SI", True, False)
    
    
    g_rsPese!ComID = Right$(cmbSetting(1).Text, 1)
    
    
    If cmbSetting(2).Text = "Parità" Then
        g_rsPese!ParityBit = "E"
    ElseIf cmbSetting(2).Text = "Dispari" Then
        g_rsPese!ParityBit = "O"
    ElseIf cmbSetting(2).Text = "Nessuna" Then
        g_rsPese!ParityBit = "N"
    End If
    
    g_rsPese!BaudeRate = cmbSetting(3).Text
    g_rsPese!DataBits = cmbSetting(4).Text
    g_rsPese!StopBit = cmbSetting(5).Text
    g_rsPese!XonXoff = cmbSetting(6).ListIndex
    
    
    g_rsPese!BadgeInStringaPesata = IIf(chkBadgeDaPesata.Value, True, False)
    g_rsPese!ModalitaLetturaPesoAutomatica = IIf(chkRilevazionePesoAuto.Value, True, False)
    
    
    g_rsPese!CmdSceltaPesa = IIf(Text1(5).Text = "", Null, Text1(5).Text)
    g_rsPese!cmdLetturaPeso = IIf(Text1(6).Text = "", Null, Text1(6).Text)
    g_rsPese!CmdScritturaAlibiM = IIf(Text1(7).Text = "", Null, Text1(7).Text)
    g_rsPese!CmdTara = IIf(Text1(8).Text = "", Null, Text1(8).Text)
    g_rsPese!CmdScritturaAlibiMRetERRCode = IIf(Text1(9).Text = "", Null, Text1(9).Text)
    g_rsPese!CmdSceltaPesaRetOKCode = IIf(Text1(10).Text = "", Null, Text1(10).Text)
    g_rsPese!cmdPesoStabileCode = IIf(Text1(11).Text = "", Null, Text1(11).Text)
    
    
    Set g_rsPese.ActiveConnection = gDBC
    g_rsPese.UpdateBatch adAffectAllChapters
    Set g_rsPese.ActiveConnection = Nothing
End Sub

