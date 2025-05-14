VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUscita 
   Caption         =   "Uscita Rifiuti"
   ClientHeight    =   10080
   ClientLeft      =   5445
   ClientTop       =   3735
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   0
      Left            =   240
      ScaleHeight     =   6975
      ScaleWidth      =   17415
      TabIndex        =   24
      Top             =   360
      Width           =   17415
      Begin VB.Frame Frame3 
         Caption         =   "Destinatario/Trasportatore"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   11415
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   6240
            TabIndex        =   1
            ToolTipText     =   "Codice dell'utente rilevato"
            Top             =   480
            Width           =   1215
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
            Height          =   600
            Index           =   0
            Left            =   7560
            Picture         =   "frmUscita.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cmbTarga 
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
            Left            =   10080
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1440
            Width           =   1935
         End
         Begin VB.ComboBox cmbConducente 
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
            Left            =   6240
            TabIndex        =   3
            Top             =   1440
            Width           =   3735
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   120
            TabIndex        =   0
            ToolTipText     =   "Codice dell'utente rilevato"
            Top             =   480
            Width           =   1215
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
            Height          =   600
            Index           =   1
            Left            =   1440
            Picture         =   "frmUscita.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cmbImpianto 
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
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Impianto di destinazione"
            Top             =   1440
            Width           =   6015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trasportatore"
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
            Height          =   225
            Index           =   26
            Left            =   6240
            TabIndex        =   44
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partita IVA"
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
            Height          =   225
            Index           =   4
            Left            =   9480
            TabIndex        =   43
            Top             =   720
            Width           =   825
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
            Index           =   5
            Left            =   9480
            TabIndex        =   42
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice Ficale"
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
            Height          =   225
            Index           =   2
            Left            =   8280
            TabIndex        =   41
            Top             =   720
            Width           =   1125
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
            Index           =   3
            Left            =   8280
            TabIndex        =   40
            Top             =   945
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   8280
            TabIndex        =   39
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ragioneso"
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
            Index           =   1
            Left            =   8280
            TabIndex        =   38
            Top             =   465
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Targa"
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
            Height          =   225
            Index           =   33
            Left            =   10080
            TabIndex        =   37
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conducente"
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
            Height          =   225
            Index           =   22
            Left            =   6240
            TabIndex        =   36
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinatario"
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
            Height          =   225
            Index           =   16
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1005
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   15
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ragioneso"
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
            Index           =   14
            Left            =   2160
            TabIndex        =   32
            Top             =   465
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partita IVA"
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
            Height          =   225
            Index           =   13
            Left            =   3360
            TabIndex        =   31
            Top             =   720
            Width           =   825
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
            Index           =   12
            Left            =   3360
            TabIndex        =   30
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice Ficale"
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
            Height          =   225
            Index           =   11
            Left            =   2160
            TabIndex        =   29
            Top             =   720
            Width           =   1125
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
            Index           =   10
            Left            =   2160
            TabIndex        =   28
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impianto"
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
            Height          =   225
            Index           =   9
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2895
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   16095
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   1
            Left            =   8520
            TabIndex        =   67
            ToolTipText     =   "Numero del formulario"
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   660
            Index           =   4
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0"
            ToolTipText     =   "Peso rilevato"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   675
            Index           =   3
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0"
            ToolTipText     =   "Peso rilevato"
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "LED BOARD REVERSED"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   690
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Text            =   "0"
            ToolTipText     =   "Peso rilevato"
            Top             =   480
            Width           =   2775
         End
         Begin MSComctlLib.ListView lvwLotti 
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
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
            NumItems        =   0
         End
         Begin MSComCtl2.DTPicker dtpickFormulario 
            Height          =   645
            Left            =   11400
            TabIndex        =   68
            ToolTipText     =   "Data formulario"
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1138
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16580609
            CurrentDate     =   39352
         End
         Begin VB.Label lblCambiaPesa 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cambia Pesa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1680
            MouseIcon       =   "frmUscita.frx":0ED4
            MousePointer    =   99  'Custom
            TabIndex        =   74
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Formulario"
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
            Height          =   225
            Index           =   17
            Left            =   11400
            TabIndex        =   70
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num.Formulario"
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
            Height          =   225
            Index           =   24
            Left            =   8520
            TabIndex        =   69
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Differenza"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   8
            Left            =   3120
            TabIndex        =   49
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Giacenza Totale"
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
            Height          =   225
            Index           =   7
            Left            =   5520
            TabIndex        =   48
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Pesata"
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
            Height          =   225
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Index           =   1
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   12015
      TabIndex        =   51
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame5 
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4215
         Left            =   240
         TabIndex        =   57
         Top             =   3720
         Width           =   11055
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
            Height          =   885
            Index           =   6
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   15
            ToolTipText     =   "Caratteristiche chimiche fisiche del rifiuto"
            Top             =   1200
            Width           =   4095
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
            Height          =   400
            Index           =   9
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Caratteristiche chimiche fisiche del rifiuto"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Index           =   10
            Left            =   4320
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Annotazioni da riportare come commento"
            Top             =   2400
            Width           =   3015
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
            Index           =   2
            Left            =   1440
            Picture         =   "frmUscita.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   11
            Left            =   4560
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Lougo di Produzione e Attività di Provenienza del Rifiuto"
            Top             =   1200
            Width           =   3015
         End
         Begin MSComctlLib.ListView lvwClassiPericolosita 
            Height          =   1695
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Selezionare i codci di classi di pericolosità del rifiuto"
            Top             =   2400
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2990
            View            =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrizione"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caratteristiche chimico fisiche"
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
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classi di pericolosità"
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
            Index           =   25
            Left            =   120
            TabIndex        =   63
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intermediario"
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
            Index           =   27
            Left            =   120
            TabIndex        =   62
            ToolTipText     =   "Inserire il codice dell'intermediario / commerciante"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Annotazioni Registro"
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
            Index           =   28
            Left            =   4320
            TabIndex        =   61
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ragioneso"
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
            Index           =   29
            Left            =   2040
            TabIndex        =   60
            Top             =   465
            Width           =   1155
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
            Index           =   30
            Left            =   2040
            TabIndex        =   59
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Luogo di Prod. / Attività di Proven. "
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
            Index           =   31
            Left            =   4560
            TabIndex        =   58
            Top             =   960
            Width           =   2745
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formulario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3375
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   11655
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   6360
            TabIndex        =   11
            ToolTipText     =   "Inserire Numero Colli"
            Top             =   2880
            Width           =   2295
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Index           =   8
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Annotazioni da riportare come commento"
            Top             =   2160
            Width           =   6015
         End
         Begin VB.Frame Frame4 
            Caption         =   "Normativa ADR / RID"
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
            Height          =   615
            Left            =   6360
            TabIndex        =   12
            ToolTipText     =   "Trasporto sottoposto a normativa ADR / RID"
            Top             =   2040
            Width           =   2175
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   315
               Left            =   75
               ScaleHeight     =   315
               ScaleWidth      =   1890
               TabIndex        =   71
               Top             =   225
               Width           =   1890
               Begin VB.OptionButton optNormativa 
                  Caption         =   "Si"
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
                  Index           =   0
                  Left            =   300
                  TabIndex        =   73
                  Top             =   0
                  Width           =   615
               End
               Begin VB.OptionButton optNormativa 
                  Caption         =   "No"
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
                  Index           =   1
                  Left            =   1020
                  TabIndex        =   72
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   735
               End
            End
         End
         Begin VB.ComboBox cmbDestRifiuto 
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
            ItemData        =   "frmUscita.frx":1790
            Left            =   120
            List            =   "frmUscita.frx":1792
            TabIndex        =   9
            Text            =   "cmbDestRifiuto"
            ToolTipText     =   "Destinazione del Rifiuto"
            Top             =   1410
            Width           =   3615
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   3000
            TabIndex        =   8
            ToolTipText     =   "Va indicato se diverso da il più breve"
            Top             =   600
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker DTPickDataOraTrasporto 
            Height          =   450
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Data formulario"
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   794
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yy HH:mm:ss"
            Format          =   16580611
            CurrentDate     =   39352
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero Colli"
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
            Index           =   32
            Left            =   6360
            TabIndex        =   66
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Annotazioni Formulario"
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
            Index           =   21
            Left            =   120
            TabIndex        =   56
            Top             =   1920
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data/Ora Trasporto"
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
            Index           =   23
            Left            =   105
            TabIndex        =   55
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinazione Rifiuto"
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
            Index           =   18
            Left            =   120
            TabIndex        =   54
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Percorso"
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
            Index           =   20
            Left            =   3000
            TabIndex        =   53
            Top             =   360
            Width           =   750
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUscita.frx":1794
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   2
      Left            =   9360
      ScaleHeight     =   1095
      ScaleWidth      =   1815
      TabIndex        =   65
      Top             =   7560
      Width           =   1815
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   975
      Left            =   7680
      TabIndex        =   50
      Top             =   7320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1720
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesata"
            Key             =   "PESATA"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formulario / Registro"
            Key             =   "FORMULARIOREGISTRO"
            Object.ToolTipText     =   "Dati aggiuntivi per la stampa del formulario  e registro di scarico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Movimenti"
            Key             =   "MOVIMENTI"
            Object.ToolTipText     =   "Visuliazza l'elenco dei movimenti in Uscita"
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
   Begin VB.CommandButton cmd 
      Caption         =   "Dettaglio "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6360
      TabIndex        =   22
      Top             =   9240
      Width           =   1215
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
      Height          =   735
      Index           =   0
      Left            =   11280
      TabIndex        =   19
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Lettura Peso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7680
      TabIndex        =   21
      Top             =   9240
      Width           =   1695
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
      Height          =   735
      Index           =   2
      Left            =   9600
      TabIndex        =   20
      Top             =   9240
      Width           =   1215
   End
End
Attribute VB_Name = "frmUscita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private mFirstTime As Boolean

Private Const TXT_COD_TRASPORTATORE = 0
Private Const TXT_NR_FORMULARIO = 1
Private Const TXT_QTY_PESA = 2
Private Const TXT_QTY_GIACENZA = 3
Private Const TXT_DIFF_PESO = 4
Private Const TXT_COD_DESTINATARIO = 5
Private Const TXT_CARATTER_CHIMICHE_FISICHE = 6
Private Const TXT_PERCORSO = 7
Private Const TXT_ANNOTAZIONI_FORMULARIO = 8
Private Const TXT_COD_INTERMEDIARIO = 9
Private Const TXT_ANNOTAZIONI_REGISTRO = 10
Private Const TXT_LUOGO_PROVENIENZA = 11
Private Const TXT_NUMCOLLI = 12


Private Const LBL_TRASPOR_RAGIONESOCIALE = 1
Private Const LBL_TRASPOR_CODICEFISCALE = 3
Private Const LBL_TRASPOR_PIVA = 5

Private Const LBL_DESTINAT_RAGIONESOCIALE = 14
Private Const LBL_DESTINAT_CODCEFISCALE = 10
Private Const LBL_DESTINAT_PIVA = 12

Private Const LBL_INTERMEDIARIO_RAGIONESOCIALE = 29



Private Const CMD_ESCI = 0
Private Const CMD_SALVA = 2
Private Const CMD_LETTURAPESO = 1
Private Const CMD_DETTAGLIOMOV = 3

Private Esci As Boolean
Private msgPesa As messaggioPesa
Private midMovimentoUscita As Long
Private midDocumentoPesata As Long
Private mnrDocumentoPesata As Long
Private midMovimentiDaStampare As String
Private mRichiestaPesoInCorso As Boolean
Private mPesoLettoDaBilancia As Boolean
Private mCodComuneCorrente As String


Private Sub cmbRicerca_Click(Index As Integer)
    Select Case Index
        Case 0
            frmRicercaUtente.schowMe txt(TXT_COD_TRASPORTATORE), "T"
            caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE).Text
        Case 1
            frmRicercaUtente.schowMe txt(TXT_COD_DESTINATARIO), "D"
            caricaDatiDestinatario txt(TXT_COD_DESTINATARIO).Text
        Case 2
            frmRicercaUtente.schowMe txt(TXT_COD_INTERMEDIARIO), "I"
            caricaIntermediario txt(TXT_COD_INTERMEDIARIO).Text
            
    End Select
End Sub

Private Sub operazioniDiCaricoScarico(codLotto As String, codCategoria As String, _
        codSottoCategoria As String, codCer As String, UnitaMisura As String, _
        tipoOpeSmaRec As String, quantita As Double, classiPericolosita As String)
    Dim rs As ADODB.Recordset
    Dim nrRegistro As Long
    Dim idRegistro As Long
    Dim qta As Double
    Dim dataOpe As String
    Dim colRif As Collection
    
    Set colRif = New Collection
    On Error GoTo operazioniDiCaricoScaricoERR
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    rs.Source = "SELECT Lotti.ID, LegameLottoRegistrazioneCarico.idRegistro" & _
                " FROM Lotti INNER JOIN LegameLottoRegistrazioneCarico ON Lotti.CodLotto = LegameLottoRegistrazioneCarico.codLotto " & _
                " WHERE Lotti.id in (" & codLotto & ")"

    Set rs.ActiveConnection = gDBC
    nrRegistro = -1
    rs.Open
    dataOpe = Format$(Now, "yyyy-MM-dd")
    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        colRif.Add rs!idRegistro.value
        rs.MoveNext
    Loop
    
    
    nrRegistro = getNumeratore(REGISTRO, False)
    'scarico
    idRegistro = creaVoceDiRegistro(nrRegistro, "S", dataOpe, txt(TXT_NR_FORMULARIO).Text, Format$(dtpickFormulario.value, "yyyy-MM-dd") _
        , codCategoria, codSottoCategoria, codCer, tipoOpeSmaRec, UnitaMisura, quantita, txt(TXT_LUOGO_PROVENIENZA).Text, Val(txt(TXT_COD_INTERMEDIARIO).Text), txt(TXT_ANNOTAZIONI_REGISTRO).Text, Now, frmLogin.codUser, classiPericolosita)
    
    If idRegistro <> -1 Then
    
        salvaRiferimentiDiCaricoSulRegistro idRegistro, colRif
                
        'stampa formulario
        stampaFormulario idRegistro, codCategoria, codSottoCategoria, codCer
        
        'crea associazione lotto uscita con lotti ingresso , svuotando il lotto ingresso
        salvaLottoUscitaLottoIngressi idRegistro, codCategoria, codSottoCategoria, codCer, quantita
            
    End If
    
    
    
    Exit Sub
operazioniDiCaricoScaricoERR:
    Err.Raise 9999, "OperazioniDiCaricoScarico", Err.Description
    
End Sub


Public Sub StampaMovimenti()
    
    'aggiorna riferimento sulla tabella documenti
    AggiornaRiferimentoInDocumento midDocumentoPesata, mnrDocumentoPesata, midMovimentiDaStampare, "I"
    
    ReportPesataUscita.Stampa midMovimentiDaStampare, pfrmUscita, CBool(gStampaPesataAnteprima), gStampaPesataNrCopie
    
    Unload ReportPesataUscita
    
    
    midDocumentoPesata = 0
    mnrDocumentoPesata = 0
    midMovimentiDaStampare = vbNullString
    
    'annullo lettura pesata
    If Not getCurrentPesa.ModalitaLetturaPesoAutomatica And cmd(CMD_LETTURAPESO).Caption = "Annulla Pesata" Then cmd_Click (CMD_LETTURAPESO)
    DoEvents
    
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim idMovimento As Long
    Dim idAlibiMemory As String
    
    Select Case Index
        Case CMD_ESCI
            Unload Me
        Case CMD_SALVA
            If (validaDati(idAlibiMemory)) Then
                
                If salva(False, idMovimento, idAlibiMemory) Then
                    
                    
                    midMovimentiDaStampare = midMovimentiDaStampare & idMovimento & ","
                    StampaMovimenti
                    
                    txt(TXT_COD_TRASPORTATORE).Text = ""
                    txt(TXT_NR_FORMULARIO).Text = ""
                    txt(TXT_QTY_PESA).Text = ""
                    txt(TXT_QTY_GIACENZA).Text = ""
                    txt(TXT_DIFF_PESO).Text = ""
                    txt(TXT_COD_DESTINATARIO).Text = ""
                    txt(TXT_CARATTER_CHIMICHE_FISICHE).Text = ""
                    txt(TXT_PERCORSO).Text = ""
                    txt(TXT_ANNOTAZIONI_FORMULARIO).Text = ""
                    txt(TXT_ANNOTAZIONI_REGISTRO).Text = ""
                    txt(TXT_COD_INTERMEDIARIO).Text = ""
                    txt(TXT_LUOGO_PROVENIENZA).Text = getProvenienzaRifiuti()
                    txt(TXT_NUMCOLLI).Text = ""
                    
                    caricaDatiDestinatario ""
                    
                    caricaDatiTrasportatore ""
                    
                    caricaIntermediario ""
                    
                    caricaClassiDiPericolosita ""
                    
                    If gDisabilitaNotifiche = 0 Then MsgBox "Operazione di scarico avvenuta con successo !!!", vbExclamation + vbOKOnly, App.Title
                    
                    MDIMain.refreshForm
                    
                    TabStrip.Tabs(1).Selected = True
                    TabStrip_Click
                End If
            End If
        Case CMD_LETTURAPESO
            If cmd(CMD_LETTURAPESO).Caption = "Annulla Pesata" Then
                cmd(CMD_LETTURAPESO).Caption = "Lettura Peso"
                cmd(CMD_ESCI).enabled = True
                cmd(CMD_SALVA).enabled = True
                picContainer(0).enabled = True
                Esci = True
            Else
                Esci = False
                cmd(CMD_ESCI).enabled = False
                cmd(CMD_SALVA).enabled = False
                picContainer(0).enabled = False
                cmd(CMD_LETTURAPESO).Caption = "Annulla Pesata"
                
                letturaPeso gPesaIsolaCod
            End If
        Case CMD_DETTAGLIOMOV
            If Not lvwLotti.SelectedItem Is Nothing Then
                lvwLotti_DblClick
            End If

    End Select
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mFirstTime Then
        cmbImpianto.SetFocus
        txt(TXT_QTY_PESA).SetFocus
        mFirstTime = False
    End If
    
    
End Sub

Private Sub Form_Load()

    TabStrip.Tabs.Remove 3
    Dim colH As ColumnHeader
    mFirstTime = True
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
    Label1(LBL_TRASPOR_CODICEFISCALE).Caption = ""
    Label1(LBL_TRASPOR_PIVA).Caption = ""

    Label1(LBL_DESTINAT_RAGIONESOCIALE).Caption = ""
    Label1(LBL_DESTINAT_CODCEFISCALE).Caption = ""
    Label1(LBL_DESTINAT_PIVA).Caption = ""

    Label1(LBL_INTERMEDIARIO_RAGIONESOCIALE).Caption = ""

    mCodComuneCorrente = GetCurrentComune!codistat
    
    With lvwLotti
        Set colH = .ColumnHeaders.Add(, , "Cod Cer")
        Set colH = .ColumnHeaders.Add(, , "Tip Rifiuto")
        
        Set colH = .ColumnHeaders.Add(, , "Qty In Gia")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Da Reg")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Da Sca")
        colH.Alignment = lvwColumnRight
        
        Set colH = .ColumnHeaders.Add(, , "Pericolo")
        
        Set colH = .ColumnHeaders.Add(, , "Qty Max")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Max Normat")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Rit")
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Rit Norm")
        Set colH = .ColumnHeaders.Add(, , "Dta Primo Mov")
    
    End With

    cmbTarga.Clear
'    cmbTarga.Text = ""
     
     
    caricaClassiDiPericolosita vbNullString
     
    caricaTotali
    
    caricaOperazioniSuRifiuti
    
    DTPickDataOraTrasporto.value = Now
    dtpickFormulario.value = Now
    
    TabStrip_Click
    
    txt(TXT_LUOGO_PROVENIENZA).Text = getProvenienzaRifiuti()
    
    lblCambiaPesa.Visible = getCountPese > 1

End Sub

Public Sub ShowME()
    
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        Me.Show vbModal
    Else
        Me.Show
    End If
End Sub

Private Function getProvenienzaRifiuti() As String
    Dim rs As Recordset
    
    getProvenienzaRifiuti = ""
    
    Set rs = gDBC.Execute("select societa.RagioneSociale ,impianti.Cap,impianti.Indirizzo,impianti.NumCivico,impianti.SubAlterno ,impianti.Interno  " & _
        " from (impianti inner join societa on " & _
        "societa.idSocieta = impianti.idsoggetto and  societa.codSoggetto = impianti.codSoggetto)" & _
        "where idImpianto=" & gIdImpiantoProduttore)
    If Not rs.EOF Then
        getProvenienzaRifiuti = rs!ragioneSociale & " " & rs!Cap & " " & rs!Indirizzo & _
            IIf(IsNull(rs!NumCivico), "", ", " & nullToBlank(rs!NumCivico)) & _
            IIf(IsNull(rs!Subalterno), "", ", " & nullToBlank(rs!Subalterno))
    End If
    
    rs.Close
    Set rs = Nothing

End Function

Private Sub caricaOperazioniSuRifiuti()
    cmbDestRifiuto.Clear
    
    If DevEnvMDB.rscmdOperazioniSuRifiuti.state = adStateClosed Then
        DevEnvMDB.cmdOperazioniSuRifiuti
    End If
    
    DevEnvMDB.rscmdOperazioniSuRifiuti.MoveFirst
    Do While Not DevEnvMDB.rscmdOperazioniSuRifiuti.EOF
    
        cmbDestRifiuto.AddItem DevEnvMDB.rscmdOperazioniSuRifiuti!codtipo & " - " & DevEnvMDB.rscmdOperazioniSuRifiuti!Descrizione
        DevEnvMDB.rscmdOperazioniSuRifiuti.MoveNext
    Loop
    If cmbDestRifiuto.ListCount > 0 Then cmbDestRifiuto.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    
    resizeControl
    
    
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
    
    TabStrip.Move offsetX, offsetY, Me.ScaleWidth - offsetX * 2, cmd(0).top - offsetY * 2
    
    For i = picContainer.LBound To picContainer.ubound
        picContainer(i).Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    Next i
    
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

    Set pfrmUscita = Nothing
End Sub


Private Sub lblCambiaPesa_Click()
    'se sono in lettura continua annullo
    If cmd(CMD_LETTURAPESO).Caption = "Annulla Lettura" Then cmd_Click (CMD_LETTURAPESO)
    mRichiestaPesoInCorso = False
    
    frmCambiaPesa.Show vbModal
    
End Sub

Private Sub lvwLotti_DblClick()
    Dim listItm As ListItem
    Dim elencoLotti As String
    Dim qty As Double
    
    
    
    qty = 0
    If lvwLotti.SelectedItem Is Nothing Then Exit Sub
    
    elencoLotti = lvwLotti.SelectedItem.ListSubItems("qtyScaricata").Tag
    
    frmDettaglioMovimenti.ShowME lvwLotti.SelectedItem.Text, lvwLotti.SelectedItem.ListSubItems("Descrizione"), qty, elencoLotti
    
    
    If elencoLotti <> "" Then
        If elencoLotti = "-1" Then Exit Sub
        
        If IsAllMovimentsRegistered(lvwLotti.SelectedItem.Text) <> vbIgnore Then Exit Sub
    End If
    
    lvwLotti.SelectedItem.ListSubItems("qtyScaricata").Text = qty
    lvwLotti.SelectedItem.ListSubItems("qtyScaricata").Tag = elencoLotti
    lvwLotti.SelectedItem.Checked = elencoLotti <> ""
    
    If elencoLotti <> "" And lvwLotti.Tag <> lvwLotti.SelectedItem.key And lvwLotti.Tag <> "" Then
        lvwLotti_ItemCheck lvwLotti.SelectedItem
    End If
    


    qty = 0
    For Each listItm In lvwLotti.ListItems
        If (listItm.Checked = True) Then
            qty = qty + getNumber(listItm.ListSubItems("qtyScaricata").Text)
        End If
    Next
    txt(TXT_QTY_GIACENZA).Text = FormatNumber(qty, 2, , , vbFalse)
End Sub

Private Function IsAllMovimentsRegistered(scer As String) As VbMsgBoxResult
    Dim rs As Recordset
    Dim codCer As String
    Dim codCategoria As String
    Dim codSottoCategoria As String
    
    On Error GoTo IsAllMovimentsRegisteredERR
    IsAllMovimentsRegistered = vbIgnore
    'testo se esistono movimenti non resitrati che questo codice CER
    codCategoria = left$(scer, 2)
    codSottoCategoria = Mid$(scer, 3, 2)
    codCer = Right$(scer, 2)
    
    Set rs = gDBC.Execute("select 1 from  movimenti where stato<>'2'  and  flgRegistrato=false and codCategoria='" & codCategoria _
        & "' and codSottoCategoria='" & codSottoCategoria & "' and codCer='" & codCer & "'")
    If Not rs.EOF Then
        If MsgBox("Attenzione esistono dei movimenti non registrati per il codice CER:" & scer & _
            vbCrLf & "Si vuole lanciare la registrazione movimenti ora, per poterlo selezionare?", vbQuestion + vbYesNo, App.Title) = VbMsgBoxResult.vbNo Then
                    
            IsAllMovimentsRegistered = vbNo
            Exit Function
        Else
        
            generazioneRegistrazioneDiCarico True
            
            IsAllMovimentsRegistered = vbYes
            
            DoEvents
            
            RefreshME

            lvwLotti.ListItems("|" & scer).Checked = True
            lvwLotti_ItemCheck lvwLotti.ListItems("|" & scer)

        End If
    End If
    rs.Close
    Set rs = Nothing
    
    Exit Function
IsAllMovimentsRegisteredERR:
    IsAllMovimentsRegistered = vbNo
    MsgBox "IsAllMovimentsRegistered:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Sub lvwLotti_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Dim r As VbMsgBoxResult
    Dim qty As Double
    Dim qty1 As Double
    
    If item.Checked Then
        r = IsAllMovimentsRegistered(item.Text)
        If r = VbMsgBoxResult.vbNo Then
            item.Checked = False
            Exit Sub
        ElseIf r = VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
        item.Selected = True
    End If
    
    'gestione singolo check
    If lvwLotti.Tag <> item.key And lvwLotti.Tag <> "" Then
        lvwLotti.ListItems(lvwLotti.Tag).Checked = False
        lvwLotti_ItemCheck lvwLotti.ListItems(lvwLotti.Tag)
    End If

    qty = 0
    If txt(TXT_QTY_GIACENZA).Text <> "" Then qty = getNumber(txt(TXT_QTY_GIACENZA).Text)
    If item.ListSubItems("qtyScaricata").Text <> "" Then qty1 = getNumber(item.ListSubItems("qtyScaricata").Text)
    
    txt(TXT_QTY_GIACENZA).Text = FormatNumber(qty + (qty1 * IIf(item.Checked, 1, -1)), 2, , , vbFalse)
    DoEvents
    
    If (item.Checked) Then
        If item.ListSubItems("qtyScaricata").Tag = "" Then
            item.ListSubItems("qtyScaricata").Tag = selezionaTuttiILotti(item.Text)
        Else
'            Item.ListSubItems("qtyScaricata").Tag = ""
        End If
        lvwLotti.Tag = item.key
        caricaClassiDiPericolosita item.Text
        
        optNormativa(0).value = item.ListSubItems("RifiutoPericoloso") = "Si"
        optNormativa(1).value = Not optNormativa(0).value
        
        
    Else
        'Item.ListSubItems("qtyScaricata").Tag = ""
        lvwLotti.Tag = ""
        
    End If
     
End Sub

Private Function selezionaTuttiILotti(ByVal codCer As String) As String
    Dim rs As Recordset
    Dim codCategoria As String
    Dim codSottoCategoria As String
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    
    codCategoria = left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codCer = Right$(codCer, 2)

    
    rs.Source = "SELECT Lotti.ID FROM Lotti WHERE QtaIngresso>0 and codCategoria ='" & codCategoria & "' and codSottoCategoria ='" & codSottoCategoria & "' and codCer ='" & codCer & "'"
    rs.Open
    Set rs.ActiveConnection = Nothing
    Do While Not rs.EOF
        selezionaTuttiILotti = selezionaTuttiILotti & rs(0) & ","
        rs.MoveNext
    Loop
    selezionaTuttiILotti = left$(selezionaTuttiILotti, Len(selezionaTuttiILotti) - 1)
    
End Function

Private Sub picContainer_Resize(Index As Integer)
    Dim posLeft As Long
    Dim posTop As Long
    
    On Error Resume Next
    Frame3.Move offsetX, offsetY, picContainer(0).ScaleWidth - offsetX * 2
    
    posTop = Frame3.Height + Frame3.top
    Frame2.Move offsetX, posTop, Frame3.width, picContainer(0).ScaleHeight - posTop
    posTop = txt(TXT_NR_FORMULARIO).top + txt(TXT_NR_FORMULARIO).Height + offsetY
    lvwLotti.Move offsetX, posTop, Frame2.width - offsetX * 2, Frame2.Height - posTop - offsetY
    
    'Trasportatore / Destinatario
    posLeft = Frame1.width / 2
    txt(TXT_COD_TRASPORTATORE).Move posLeft
    Label1(26).Move posLeft
    cmbRicerca(0).Move posLeft + Label1(26).width + offsetX
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Move cmbRicerca(0).left + cmbRicerca(0).width + offsetX
    Label1(LBL_TRASPOR_RAGIONESOCIALE - 1).Move Label1(LBL_TRASPOR_RAGIONESOCIALE).left
    Label1(LBL_TRASPOR_CODICEFISCALE).Move Label1(LBL_TRASPOR_RAGIONESOCIALE).left
    Label1(LBL_TRASPOR_CODICEFISCALE - 1).Move Label1(LBL_TRASPOR_RAGIONESOCIALE).left

    If (Label1(LBL_TRASPOR_CODICEFISCALE - 1).width > Label1(LBL_TRASPOR_CODICEFISCALE).width) Then
        Label1(LBL_TRASPOR_PIVA).left = Label1(LBL_TRASPOR_CODICEFISCALE - 1).width + Label1(LBL_TRASPOR_CODICEFISCALE - 1).left + 2 * offsetX
        Label1(LBL_TRASPOR_PIVA - 1).left = Label1(LBL_TRASPOR_CODICEFISCALE - 1).width + Label1(LBL_TRASPOR_CODICEFISCALE - 1).left + 2 * offsetX
    Else
        Label1(LBL_TRASPOR_PIVA).left = Label1(LBL_TRASPOR_CODICEFISCALE).width + Label1(LBL_TRASPOR_CODICEFISCALE).left + 2 * offsetX
        Label1(LBL_TRASPOR_PIVA - 1).left = Label1(LBL_TRASPOR_CODICEFISCALE).width + Label1(LBL_TRASPOR_CODICEFISCALE).left + 2 * offsetX
    End If
    
    
    cmbConducente.Move posLeft
    Label1(22).Move posLeft
    
    cmbTarga.Move cmbConducente.left + cmbConducente.width + offsetX
    Label1(33).Move cmbTarga.left
    
    posLeft = txt(TXT_QTY_GIACENZA).left + txt(TXT_QTY_GIACENZA).width + offsetX
    Label1(24).left = posLeft
    txt(TXT_NR_FORMULARIO).left = posLeft
    Label1(17).left = txt(TXT_NR_FORMULARIO).left + txt(TXT_NR_FORMULARIO).width + offsetX
    dtpickFormulario.left = Label1(17).left
    
    
    Frame1.Move offsetX, offsetY, picContainer(0).ScaleWidth - offsetX * 2
    Frame5.Move offsetX, Frame1.top + Frame1.Height + offsetY * 3, Frame1.width
    
    'Formulario
    txt(TXT_PERCORSO).width = Frame1.width - txt(TXT_PERCORSO).left - offsetX
    cmbDestRifiuto.width = Frame1.width - offsetX * 2
    Label1(32).Move Frame4.left + Frame4.width + offsetX, Frame4.top
    txt(TXT_NUMCOLLI).Move Label1(32).left, Label1(32).top + Label1(32).Height
    
    
    'Registro
    txt(TXT_CARATTER_CHIMICHE_FISICHE).width = Frame1.width / 2 - offsetX * 2
    posLeft = txt(TXT_CARATTER_CHIMICHE_FISICHE).width + txt(TXT_CARATTER_CHIMICHE_FISICHE).left + offsetX
    txt(TXT_LUOGO_PROVENIENZA).Move posLeft, txt(TXT_LUOGO_PROVENIENZA).top, txt(TXT_CARATTER_CHIMICHE_FISICHE).width
    lvwClassiPericolosita.width = txt(TXT_CARATTER_CHIMICHE_FISICHE).width
    txt(TXT_ANNOTAZIONI_REGISTRO).Move posLeft, txt(TXT_ANNOTAZIONI_REGISTRO).top, txt(TXT_CARATTER_CHIMICHE_FISICHE).width
    
    
    DoEvents
    lvwClassiPericolosita.View = lvwList
    lvwClassiPericolosita.View = lvwSmallIcon
    
    Label1(31).left = posLeft
    Label1(28).left = Label1(31).left
End Sub


Private Sub caricaDatiTrasportatore(CodTrasportatore As String)
    Dim rs As Recordset
    
    

    Label1(LBL_TRASPOR_CODICEFISCALE).Caption = ""
    Label1(LBL_TRASPOR_PIVA).Caption = ""
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
     cmbTarga.Clear
    cmbConducente.Clear
    
    If CodTrasportatore = "" Then Exit Sub
    If CodTrasportatore <> vbNullString Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.LockType = adLockBatchOptimistic
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Source = "select *,comuni.comune from trasportatori LEFT JOIN Comuni ON trasportatori.CodComuneSedeLegale = Comuni.CODISTAT where idTrasportatore=" & CodTrasportatore
        
        Set rs.ActiveConnection = gDBC
        rs.Open
        Set rs.ActiveConnection = Nothing
                
        If rs.EOF Then
            MsgBox "Attenzione Trasportatore:" & CodTrasportatore & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
        Else
            
            Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = rs!ragioneSociale
            Label1(LBL_TRASPOR_CODICEFISCALE).Caption = nullToBlank(rs!CodFiscale)
            
            DoEvents
            
            If (Label1(LBL_TRASPOR_CODICEFISCALE - 1).width > Label1(LBL_TRASPOR_CODICEFISCALE).width) Then
                Label1(LBL_TRASPOR_PIVA).left = Label1(LBL_TRASPOR_CODICEFISCALE - 1).width + Label1(LBL_TRASPOR_CODICEFISCALE - 1).left + 2 * offsetX
                Label1(LBL_TRASPOR_PIVA - 1).left = Label1(LBL_TRASPOR_CODICEFISCALE - 1).width + Label1(LBL_TRASPOR_CODICEFISCALE - 1).left + 2 * offsetX
            Else
                Label1(LBL_TRASPOR_PIVA).left = Label1(LBL_TRASPOR_CODICEFISCALE).width + Label1(LBL_TRASPOR_CODICEFISCALE).left + 2 * offsetX
                Label1(LBL_TRASPOR_PIVA - 1).left = Label1(LBL_TRASPOR_CODICEFISCALE).width + Label1(LBL_TRASPOR_CODICEFISCALE).left + 2 * offsetX
            End If
            Label1(LBL_TRASPOR_PIVA).Caption = nullToBlank(rs!PartitaIva)
            
            DoEvents
            
        End If
        rs.Close
    End If

    'carica combo targhe

    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    DevEnvMDB.cmdAutomezzi CodTrasportatore, "T"
    Do While Not DevEnvMDB.rscmdAutomezzi.EOF
        cmbTarga.AddItem DevEnvMDB.rscmdAutomezzi!Targa
        DevEnvMDB.rscmdAutomezzi.MoveNext
    Loop
    If cmbTarga.ListCount > 0 Then cmbTarga.ListIndex = 0
    
    
    'carica combo conducenti

    If DevEnvMDB.rscmdConducenti.state = adStateOpen Then
        DevEnvMDB.rscmdConducenti.Close
    End If
    DevEnvMDB.cmdConducenti CodTrasportatore
    Do While Not DevEnvMDB.rscmdConducenti.EOF
        cmbConducente.AddItem DevEnvMDB.rscmdConducenti!Cognome & " " & DevEnvMDB.rscmdConducenti!Nome
        cmbConducente.ItemData(cmbConducente.ListCount - 1) = DevEnvMDB.rscmdConducenti!idConducente
        DevEnvMDB.rscmdConducenti.MoveNext
    Loop
    If cmbConducente.ListCount > 0 Then cmbConducente.ListIndex = 0
    
End Sub


Private Sub caricaIntermediario(cod As String)
    Dim rs As Recordset
    

    Label1(LBL_INTERMEDIARIO_RAGIONESOCIALE).Caption = ""
    
    If cod = "" Then Exit Sub
    
    If cod <> vbNullString Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.LockType = adLockBatchOptimistic
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Source = "select *,comuni.comune from Intermediari LEFT JOIN Comuni ON Intermediari.CodComune = Comuni.CODISTAT where idIntermediario=" & cod
        
        Set rs.ActiveConnection = gDBC
        rs.Open
        Set rs.ActiveConnection = Nothing
                
        If rs.EOF Then
            MsgBox "Attenzione Intermediario: " & cod & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
        Else
            
            Label1(LBL_DESTINAT_RAGIONESOCIALE).Caption = rs!ragioneSociale
            DoEvents
            
            
            'Label1(LBL_DESTINAT_CODCEFISCALE).Caption = nullToBlank(rs!CodFiscale)
            'DoEvents
            
            'If (Label1(LBL_DESTINAT_CODCEFISCALE + 1).Width > Label1(LBL_DESTINAT_CODCEFISCALE).Width) Then
            '    Label1(LBL_DESTINAT_PIVA).Left = Label1(LBL_DESTINAT_CODCEFISCALE + 1).Width + Label1(LBL_DESTINAT_CODCEFISCALE + 1).Left + 2 * offsetX
            '    Label1(LBL_DESTINAT_PIVA + 1).Left = Label1(LBL_DESTINAT_CODCEFISCALE + 1).Width + Label1(LBL_DESTINAT_CODCEFISCALE + 1).Left + 2 * offsetX
            'Else
            '    Label1(LBL_DESTINAT_PIVA).Left = Label1(LBL_DESTINAT_CODCEFISCALE).Width + Label1(LBL_DESTINAT_CODCEFISCALE).Left + 2 * offsetX
            '    Label1(LBL_DESTINAT_PIVA + 1).Left = Label1(LBL_DESTINAT_CODCEFISCALE).Width + Label1(LBL_DESTINAT_CODCEFISCALE).Left + 2 * offsetX
            'End If
            'Label1(LBL_DESTINAT_PIVA).Caption = nullToBlank(rs!PartitaIva)
            
            DoEvents
            
        End If
        rs.Close
    End If
    
End Sub


Private Sub caricaDatiDestinatario(cod As String)
    Dim rs As Recordset
    

    Label1(LBL_DESTINAT_CODCEFISCALE).Caption = ""
    Label1(LBL_DESTINAT_PIVA).Caption = ""
    Label1(LBL_DESTINAT_RAGIONESOCIALE).Caption = ""
    cmbImpianto.Clear
    
    If cod = "" Then Exit Sub
    If cod <> vbNullString Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.LockType = adLockBatchOptimistic
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Source = "select *,comuni.comune from destinatari LEFT JOIN Comuni ON destinatari.CodComune = Comuni.CODISTAT where idDEstinatario=" & cod
        
        Set rs.ActiveConnection = gDBC
        rs.Open
        Set rs.ActiveConnection = Nothing
                
        If rs.EOF Then
            MsgBox "Attenzione Destinatario: " & cod & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
        Else
            
            Label1(LBL_DESTINAT_RAGIONESOCIALE).Caption = rs!ragioneSociale
            DoEvents
            
            
            Label1(LBL_DESTINAT_CODCEFISCALE).Caption = nullToBlank(rs!CodFiscale)
            DoEvents
            
            If (Label1(LBL_DESTINAT_CODCEFISCALE + 1).width > Label1(LBL_DESTINAT_CODCEFISCALE).width) Then
                Label1(LBL_DESTINAT_PIVA).left = Label1(LBL_DESTINAT_CODCEFISCALE + 1).width + Label1(LBL_DESTINAT_CODCEFISCALE + 1).left + 2 * offsetX
                Label1(LBL_DESTINAT_PIVA + 1).left = Label1(LBL_DESTINAT_CODCEFISCALE + 1).width + Label1(LBL_DESTINAT_CODCEFISCALE + 1).left + 2 * offsetX
            Else
                Label1(LBL_DESTINAT_PIVA).left = Label1(LBL_DESTINAT_CODCEFISCALE).width + Label1(LBL_DESTINAT_CODCEFISCALE).left + 2 * offsetX
                Label1(LBL_DESTINAT_PIVA + 1).left = Label1(LBL_DESTINAT_CODCEFISCALE).width + Label1(LBL_DESTINAT_CODCEFISCALE).left + 2 * offsetX
            End If
            Label1(LBL_DESTINAT_PIVA).Caption = nullToBlank(rs!PartitaIva)
            
            DoEvents
            
        End If
        rs.Close
    End If
    
    'carica combo impianti
    If DevEnvMDB.rscmdImpianti.state = adStateOpen Then
        DevEnvMDB.rscmdImpianti.Close
    End If
    DevEnvMDB.cmdImpianti
    
    DevEnvMDB.rscmdImpianti.Filter = " idSoggetto=" & cod & " and codSoggetto='D'"
    
    Do While Not DevEnvMDB.rscmdImpianti.EOF
        cmbImpianto.AddItem DevEnvMDB.rscmdImpianti!ragioneSociale & " - " & DevEnvMDB.rscmdImpianti!Indirizzo
        cmbImpianto.ItemData(cmbImpianto.ListCount - 1) = DevEnvMDB.rscmdImpianti!idImpianto
        DevEnvMDB.rscmdImpianti.MoveNext
    Loop
    If cmbImpianto.ListCount > 0 Then cmbImpianto.ListIndex = 0
End Sub


Private Sub caricaTotali()
    Dim itm As ListItem
    Dim cer As String
    Dim qry As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    
    On Error GoTo caricaTotaliERR
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = DevEnvMDB.rscmdGiacenzeTotali.Source
    Set qry.ActiveConnection = gDBC
    qry.Parameters.Refresh
    qry.Parameters(0).value = gGGAllarmeGiacGestionale
    qry.Parameters(1).value = gGGAllarmeGiacNormativa

    Set rs = qry.Execute
    
    lvwLotti.ListItems.Clear
    Do While Not rs.EOF
        
        cer = rs!codCategoria & _
              rs!codSottoCategoria & _
              rs!codCer
        With lvwLotti
            Set itm = .ListItems.Add(, "|" & cer, cer)
            itm.ListSubItems.Add , rs!Descrizione.Name, rs!Descrizione
            
            itm.ListSubItems.Add , rs!qtyResidua.Name, FormatNumber(rs!qtyResidua, 2, , , vbFalse)
            itm.ListSubItems.Add , rs!QtyDaRegistrare.Name, FormatNumber(rs!QtyDaRegistrare, 2, , , vbFalse)
            itm.ListSubItems.Add , "qtyScaricata", FormatNumber(rs!qtyResidua, 2, , , vbFalse)
            itm.ListSubItems.Add , rs!RifiutoPericoloso.Name, IIf(rs!RifiutoPericoloso, "Si", "No")
            
            itm.ListSubItems.Add , rs!qtyMassimaGiacenza.Name, toZero(rs!qtyMassimaGiacenza)
            itm.ListSubItems.Add , rs!qtyMassimaGiacenzaNormativa.Name, toZero(rs!qtyMassimaGiacenzaNormativa)
            itm.ListSubItems.Add , rs!DataPrevistaRitiro.Name, Format$(rs!DataPrevistaRitiro, "dd/mm/yyyy")
            itm.ListSubItems.Add , rs!DataPrevistaRitiroNormativa.Name, Format$(rs!DataPrevistaRitiroNormativa, "dd/mm/yyyy")
            itm.ListSubItems.Add , rs!MinDiDataIns.Name, Format$(rs!MinDiDataIns, "dd/mm/yyyy")
        End With
        
        rs.MoveNext
    Loop
    
    Set rs1 = CreateObject("ADODB.Recordset")
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.LockType = adLockBatchOptimistic
    
    
    If DevEnvMDB.rscmdCerNonCaricati.state = adStateOpen Then
        DevEnvMDB.rscmdCerNonCaricati.Close
    End If
    DevEnvMDB.cmdCerNonCaricati
    
    
    Set rs1 = CreateObject("ADODB.Recordset")
    rs1.LockType = adLockBatchOptimistic
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    
    rs1.Source = DevEnvMDB.rscmdCerNonCaricati.Source
    Set rs1.ActiveConnection = gDBC
    rs1.Open
    Set rs1.ActiveConnection = Nothing
     
    Do While Not rs1.EOF
        
        cer = rs1!codCategoria & _
              rs1!codSottoCategoria & _
              rs1!codCer
        With lvwLotti
            
            Set itm = .ListItems.Add(, "|" & cer, cer)
            itm.ListSubItems.Add , rs1!Descrizione.Name, rs1!Descrizione
            
            itm.ListSubItems.Add , , 0
            itm.ListSubItems.Add , , FormatNumber(rs1!quantita, 2, , , vbFalse)
            itm.ListSubItems.Add , , FormatNumber(rs1!quantita, 2, , , vbFalse)
            itm.ListSubItems.Add , , "No"
            
            itm.ListSubItems.Add , , 0
            itm.ListSubItems.Add , , 0
            itm.ListSubItems.Add , , Format$(Now, "dd/mm/yyyy")
            itm.ListSubItems.Add , , Format$(Now, "dd/mm/yyyy")
            itm.ListSubItems.Add , , Format$(Now, "dd/mm/yyyy")
            
            coloraRiga itm, vbBlue
        End With
        
        rs1.MoveNext
    Loop
     
    AllargaLeColonneLVW lvwLotti, False
    Exit Sub
caricaTotaliERR:
    MsgBox "CaricaTotali:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub letturaPeso(codPesa As String)
    Dim peso As Double
    Dim messaggio As String
    
    On Error GoTo letturaPesoERR
    peso = 0
    Do While peso = 0 And Not Esci
        
        'richiesta peso alla bilancia
        If Not mRichiestaPesoInCorso Then
            If InvioComandoAllaBilancia(cmd_RichiestaPesa, True) Then mRichiestaPesoInCorso = True
        End If
         
        Select Case pesa(codPesa).ReadMessage(messaggio)
            Case ErroreRicevuto
            Case InLettura
            Case LetturaTerminata

            
                Select Case getCurrentPesa.ModelloPesa
                    Case "LAUMAS400"
                        msgPesa = getMessaggioPesataLAumas400(messaggio, vbNullString)
                        If pesa(codPesa).COMControl.PortOpen Then
                             pesa(codPesa).COMControl.PortOpen = False
                             pesa(codPesa).COMControl.PortOpen = True
                        End If
                    Case "RETEL"
                        msgPesa = getMessaggioPesataRETEL(messaggio, vbNullString)
                        If pesa(codPesa).COMControl.PortOpen Then
                             pesa(codPesa).COMControl.PortOpen = False
                             pesa(codPesa).COMControl.PortOpen = True
                        End If
                    Case Else
                        msgPesa = getMessaggioPesata(messaggio)
                End Select
                    
                mRichiestaPesoInCorso = False
                mPesoLettoDaBilancia = True
                'annulla pesata
                If Not getCurrentPesa.ModalitaLetturaPesoAutomatica Then
                    cmd_Click CMD_LETTURAPESO
                Else
                    peso = 0
                End If
                
                txt(TXT_QTY_PESA) = Format$(msgPesa.PesoNetto, "#.00")
         End Select
         DoEvents
    Loop
    
    If Esci Then peso = 0
    
    Exit Sub
letturaPesoERR:
    cmd_Click CMD_LETTURAPESO
End Sub


Private Function salva(forzatura As Boolean, ByRef idMovimento As Long, idAlibiMemory As String) As Boolean
    Dim nrProdottoScaricati As Integer
    Dim listItm As ListItem
    Dim DataMovimento As String
    Dim cmdq As Command
    Dim P As Parameter
    Dim codCer As String
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codSubCer As String
    Dim quantita As Double
    Dim UnitaDiMisura As String
    Dim classiPericolosita As String
    Dim nrMovimento As Long
    
    UnitaDiMisura = "KG"
    
    On Error GoTo salvaErr
    gDBC.BeginTrans
    
    nrMovimento = getNumeratore(MOVIMENTI, False)
    salva = False
    
    nrProdottoScaricati = 0
    
    For Each listItm In lvwLotti.ListItems
        If listItm.Checked Then
            codCer = Mid$(listItm.key, 2)
            codCategoria = left$(codCer, 2)
            codSottoCategoria = Mid$(codCer, 3, 2)
            codCer = Mid$(codCer, 5, 2)
            codSubCer = Right$(codCer, 2)
            
            nrProdottoScaricati = nrProdottoScaricati + 1
            
            If nrProdottoScaricati > 1 Then
                codCer = "US"
                codCategoria = "US"
                codSottoCategoria = "US"
                Exit For
            End If
        End If
    Next
    
    
    DataMovimento = IIf(CLng(msgPesa.Data) <= 0, Now, msgPesa.Data)
    quantita = Round(getNumber(txt(TXT_QTY_PESA).Text), 2)
    
    'salva documento e riferimento al movimento
    If mnrDocumentoPesata = 0 Then
        mnrDocumentoPesata = getNumeratore(enumEntita.DOCUMENTO_PESATA_OUT, False)
        'aggiorna riferimento sulla tabella documenti
        AggiornaRiferimentoInDocumento midDocumentoPesata, mnrDocumentoPesata, midMovimentiDaStampare, "O"
    End If
    
    
    midMovimentoUscita = creaMovimento(nrMovimento, -1, -1, mCodComuneCorrente, codCategoria, codSottoCategoria, codCer, codSubCer, _
                DataMovimento, quantita, "0", IIf(CLng(msgPesa.Data) <= 0, True, False), _
                frmLogin.codUser, 0, 0, UnitaDiMisura, "", txt(TXT_NR_FORMULARIO).Text, _
                Format$(dtpickFormulario.value, "YYYY-MM-DD"), forzatura, IIf(cmbTarga.Text = "", Null, cmbTarga.Text), _
                txt(TXT_COD_TRASPORTATORE).Text, "01", True, gPesaIsolaCod, midDocumentoPesata, idAlibiMemory)

    
    'salvo la targa trasportatore se non presente in anagrafica
    'salvaTarga txt(TXT_COD_TRASPORTATORE), cmbTarga, "T"
    classiPericolosita = ""
    For Each listItm In lvwClassiPericolosita.ListItems
        If listItm.Checked = True Then
            classiPericolosita = classiPericolosita & listItm.Text & ", "
        End If
    Next
    
    For Each listItm In lvwLotti.ListItems
        If listItm.Checked = True Then
            operazioniDiCaricoScarico listItm.ListSubItems("qtyScaricata").Tag, codCategoria, _
                codSottoCategoria, codCer, UnitaDiMisura, left$(cmbDestRifiuto.Text, 3), quantita, classiPericolosita
            Exit For
        End If
    Next
    

    gDBC.CommitTrans
    
    idMovimento = midMovimentoUscita
    
    salva = True
    Exit Function
salvaErr:
    gDBC.RollbackTrans
    
    MsgBox "FrmUscita Salva:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Sub salvaLottoUscitaLottoIngressi(idRegistro As Long, codCategoria As String, codSottoCategoria As String, pcodCer As String, quantita As Double)
    Dim listItm As ListItem
    Dim qryUpdate As Command
    Dim qryAssociazione As Command
    Dim codLotti() As String
    Dim i As Long
    Dim codLotto  As String
    
    On Error GoTo salvaLottoUscitaLottoIngressiERR

    
    codLotto = creaLotto(idRegistro, codCategoria, codSottoCategoria, pcodCer, _
        quantita, 0, "U", 0, Now, frmLogin.codUser, Now, frmLogin.codUser)
    
    Set qryUpdate = CreateObject("ADODB.Command")
    qryUpdate.CommandType = adCmdText
    Set qryUpdate.ActiveConnection = gDBC
     
    Set qryAssociazione = CreateObject("ADODB.Command")
    qryAssociazione.CommandType = adCmdText
    qryAssociazione.CommandText = "insert into LegameLottoIngressoLottoUscita (LottoUscita,idLottoEntrata,DataIns,CodUser) values (?,?,?,?)"
    
    qryAssociazione.Parameters.Append qryAssociazione.CreateParameter(, adVarChar, adParamInput, 50)
    qryAssociazione.Parameters.Append qryAssociazione.CreateParameter(, adInteger, adParamInput)
    qryAssociazione.Parameters.Append qryAssociazione.CreateParameter(, adDate, adParamInput)
    qryAssociazione.Parameters.Append qryAssociazione.CreateParameter(, adVarChar, adParamInput, 10)
    
    Set qryAssociazione.ActiveConnection = gDBC
    
    For Each listItm In lvwLotti.ListItems
        If listItm.Checked = True Then
            codLotti = Split(listItm.ListSubItems("qtyScaricata").Tag, ",")
            
            'azzero le giacenze per lotto
            qryUpdate.CommandText = "Update lotti set qtaUscita=QtaIngresso,QtaIngresso=0   where id in (" & listItm.ListSubItems("qtyScaricata").Tag & ")"
            qryUpdate.Execute
            
            'creo associazione
            For i = LBound(codLotti) To UBound(codLotti)
                qryAssociazione.Parameters(0) = codLotto
                qryAssociazione.Parameters(1) = codLotti(i)
                qryAssociazione.Parameters(2) = Date
                qryAssociazione.Parameters(3) = frmLogin.codUser
                qryAssociazione.Execute
            Next
        End If
    Next
    
    Set qryUpdate.ActiveConnection = Nothing
    Set qryAssociazione.ActiveConnection = Nothing
    Exit Sub
salvaLottoUscitaLottoIngressiERR:
    ScriviLog TPL_ERRORE, App.Title, "salvaLottoUscitaLottoIngressi", Err.Description

End Sub

Private Sub TabStrip_Click()
    Dim i As Integer
    
    For i = picContainer.LBound To picContainer.ubound
        picContainer(i).Visible = False
    Next
    
    picContainer(TabStrip.SelectedItem.Index - 1).Visible = True
    
    
    cmd(CMD_LETTURAPESO).enabled = TabStrip.SelectedItem.key = "PESATA"
    cmd(CMD_DETTAGLIOMOV).enabled = TabStrip.SelectedItem.key = "PESATA"
    
    Select Case TabStrip.SelectedItem.key
        Case "FORMULARIOREGISTRO"
            cmbDestRifiuto.SetFocus
            DTPickDataOraTrasporto.SetFocus
    End Select
    
End Sub

Private Sub txt_Change(Index As Integer)
    Dim qty As Double
    Dim qty1 As Double
On Error GoTo txt_ChangeERR

    If (Index = TXT_QTY_PESA) Or _
     (Index = TXT_QTY_GIACENZA) Then
     
        If txt(TXT_QTY_PESA) <> "" Then qty = getNumber(txt(TXT_QTY_PESA).Text)
        If txt(TXT_QTY_GIACENZA) <> "" Then qty1 = getNumber(txt(TXT_QTY_GIACENZA).Text)
        
        txt(TXT_DIFF_PESO).Text = Format$(qty - qty1, "00.00")
    End If
    Exit Sub
txt_ChangeERR:
    ScriviLog TPL_ERRORE, App.Title, "txt_Change", Err.Description

End Sub


Private Function validaDati(ByRef idAlibiMemory As String) As Boolean
    Dim campoInErrore1 As Boolean
    Dim campoInErrore2 As Boolean
    Dim rs As ADODB.Recordset
    Dim qty As Double
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codCer As String
    Dim codSubCer As String
    Dim listItm As ListItem
    Dim bTrans As Boolean
    Dim ColorError As Long
    On Error GoTo validaDatiERR
    
    bTrans = False
    validaDati = False
    campoInErrore1 = False
    campoInErrore2 = False
    
    txt(TXT_COD_TRASPORTATORE).ToolTipText = ""
    txt(TXT_COD_DESTINATARIO).ToolTipText = ""
    cmbConducente.ToolTipText = ""
    txt(TXT_NR_FORMULARIO).ToolTipText = ""
    cmbImpianto.ToolTipText = ""
    txt(TXT_QTY_PESA).ToolTipText = ""
    lvwLotti.ToolTipText = ""
    dtpickFormulario.ToolTipText = ""
    cmbTarga.ToolTipText = ""
    
    txt(TXT_COD_TRASPORTATORE).BackColor = -2147483643
    txt(TXT_COD_DESTINATARIO).BackColor = -2147483643
    txt(TXT_QTY_PESA).BackColor = vbBlack
    txt(TXT_NR_FORMULARIO).BackColor = -2147483643
    lvwLotti.BackColor = -2147483643
    dtpickFormulario.CalendarBackColor = -2147483643
    cmbTarga.BackColor = -2147483643
    cmbImpianto.BackColor = -2147483643
    cmbConducente.BackColor = -2147483643
    
    ColorError = &H80C0FF
    
    If (cmbTarga.Text = "") Then
        cmbTarga.ToolTipText = "Attenzione la targa è obbligatoria"
        cmbTarga.BackColor = ColorError
        If (TabStrip.SelectedItem.key = "PESATA") Then cmbTarga.SetFocus
        
        campoInErrore1 = True
    Else
        Set rs = gDBC.Execute("select codutente,codSoggetto from automezzi where codSoggetto='T' and  codutente<>" & txt(TXT_COD_TRASPORTATORE) & " and targa='" & cmbTarga.Text & "'")
        If Not rs.EOF Then
            cmbTarga.ToolTipText = "Attenzione la targa è già presente ed assegnata al soggetto:" & rs(0) & " tipo:" & rs(1)
            cmbTarga.BackColor = ColorError
            If (TabStrip.SelectedItem.key = "PESATA") Then cmbTarga.SetFocus
            
            campoInErrore1 = True
        End If
        rs.Close
    End If

    
    If txt(TXT_COD_TRASPORTATORE) = "" Then
        txt(TXT_COD_TRASPORTATORE).ToolTipText = "Attenzione codice trasportatore obbligatorio"
        txt(TXT_COD_TRASPORTATORE).BackColor = ColorError
        'txt(TXT_COD_TRASPORTATORE).SetFocus

        campoInErrore1 = True
    End If
    
    If txt(TXT_COD_DESTINATARIO) = "" Then
        txt(TXT_COD_DESTINATARIO).ToolTipText = "Attenzione codice destinatario obbligatorio"
        txt(TXT_COD_DESTINATARIO).BackColor = ColorError
        'txt(TXT_COD_DESTINATARIO).SetFocus
            
        campoInErrore1 = True
    ElseIf cmbImpianto.ListIndex < 0 Then
        cmbImpianto.ToolTipText = "Attenzione impianto obbligatorio"
        cmbImpianto.BackColor = ColorError
        If (TabStrip.SelectedItem.key = "PESA") Then cmbImpianto.SetFocus
            
        campoInErrore1 = True
    
    End If
    
    If cmbConducente.ListIndex < 0 Then
        cmbConducente.ToolTipText = "Attenzione conducente obbligatorio"
        cmbConducente.BackColor = ColorError
        If (TabStrip.SelectedItem.key = "PESA") Then cmbConducente.SetFocus
            
        campoInErrore1 = True
    
    End If
    
    qty = getNumber(txt(TXT_QTY_GIACENZA).Text)
    If qty <= 0 Then
        lvwLotti.ToolTipText = "Attenzione il rifiuto da scaricare è obbligatorio, la quantità deve essere > 0"
        lvwLotti.BackColor = ColorError
        If campoInErrore1 = False And TabStrip.SelectedItem.key = "PESATA" Then lvwLotti.SetFocus
        campoInErrore1 = True
    End If
    
    qty = getNumber(txt(TXT_QTY_PESA).Text)
    If qty <= 0 Then
        txt(TXT_QTY_PESA).ToolTipText = "Attenzione la quantità deve essere > 0"
        txt(TXT_QTY_PESA).BackColor = ColorError
        If campoInErrore1 = False And (TabStrip.SelectedItem.key = "PESATA") Then txt(TXT_QTY_PESA).SetFocus
        campoInErrore1 = True
    ElseIf Not msgPesa.PesoStabile Then
         MsgBox "Attenzione il peso non è stabile, attendere che si stabilizzi !!"
        If campoInErrore1 = False Then txt(TXT_QTY_PESA).SetFocus
        campoInErrore1 = True
    End If
    
    If (txt(TXT_NR_FORMULARIO).Text = "") Then
        txt(TXT_NR_FORMULARIO).ToolTipText = "Attenzione numero formulario è obbligatorio"
        txt(TXT_NR_FORMULARIO).BackColor = ColorError
        If (TabStrip.SelectedItem.key = "FORMULARIOREGISTRO") Then txt(TXT_NR_FORMULARIO).SetFocus

        campoInErrore1 = True
    Else 'verifico che il formulario non esista già
        Set rs = gDBC.Execute("select 1 from FormularioRifiuti where id='" & txt(TXT_NR_FORMULARIO).Text & "'")
        If Not rs.EOF Then
            txt(TXT_NR_FORMULARIO).ToolTipText = "Attenzione numero formulario è già presente in archivio"
            txt(TXT_NR_FORMULARIO).BackColor = ColorError
            If (TabStrip.SelectedItem.key = "FORMULARIOREGISTRO") Then txt(TXT_NR_FORMULARIO).SetFocus
        End If
        rs.Close
    End If
        
    If (IsNull(dtpickFormulario.value)) Then
        dtpickFormulario.ToolTipText = "Attenzione la data formulario è obbligatoria"
        dtpickFormulario.CalendarBackColor = ColorError
        If (TabStrip.SelectedItem.key = "FORMULARIOREGISTRO") Then dtpickFormulario.SetFocus
        
        campoInErrore1 = True
    End If
    
    
    qty = getNumber(txt(TXT_DIFF_PESO).Text)
    If campoInErrore1 = False And campoInErrore2 = False And qty <> 0 Then
        If qty > 0 Then
            txt(TXT_QTY_PESA).ToolTipText = "Attenzione la quantità rilevata è superiore a quella di giacenza. " & _
             "Occorre generare un movimento di correzione per la quantità in eccesso." & vbCrLf & _
             "Si vuole generare il MOVIMENTO pari a " & qty & " ?"
             
             If (MsgBox(txt(TXT_QTY_PESA).ToolTipText, vbExclamation + vbYesNo, App.Title) = vbYes) Then
                'se il peso risulata essere maggiore della giacenza creo un movimento di ingresso pari alla
                'differenza e lo registro
                'cerco i codcer selezionato
                For Each listItm In lvwLotti.ListItems
                    If listItm.Checked Then
                        codCer = Mid$(listItm.key, 2)
                        codCategoria = left$(codCer, 2)
                        codSottoCategoria = Mid$(codCer, 3, 2)
                        codCer = Mid$(codSubCer, 5, 2)
                        codSubCer = Right$(codCer, 2)
                    End If
                Next
                
                bTrans = True
                gDBC.BeginTrans
                creaMovimento getNumeratore(MOVIMENTI, False), gCodImmobileMovManuale, gCodUtenteMovManuale, mCodComuneCorrente, codCategoria, codSottoCategoria, codCer, codSubCer, _
                    Format$(Now, "yyyy-MM-dd hh:mm:ss"), qty, "0", True, gCodUserMovManuale, 0, 0, "KG", "", "", _
                     "", False, "", -1, gCodCausaleMovManuale, False, gPesaIsolaCod, 0, vbNullString
                                
                generazioneRegistrazioneDiCarico True
                gDBC.CommitTrans
                 bTrans = False
                
                RefreshME
                    
                DoEvents
                
                TabStrip.Tabs("PESATA").Selected = True
                TabStrip_Click
                
                lvwLotti.ListItems("|" & codCategoria & codSottoCategoria & codCer).Checked = True
                lvwLotti_ItemCheck lvwLotti.ListItems("|" & codCategoria & codSottoCategoria & codCer)
                
                If MsgBox("Ora è possibile ripetere l'operazione di salvataggio." & vbCrLf & _
                    "Lo si vuole fare adesso?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    
                    cmd_Click CMD_SALVA
                End If
                
                 
                Exit Function
             Else
                txt(TXT_QTY_PESA).BackColor = ColorError
                If campoInErrore1 = False And (TabStrip.SelectedItem.key = "PESATA") Then txt(TXT_QTY_PESA).SetFocus
                campoInErrore1 = True
             End If
             
        Else
            txt(TXT_QTY_PESA).ToolTipText = "Attenzione la quantità rilevata è inferiore a quella di giacenza. " & _
            "Provare a deselezionare alcuni di movimenti di carico per raggiungere la quantità rilevata."
            
            txt(TXT_QTY_PESA).BackColor = ColorError
            If campoInErrore1 = False And (TabStrip.SelectedItem.key = "PESATA") Then txt(TXT_QTY_PESA).SetFocus
            campoInErrore1 = True
        End If
        
    End If
    
    validaDati = Not campoInErrore1 And Not campoInErrore2
    
    
    If campoInErrore1 And Not campoInErrore2 Then
        TabStrip.Tabs("PESATA").Selected = True
        TabStrip_Click
    ElseIf campoInErrore2 Then
        TabStrip.Tabs("FORMULARIOREGISTRO").Selected = True
        TabStrip_Click
    End If
    
    'se richiesto ALIBI MEMORY WRITE e sono in peso proveniente da bilancia, mando il messaggio di scrittura
    If validaDati Then
        If getCurrentPesa.CmdScritturaAlibiM <> "" And mPesoLettoDaBilancia Then
            'If Not InvioComandoAllaBilancia(cmd_ScritturaAlibi, True, idAlibiMemory) Then
            '    validaDati = False
            'Else
            '    'resetto la modalità da PESA
            '    mPesoLettoDaBilancia = False
            'End If
            InvioComandoAllaBilancia cmd_ScritturaAlibi, True, idAlibiMemory
            mPesoLettoDaBilancia = False
        End If
    End If
    
    
    Exit Function
validaDatiERR:
    If bTrans Then gDBC.RollbackTrans
    MsgBox "FrmUscita ValidaDati:" & Err.Description, vbExclamation + vbOKOnly, App.Title
    
End Function

Private Sub stampaFormulario(nrRegistro As Long, codCategoria As String, codSottoCategoria As String, codCer As String)
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    
    
    On Error GoTo stampaFormularioERR
    
    
    'formulario
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "INSERT INTO FormularioRifiuti (id,data,nrRegistro,  idMovimento, idImpProduttore, idImpDestinatario, idTrasportatore, idConducente, " & _
                "Annotazioni, codCategoria, codSottoCategoria, codCer, codOpeSmalRecup, carattChimicheFisiche, quantita, percorso, flgNormativaADR_RID, targa, dataInizioTrasporto,numColli, DataIns, CodUser ) " & _
                       "values (?,?,?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "

    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 35)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    Set cmdq.ActiveConnection = gDBC
    
    cmdq.Parameters(0).value = txt(TXT_NR_FORMULARIO).Text
    cmdq.Parameters(1).value = dtpickFormulario.value
    cmdq.Parameters(2).value = nrRegistro
    cmdq.Parameters(3).value = midMovimentoUscita
    cmdq.Parameters(4).value = gIdImpiantoProduttore
    cmdq.Parameters(5).value = cmbImpianto.ItemData(cmbImpianto.ListIndex)
    cmdq.Parameters(6).value = txt(TXT_COD_TRASPORTATORE).Text
    cmdq.Parameters(7).value = cmbConducente.ItemData(cmbConducente.ListIndex)
    cmdq.Parameters(8).value = txt(TXT_ANNOTAZIONI_FORMULARIO).Text
    cmdq.Parameters(9).value = codCategoria
    cmdq.Parameters(10).value = codSottoCategoria
    cmdq.Parameters(11).value = codCer
    cmdq.Parameters(12).value = left$(cmbDestRifiuto.Text, 3)
    cmdq.Parameters(13).value = txt(TXT_CARATTER_CHIMICHE_FISICHE).Text
    cmdq.Parameters(14).value = getNumber(txt(TXT_QTY_PESA).Text)
    cmdq.Parameters(15).value = txt(TXT_PERCORSO).Text
    cmdq.Parameters(16).value = IIf(optNormativa(0).value = -1, True, False)
    cmdq.Parameters(17).value = cmbTarga.Text
    cmdq.Parameters(18).value = DTPickDataOraTrasporto.value
    cmdq.Parameters(19).value = txt(TXT_NUMCOLLI).Text
    
    cmdq.Parameters(20).value = Now
    cmdq.Parameters(21).value = frmLogin.codUser
    
    cmdq.Execute
    
    Exit Sub
stampaFormularioERR:
        
    Err.Raise 9999, "StampaFormulario", Err.Description
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo txt_KeyPressERR

    Select Case Index
    
        Case TXT_COD_DESTINATARIO, TXT_COD_TRASPORTATORE
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
        
        Case TXT_QTY_PESA, TXT_QTY_GIACENZA
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete _
                And KeyAscii <> Asc(",") Then
                KeyAscii = 0
            Else
                mPesoLettoDaBilancia = False
                msgPesa.PesoStabile = True
            End If
    End Select
    Exit Sub
txt_KeyPressERR:
    ScriviLog TPL_ERRORE, App.Title, "txt_KeyPress", Err.Description

End Sub


Private Sub caricaClassiDiPericolosita(codCer As String)
    Dim rs As Recordset
    Dim itm As ListItem
    Dim codCategoria As String
    Dim codSottoCategoria As String
    
On Error GoTo caricaClassiDiPericolositaERR

    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    If codCer <> "" Then
        codCategoria = left$(codCer, 2)
        codSottoCategoria = Mid$(codCer, 3, 2)
        codCer = Right$(codCer, 2)
        
        rs.Source = "select * from AssCodCerClassidiPericolo " & _
            " where (AssCodCerClassidiPericolo.codCategoria='" & codCategoria & "'" & _
            " and AssCodCerClassidiPericolo.codSottoCategoria='" & codSottoCategoria & "' and AssCodCerClassidiPericolo.CodCer='" & codCer & "') "
            
        For Each itm In lvwClassiPericolosita.ListItems
            itm.Checked = False
        Next
    Else
        lvwClassiPericolosita.ListItems.Clear
        rs.Source = "select ClassiPericolo.* from  ClassiPericolo order by id"
    End If

    Set rs.ActiveConnection = gDBC
    rs.Open
        
    lvwClassiPericolosita.View = lvwIcon
    lvwClassiPericolosita.View = lvwList
    
    Do While Not rs.EOF
        If codCer = "" Then
            Set itm = lvwClassiPericolosita.ListItems.Add(, rs!CodClasse.value, rs!CodClasse.value, 0, 0)
        Else

            lvwClassiPericolosita.ListItems(rs!CodClasse.value).Checked = True
        End If
        
        rs.MoveNext
    Loop
    
    
    rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
    Exit Sub
caricaClassiDiPericolositaERR:
    ScriviLog TPL_ERRORE, App.Title, "caricaClassiDiPericolosita", Err.Description

End Sub

Public Sub RefreshME()
    
    lvwLotti.Tag = ""
    txt(TXT_QTY_GIACENZA).Text = 0
    txt(TXT_DIFF_PESO).Text = 0
    
    caricaTotali
    
  
    
End Sub
