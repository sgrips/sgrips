VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCalcoloTariffa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calcolo della Tari"
   ClientHeight    =   9765
   ClientLeft      =   2685
   ClientTop       =   1350
   ClientWidth     =   12690
   Icon            =   "FrmCalcoloTariffa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   2
      Left            =   60
      ScaleHeight     =   8910
      ScaleWidth      =   12600
      TabIndex        =   171
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari Comma 652 2/2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   13
         Left            =   60
         TabIndex        =   172
         Top             =   -45
         Width           =   12540
         Begin MSComctlLib.ListView lvw10 
            Height          =   8220
            Left            =   135
            TabIndex        =   173
            Top             =   540
            Width           =   12150
            _ExtentX        =   21431
            _ExtentY        =   14499
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7710
      Index           =   1
      Left            =   45
      ScaleHeight     =   7710
      ScaleWidth      =   12615
      TabIndex        =   153
      Top             =   45
      Width           =   12615
      Begin VB.Frame Frame1 
         Caption         =   "Tari Comma 652 1/2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7755
         Index           =   12
         Left            =   60
         TabIndex        =   154
         Top             =   -45
         Width           =   12435
         Begin VB.CheckBox chkTarsu507perNComponenti 
            Caption         =   "Tari per numero componenti"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   480
            TabIndex        =   175
            Top             =   4440
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   55
            Left            =   4545
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   169
            Text            =   "0"
            Top             =   4005
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   53
            Left            =   4545
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   160
            Text            =   "0"
            Top             =   3390
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   52
            Left            =   4545
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   159
            Text            =   "0"
            Top             =   2655
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   49
            Left            =   4560
            MaxLength       =   3
            TabIndex        =   158
            Text            =   "0"
            Top             =   615
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   50
            Left            =   4530
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   156
            Text            =   "0"
            Top             =   1275
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   51
            Left            =   4530
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   155
            Text            =   "0"
            Top             =   1950
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   54
            Left            =   8730
            Locked          =   -1  'True
            TabIndex        =   157
            Text            =   "0"
            Top             =   6465
            Width           =   3465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Produzione Media kg/mq Anno (Km):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   68
            Left            =   435
            TabIndex        =   170
            Top             =   4005
            Width           =   3855
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "()"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   67
            Left            =   3225
            TabIndex        =   168
            Top             =   630
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   79
            Left            =   6915
            TabIndex        =   167
            Top             =   660
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Kg/Anno rifiuti smaltiti:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   65
            Left            =   435
            TabIndex        =   166
            Top             =   3390
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cmg al netto delle detrazioni:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   64
            Left            =   435
            TabIndex        =   165
            Top             =   2640
            Width           =   3060
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Detrazione spazzamento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   61
            Left            =   435
            TabIndex        =   164
            Top             =   615
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Entrate Tariffe Attese:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   66
            Left            =   5160
            TabIndex        =   163
            Top             =   6495
            Width           =   3510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Tot. al netto dello spazzamento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   62
            Left            =   420
            TabIndex        =   162
            Top             =   1290
            Width           =   4065
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Medio Generale  (Cmg) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   63
            Left            =   420
            TabIndex        =   161
            Top             =   1965
            Width           =   3270
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   7
      Left            =   45
      ScaleHeight     =   8910
      ScaleWidth      =   12600
      TabIndex        =   137
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari - Verifica Tariffe 7/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   11
         Left            =   60
         TabIndex        =   138
         Top             =   -45
         Width           =   12540
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   48
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   150
            Text            =   "0"
            Top             =   555
            Width           =   750
         End
         Begin VB.ComboBox cmb 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "FrmCalcoloTariffa.frx":000C
            Left            =   9750
            List            =   "FrmCalcoloTariffa.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   1005
            Width           =   2550
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   47
            Left            =   5820
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "0"
            Top             =   555
            Width           =   2265
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   46
            Left            =   5820
            Locked          =   -1  'True
            TabIndex        =   142
            Text            =   "0"
            Top             =   1020
            Width           =   2265
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   45
            Left            =   1185
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "0"
            Top             =   1020
            Width           =   2250
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   43
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   140
            Text            =   "0"
            Top             =   555
            Width           =   2250
         End
         Begin MSComctlLib.ListView lvw9 
            Height          =   7275
            Left            =   180
            TabIndex        =   139
            Top             =   1500
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   12832
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fattore Correttivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9915
            TabIndex        =   149
            Top             =   585
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Gettito Dichiarato:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   60
            Left            =   3960
            TabIndex        =   147
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Indice di Correzione:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   59
            Left            =   3975
            TabIndex        =   146
            Top             =   1065
            Width           =   1815
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Gettito UA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   58
            Left            =   195
            TabIndex        =   145
            Top             =   1065
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Gettito UD:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   56
            Left            =   195
            TabIndex        =   144
            Top             =   585
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   6
      Left            =   45
      ScaleHeight     =   8910
      ScaleWidth      =   12600
      TabIndex        =   126
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari - Utenze Attività 6/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   3
         Left            =   120
         TabIndex        =   127
         Top             =   -45
         Width           =   12540
         Begin MSComctlLib.ListView lvw8 
            Height          =   7290
            Left            =   180
            TabIndex        =   132
            Top             =   1500
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   12859
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   44
            Left            =   3165
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "0"
            Top             =   570
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   42
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   130
            Text            =   "0"
            Top             =   1035
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   41
            Left            =   10020
            Locked          =   -1  'True
            TabIndex        =   129
            Text            =   "0"
            Top             =   1035
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   40
            Left            =   10020
            Locked          =   -1  'True
            TabIndex        =   128
            Text            =   "0"
            Top             =   570
            Width           =   2295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Quota unitaria parte fissa (Quap):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   57
            Left            =   195
            TabIndex        =   136
            Top             =   585
            Width           =   2910
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Unitario al kg (Cu):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   55
            Left            =   195
            TabIndex        =   135
            Top             =   1065
            Width           =   2190
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Quota Variabile:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   54
            Left            =   7680
            TabIndex        =   134
            Top             =   1080
            Width           =   2070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Quota Parte Fissa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   53
            Left            =   7695
            TabIndex        =   133
            Top             =   600
            Width           =   2280
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   5
      Left            =   45
      ScaleHeight     =   8910
      ScaleWidth      =   12600
      TabIndex        =   109
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari - Utenze Attività 5/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   10
         Left            =   60
         TabIndex        =   110
         Top             =   -45
         Width           =   12540
         Begin MSComctlLib.ListView lvw7 
            Height          =   8295
            Left            =   180
            TabIndex        =   111
            Top             =   480
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   14631
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   4
      Left            =   45
      ScaleHeight     =   8910
      ScaleWidth      =   12600
      TabIndex        =   100
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari - Utenze Domestiche 4/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   9
         Left            =   60
         TabIndex        =   101
         Top             =   -45
         Width           =   12420
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   39
            Left            =   9810
            Locked          =   -1  'True
            TabIndex        =   124
            Text            =   "0"
            Top             =   585
            Width           =   2205
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   38
            Left            =   9810
            Locked          =   -1  'True
            TabIndex        =   122
            Text            =   "0"
            Top             =   1125
            Width           =   2205
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   35
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   "0"
            Top             =   1665
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   33
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "0"
            Top             =   1110
            Width           =   2310
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   32
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   103
            Text            =   "0"
            Top             =   570
            Width           =   2295
         End
         Begin MSComctlLib.ListView lvw6 
            Height          =   6630
            Left            =   180
            TabIndex        =   102
            Top             =   2160
            Width           =   12075
            _ExtentX        =   21299
            _ExtentY        =   11695
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Quota Parte Fissa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   52
            Left            =   7470
            TabIndex        =   125
            Top             =   600
            Width           =   2280
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Quota Variabile:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   51
            Left            =   7455
            TabIndex        =   123
            Top             =   1140
            Width           =   2070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Unitario al kg (Cu):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   44
            Left            =   195
            TabIndex        =   108
            Top             =   1695
            Width           =   2190
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Quota unitaria parte variabile (Quv):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   42
            Left            =   180
            TabIndex        =   107
            Top             =   1125
            Width           =   3135
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Quota unitaria parte fissa (Quf):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   41
            Left            =   195
            TabIndex        =   106
            Top             =   585
            Width           =   2715
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8790
      Index           =   3
      Left            =   45
      ScaleHeight     =   8790
      ScaleWidth      =   12600
      TabIndex        =   97
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari - Utenze Domestiche 3/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   8
         Left            =   60
         TabIndex        =   98
         Top             =   -45
         Width           =   12540
         Begin MSComctlLib.ListView lvw5 
            Height          =   8310
            Left            =   180
            TabIndex        =   99
            Top             =   480
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   14658
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8790
      Index           =   2
      Left            =   45
      ScaleHeight     =   8790
      ScaleWidth      =   12600
      TabIndex        =   81
      Top             =   45
      Width           =   12600
      Begin VB.Frame Frame1 
         Caption         =   "Tari 2/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   7
         Left            =   60
         TabIndex        =   82
         Top             =   -45
         Width           =   12540
         Begin VB.OptionButton optUtenzaDA 
            Caption         =   "&Utenza Domestica e Attività"
            Height          =   315
            Index           =   0
            Left            =   525
            TabIndex        =   96
            Top             =   8535
            Value           =   -1  'True
            Width           =   2325
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   29
            Left            =   9960
            TabIndex        =   94
            Text            =   "0"
            Top             =   1260
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   28
            Left            =   9960
            TabIndex        =   92
            Top             =   570
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   27
            Left            =   9945
            TabIndex        =   90
            Text            =   "0"
            Top             =   1935
            Width           =   1965
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   34
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "0"
            Top             =   1905
            Width           =   2430
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   31
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   84
            Text            =   "0"
            Top             =   1245
            Width           =   2445
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   30
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "0"
            Top             =   570
            Width           =   2430
         End
         Begin MSComctlLib.ListView lvw4 
            Height          =   5940
            Left            =   180
            TabIndex        =   89
            Top             =   2490
            Width           =   12165
            _ExtentX        =   21458
            _ExtentY        =   10478
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Image Image1 
            Height          =   285
            Index           =   0
            Left            =   195
            Stretch         =   -1  'True
            Top             =   8535
            Width           =   285
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "% Ripartizione Canone UA:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   38
            Left            =   7035
            TabIndex        =   95
            Top             =   1275
            Width           =   2895
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "% Ripartizione Canone UD:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   37
            Left            =   7035
            TabIndex        =   93
            Top             =   585
            Width           =   2910
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Passo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   36
            Left            =   9180
            TabIndex        =   91
            Top             =   1950
            Width           =   1350
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Entrate Tariffe Attese:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   43
            Left            =   165
            TabIndex        =   88
            Top             =   1920
            Width           =   3510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Entrate Tariffe Parte Variabile:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   40
            Left            =   180
            TabIndex        =   87
            Top             =   1260
            Width           =   3195
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Entrate Tariffe Parte Fissa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   39
            Left            =   195
            TabIndex        =   86
            Top             =   585
            Width           =   2850
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7710
      Index           =   1
      Left            =   45
      ScaleHeight     =   7710
      ScaleWidth      =   12330
      TabIndex        =   35
      Top             =   45
      Width           =   12330
      Begin VB.Frame Frame1 
         Caption         =   "Tari 1/7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7755
         Index           =   4
         Left            =   60
         TabIndex        =   36
         Top             =   -45
         Width           =   12285
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   37
            Left            =   9420
            TabIndex        =   113
            Text            =   "0"
            Top             =   2625
            Width           =   1050
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   36
            Left            =   9420
            TabIndex        =   112
            Text            =   "0"
            Top             =   3360
            Width           =   1035
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   26
            Left            =   9420
            Locked          =   -1  'True
            TabIndex        =   77
            Text            =   "0"
            Top             =   585
            Width           =   2520
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   25
            Left            =   9420
            Locked          =   -1  'True
            TabIndex        =   76
            Text            =   "0"
            Top             =   1260
            Width           =   2520
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   9420
            Locked          =   -1  'True
            TabIndex        =   75
            Text            =   "0"
            Top             =   1935
            Width           =   2520
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   23
            Left            =   8850
            Locked          =   -1  'True
            TabIndex        =   73
            Text            =   "0"
            Top             =   5040
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   22
            Left            =   8835
            Locked          =   -1  'True
            TabIndex        =   71
            Text            =   "0"
            Top             =   5745
            Width           =   2985
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   21
            Left            =   3105
            Locked          =   -1  'True
            TabIndex        =   69
            Text            =   "0"
            Top             =   1935
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   20
            Left            =   3105
            Locked          =   -1  'True
            TabIndex        =   67
            Text            =   "0"
            Top             =   1260
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   19
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "0"
            Top             =   6480
            Width           =   3465
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   18
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "0"
            Top             =   585
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   17
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   58
            Text            =   "0"
            Top             =   2640
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   16
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   57
            Text            =   "0"
            Top             =   3375
            Width           =   2280
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "CU Domestico:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   46
            Left            =   7755
            TabIndex        =   115
            Top             =   2655
            Width           =   1605
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "CU Attività:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   45
            Left            =   7755
            TabIndex        =   114
            Top             =   3390
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   50
            Left            =   10920
            TabIndex        =   119
            Top             =   2655
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   49
            Left            =   10845
            TabIndex        =   118
            Top             =   3390
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   48
            Left            =   10560
            TabIndex        =   117
            Top             =   2640
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   47
            Left            =   10560
            TabIndex        =   116
            Top             =   3360
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Parte Fissa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   35
            Left            =   7815
            TabIndex        =   80
            Top             =   615
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Parte Variabile:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   34
            Left            =   7830
            TabIndex        =   79
            Top             =   1290
            Width           =   1605
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costo Totale:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   18
            Left            =   7815
            TabIndex        =   78
            Top             =   1965
            Width           =   1410
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Entrate Tariffe Parte Fissa:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   33
            Left            =   4815
            TabIndex        =   74
            Top             =   5055
            Width           =   2850
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Entrate Tariffe Parte Variabile:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   32
            Left            =   4800
            TabIndex        =   72
            Top             =   5760
            Width           =   3195
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costi uso Capitale (CK):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   31
            Left            =   420
            TabIndex        =   70
            Top             =   1965
            Width           =   2505
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costi Comuni (CC):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   30
            Left            =   420
            TabIndex        =   68
            Top             =   1290
            Width           =   2010
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Entrate Tariffe Attese:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   29
            Left            =   4800
            TabIndex        =   66
            Top             =   6495
            Width           =   3510
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Costi di Gestione (CG):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   28
            Left            =   435
            TabIndex        =   65
            Top             =   615
            Width           =   2445
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Inflazione Programmata:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   27
            Left            =   435
            TabIndex        =   64
            Top             =   2640
            Width           =   2595
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Recupero Produttività:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   26
            Left            =   435
            TabIndex        =   63
            Top             =   3390
            Width           =   2355
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   20
            Left            =   5475
            TabIndex        =   62
            Top             =   3405
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   19
            Left            =   5475
            TabIndex        =   61
            Top             =   2670
            Width           =   210
         End
      End
   End
   Begin VB.PictureBox picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Index           =   4
      Left            =   45
      ScaleHeight     =   8910
      ScaleWidth      =   12570
      TabIndex        =   54
      Top             =   45
      Width           =   12570
      Begin VB.Frame Frame1 
         Caption         =   "Metodo Transitorio 4-4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   6
         Left            =   60
         TabIndex        =   55
         Top             =   -45
         Width           =   12405
         Begin MSComctlLib.ListView lvw3 
            Height          =   8280
            Left            =   210
            TabIndex        =   56
            Top             =   495
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   14605
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8670
      Index           =   3
      Left            =   45
      ScaleHeight     =   8670
      ScaleWidth      =   12450
      TabIndex        =   32
      Top             =   45
      Width           =   12450
      Begin VB.Frame Frame1 
         Caption         =   "Metodo Transitorio 3-3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   2
         Left            =   60
         TabIndex        =   33
         Top             =   -45
         Width           =   12420
         Begin MSComctlLib.ListView lvw2 
            Height          =   8430
            Left            =   210
            TabIndex        =   34
            Top             =   495
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   14870
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8790
      Index           =   2
      Left            =   45
      ScaleHeight     =   8790
      ScaleWidth      =   12330
      TabIndex        =   37
      Top             =   45
      Width           =   12330
      Begin VB.Frame Frame1 
         Caption         =   "Metodo Transitorio 2-3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   5
         Left            =   60
         TabIndex        =   38
         Top             =   -45
         Width           =   12270
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   15
            Left            =   9720
            MaxLength       =   15
            TabIndex        =   47
            Text            =   "0"
            Top             =   1530
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   14
            Left            =   9720
            MaxLength       =   15
            TabIndex        =   46
            Text            =   "0"
            Top             =   1065
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   2355
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "0"
            Top             =   600
            Width           =   2235
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   12
            Left            =   2340
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "0"
            Top             =   1545
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   2355
            MaxLength       =   15
            TabIndex        =   43
            Text            =   "0"
            Top             =   1065
            Width           =   2250
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   13
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   42
            Text            =   "0"
            Top             =   600
            Width           =   2280
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   6420
            Left            =   210
            TabIndex        =   41
            Top             =   2070
            Width           =   11880
            _ExtentX        =   20955
            _ExtentY        =   11324
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.OptionButton optUtenza 
            Caption         =   "&Utenza Attività"
            Height          =   315
            Index           =   1
            Left            =   2760
            TabIndex        =   40
            Top             =   8670
            Width           =   1725
         End
         Begin VB.OptionButton optUtenza 
            Caption         =   "&Utenza Domestica"
            Height          =   315
            Index           =   0
            Left            =   570
            TabIndex        =   39
            Top             =   8670
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tot.Tariffe Attese:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   25
            Left            =   7290
            TabIndex        =   53
            Top             =   600
            Width           =   1905
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Utenza Attivià:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   24
            Left            =   210
            TabIndex        =   52
            Top             =   1560
            Width           =   1545
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Utenza Domestica:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   23
            Left            =   210
            TabIndex        =   51
            Top             =   1080
            Width           =   2040
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Gettito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   22
            Left            =   225
            TabIndex        =   50
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tot.Tariffe Attese D:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   21
            Left            =   7320
            TabIndex        =   49
            Top             =   1080
            Width           =   2145
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Tot.Tariffe Attese A:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   17
            Left            =   7320
            TabIndex        =   48
            Top             =   1560
            Width           =   2130
         End
         Begin VB.Image Image1 
            Height          =   285
            Index           =   3
            Left            =   240
            Stretch         =   -1  'True
            Top             =   8670
            Width           =   285
         End
         Begin VB.Image Image1 
            Height          =   285
            Index           =   2
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   8670
            Width           =   285
         End
      End
   End
   Begin VB.PictureBox picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7710
      Index           =   1
      Left            =   45
      ScaleHeight     =   7710
      ScaleWidth      =   12210
      TabIndex        =   7
      Top             =   45
      Width           =   12210
      Begin VB.Frame Frame1 
         Caption         =   "Metodo Transitorio 1-3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7755
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   -45
         Width           =   12165
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   6135
            MaxLength       =   3
            TabIndex        =   23
            Text            =   "0"
            Top             =   3090
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   6135
            MaxLength       =   3
            TabIndex        =   22
            Text            =   "0"
            Top             =   2625
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   6135
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "0"
            Top             =   2175
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   6135
            MaxLength       =   3
            TabIndex        =   20
            Text            =   "0"
            Top             =   1725
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   6135
            MaxLength       =   3
            TabIndex        =   19
            Text            =   "0"
            Top             =   1260
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   6135
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   795
            Width           =   2970
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   4200
            Width           =   3465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(max 3,5)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   16
            Left            =   3450
            TabIndex        =   31
            Top             =   2655
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(max 1,5)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   15
            Left            =   4740
            TabIndex        =   30
            Top             =   2190
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(max 2,0)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   14
            Left            =   3810
            TabIndex        =   29
            Top             =   1740
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   13
            Left            =   9360
            TabIndex        =   28
            Top             =   1290
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   12
            Left            =   9360
            TabIndex        =   27
            Top             =   1755
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   9375
            TabIndex        =   26
            Top             =   3120
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   10
            Left            =   9375
            TabIndex        =   25
            Top             =   2640
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   9
            Left            =   9375
            TabIndex        =   24
            Top             =   2205
            Width           =   210
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Coefficiente incremento Entrate Tariffarie:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   8
            Left            =   270
            TabIndex        =   17
            Top             =   3135
            Width           =   4425
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Riduzione per Imballi Primari:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   270
            TabIndex        =   16
            Top             =   2655
            Width           =   3045
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Riduzione per Imballi Secondari e Terziari:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   270
            TabIndex        =   15
            Top             =   2190
            Width           =   4425
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Percentuale Traslazione Minima:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   270
            TabIndex        =   14
            Top             =   1740
            Width           =   3420
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Inflazione Programmata:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   270
            TabIndex        =   13
            Top             =   1260
            Width           =   2595
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Entrate Ruolo Anno Precedente:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   270
            TabIndex        =   12
            Top             =   825
            Width           =   4185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Totale Entrate Tariffe Attese:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   4665
            TabIndex        =   10
            Top             =   4200
            Width           =   3510
         End
      End
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8940
      Index           =   0
      Left            =   45
      ScaleHeight     =   8940
      ScaleWidth      =   12450
      TabIndex        =   3
      Top             =   45
      Width           =   12450
      Begin VB.Frame Frame1 
         Caption         =   "Scelta Applicazione Metodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   8955
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   -120
         Width           =   12315
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   870
            Width           =   3195
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   2865
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1560
            Width           =   3195
         End
         Begin VB.ComboBox DTPAnno 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2880
            TabIndex        =   176
            Top             =   2280
            Width           =   1935
         End
         Begin VB.OptionButton opt 
            Caption         =   "Integrale Applicazione del metodo"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   120
            Top             =   3975
            Width           =   8805
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   2
            Left            =   6795
            MaxLength       =   4
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Anno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   69
            Left            =   1065
            TabIndex        =   174
            Top             =   2295
            Width           =   1710
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Copertura"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9705
            TabIndex        =   152
            Top             =   1260
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Calcolo Tari"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   195
            TabIndex        =   121
            Top             =   3540
            Width           =   2385
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   9090
            TabIndex        =   9
            Top             =   1162
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Costo Totale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1065
            TabIndex        =   6
            Top             =   1680
            Width           =   1710
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Gettito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1065
            TabIndex        =   5
            Top             =   960
            Width           =   1710
         End
      End
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   495
      ScaleWidth      =   8655
      TabIndex        =   177
      Top             =   9000
      Width           =   8655
      Begin VB.CommandButton Command1 
         Caption         =   "&Avanti >"
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   5025
         TabIndex        =   183
         Top             =   100
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "< In&dietro"
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   3780
         TabIndex        =   182
         Top             =   100
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Annulla"
         Default         =   -1  'True
         Height          =   390
         Index           =   2
         Left            =   7530
         TabIndex        =   181
         Top             =   100
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "In Lire"
         Height          =   390
         Index           =   4
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   180
         Top             =   100
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Alla &Fine >>"
         Enabled         =   0   'False
         Height          =   390
         Index           =   5
         Left            =   6285
         TabIndex        =   179
         Top             =   100
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<< Alla &Inizio"
         Enabled         =   0   'False
         Height          =   390
         Index           =   6
         Left            =   2520
         TabIndex        =   178
         Top             =   100
         Width           =   1125
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostra &Riduzioni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   151
      Top             =   9015
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "FrmCalcoloTariffa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Txt Pagina Calcolo tipo metodo
Private Const PAG0_GETTITO_ID = 0
Private Const PAG0_TOTSERVIZIO_ID = 1
Private Const PAG0_TIPOTARIFFA_ID = 2

'option button metodo calcolo tariffa
Private Const METODOINTEGRALE_ID = tpoTariffa.TariffaNormalizzata
Private Const METODO507_ID = tpoTariffa.Tariffa507
Private Const METODOTRANSITORIO_ID = tpoTariffa.TariffaTransitorio
Private Const METODOPORTAAPORTA_ID = tpoTariffa.TariffaPortaAPorta
Private Const METODOCOSTIINDIVISIBILI_ID = tpoTariffa.TariffaCostiIndivisibili

'Button
Private Const AVANTI = 0
Private Const INDIETRO = 1
Private Const ANNULLA = 2
Private Const MOSTRARIDUZIONI = 3
Private Const MOSTRAINVALUTA = 4
Private Const ALLAFINE = 5
Private Const ALLINIZIO = 6


'Txt Tariffa Metodo Transitorio pag 1
Private Const PAG1_T_ET1_ID = 3
Private Const PAG1_T_ET0_ID = 4
Private Const PAG1_T_IP_ID = 5
Private Const PAG1_T_X_ID = 6
Private Const PAG1_T_Y_ID = 7
Private Const PAG1_T_CCON_ID = 8
Private Const PAG1_T_K_ID = 9

'Txt Tariffa Metodo Transitorio pag 2
Private Const PAG2_T_GETTITO_ID = 10
Private Const PAG2_T_UDOMESTICA_ID = 11
Private Const PAG2_T_UATTIVITA_ID = 12
Private Const PAG2_T_ET1_ID = 13
Private Const PAG2_T_ET1D_ID = 14
Private Const PAG2_T_ET1A_ID = 15


'Txt Tariffa Metodo Normalizzato pag 1
Private Const PAG1_N_ET_ID = 19
Private Const PAG1_N_ETF_ID = 23
Private Const PAG1_N_ETV_ID = 22
Private Const PAG1_N_IP_ID = 17
Private Const PAG1_N_RP_ID = 16
Private Const PAG1_N_CK_ID = 21
Private Const PAG1_N_CC_ID = 20
Private Const PAG1_N_CG_ID = 18
Private Const PAG1_N_CT_ID = 24
Private Const PAG1_N_PV_ID = 25
Private Const PAG1_N_PF_ID = 26
Private Const PAG1_N_CUD_ID = 37
Private Const PAG1_N_CUA_ID = 36

'lbl Tariffa Metodo Normalizzato pag 1
Private Const PAG1_N_LBL_ETF = 33
Private Const PAG1_N_LBL_ETV = 32
Private Const PAG1_N_LBL_PF = 35
Private Const PAG1_N_LBL_PV = 34
Private Const PAG1_N_LBL_CUA = 49
Private Const PAG1_N_LBL_CUD = 50

'txt Tariffa Metodo Normalizzato pag 2
Private Const PAG2_N_ET_ID = 34
Private Const PAG2_N_ETF_ID = 30
Private Const PAG2_N_ETV_ID = 31
Private Const PAG2_N_PASSO_ID = 27
Private Const PAG2_N_PERCINIUD_ID = 28
Private Const PAG2_N_PERCINIUA_ID = 29
'lbl Tariffa Metodo Normalizzato pag 2
Private Const PAG2_N_LBL_ETF = 39
Private Const PAG2_N_LBL_ETV = 40


'txt Pagina  Fine calcolo tariffa UD pag 4
Private Const PAG4_QUF_ID = 32
Private Const PAG4_QUV_ID = 33
Private Const PAG4_TotQF_ID = 39
Private Const PAG4_TotQV_ID = 38
Private Const PAG4_CU_ID = 35

'Risultati calcolo tariffa metodo normalizzato UD
Private QUV As Double, CUA As Double, CUD As Double, CUASave As Double
Private QUFUD As Double, TotQFUD As Double, ToTQVUD As Double, ToTQVUA As Double, ToTQFUA As Double
Private SupCorrettaConRiduzioneUD() As Double, QUAPUA As Double
Private SupCorrettaSenzaRiduzioneUD() As Double
Private NumeroUtenzaDomesticaConRiduzione() As Double
Private SupConRiduzioneUAPF() As Double
Private SupConRiduzioneUAPV() As Double

Private GettitoCalcUDPF As Double, GettitoCalcUDPV As Double
Private GettitoCalcUAPV As Double, GettitoCalcUAPF As Double
Private GettitoCalc507UD As Double, GettitoCalc507UA As Double
Private CoeffCorrettivo507IPS As Double
Private CoeffCorrTariffa507 As Double

'con riduzioni
Private GettitoCalcUDPFRidotto As Double, GettitoCalcUDPVRidotto As Double
Private GettitoCalcUAPVRidotto As Double, GettitoCalcUAPFRidotto As Double


'Riduzioni
Private TotMancatoIntroitoDAUA_PF As Double
Private TotMancatoIntroitoDAUD_PF As Double
Private TotMancatoIntroitoDAUA_PV As Double
Private TotMancatoIntroitoDAUD_PV  As Double
Private TotMancatoIntroito As Double
Private AlfaPrec As Double
Private QUFUDsav As Double
Private CUDsav As Double
Private CUAsav As Double
Private QUAPUAsav As Double
Private QUVsav As Double
'ANTONIO 26-02-02Private mRiduzione As Boolean
'ANTONIO 26-02-02Private pQRifiutiUDRiduzioni As Double
'ANTONIO 26-02-02Private pQRifiutiUARiduzioni As Double

'507
Private CostoTotAlNettoSpazzamento As Double
Private CalcoloCostoCopertura  As Double
Private CMG As Double
Private KM As Double
Private TotMancatoIntroito507 As Double
Private SuperficiCorrette As Long


'txt Pagina  Fine calcolo tariffa UA pag 6
Private Const PAG6_QUAP_ID = 44
Private Const PAG6_TOTQF_ID = 40
Private Const PAG6_TOTQV_ID = 41
Private Const PAG6_CU_ID = 42


'txt Pagina  Fattore Correttivo
Private Const PAG7_GETTITOUD_ID = 43
Private Const PAG7_GETTITOUA_ID = 45
Private Const PAG7_GETTITODICHIARATO_ID = 47
Private Const PAG7_GUADAGNO_ID = 46
Private Const PAG7_FATTORECORRETTIVO_ID = 48


'txt calcolo 507 Pagina 1
Private Const PAG1_PERCDETRASPAZZAMENTO_ID = 49
Private Const PAG1_COSTOTOTALNETTO_ID = 50
Private Const PAG1_CMG_ID = 51
Private Const PAG1_CMGCONRIDUZIONI_ID = 52
Private Const PAG1_COSTOTOTRIFIUTI_ID = 53
Private Const PAG1_COSTOCOPERTURA_ID = 54
Private Const PAG1_KM_ID = 55

'PAGINA id
Private Const PAG0_CALCOLOMETODO = 0
Private Const PAG1_CALCOLOET1 = 1
Private Const PAG2_DISTRIBUZIONEUTENZE = 2
Private Const PAG3_TARIFFACATEGORIE = 3
Private Const PAG4_TARIFFAUTENZEDOMESTICHE = 4
Private Const PAG1_CALCOLOINTEGRALE = 1
Private Const PAG2_SUDDIVISIONEPARTEFISSA = 2
Private Const PAG3_TARIFFAUDNORMALIZZATE = 3
Private Const PAG4_TARIFFAUDNORMALIZZATE_END = 4
Private Const PAG5_TARIFFAUANORMALIZZATE = 5
Private Const PAG6_TARIFFAUANORMALIZZATE_END = 6
Private Const PAG7_CORREZIONETARIFFE = 7
Private Const PAG1_CALCOLO507 = 1
Private Const PAG2_CALCOLO507RISULTATO = 2

'MIN lbl Value
Private Const MINTRASLMIN_ID = 14
Private Const MINRIDIMB23_ID = 15
Private Const MINRIDIMB1_ID = 16
Private Const MINMAXPERCDETRAZIONE_ID = 67

'Varibili MIN
Private mMINTRASLMIN As Double
Private mMINRIDIMB23 As Double
Private mMINRIDIMB1 As Double
Private mMAXPERCDETRAZIONE As Double
Private mMINPERCDETRAZIONE As Double


'Option Button scelta colonna utenza
Private Const optColonnaDomestica = 0
Private Const optColonnaAttivita = 0

'ColonnaUDomestica
Private keyColUDomestica As String
'ColonnaUAttivita
Private keyColUAttivita As String

'ColonnaUDA
Private keyColUDA As String

'Metodo tariffa
Private mOptMetodoTariffa As tpoTariffa

'Pagina corrente
Private mPagina As Integer
Private PicPage As Object
Private mPicName As String


'Variabili Globali per identificare di quale comune si stà cacolando la tariffa
Private mCodComune As String
Private mDesComune As String
Private mAnnoRiferimento As String
Private mParAbitanti As String

'Recorset
Private rsDatiTecniciComune As Recordset
Private rsDatiAnagraficiComune As Recordset
Private rsCostiComune As Recordset
Private rsRIDUA As Recordset
Private rsRIDUD As Recordset

Private OldTxtValue As String
Private Const TAG_TAR_PVUA_FLD = SEPARATORE & "PV"
Private Const TAG_TAR_PFUA_FLD = SEPARATORE & "PF"
Private Const TAG_TAR_PVUD_FLD = SEPARATORE & "PV"
Private Const TAG_TAR_PFUD_FLD = SEPARATORE & "PF"
Private CUAPerc As Double
Private CUDPerc As Double
Private idPicture As Long
Private mCoeffMedioSvuotamentoLT As Double
Private mDisabilitaEvento As Boolean
Private mUscitaLoopAllFine As Boolean


Public Sub ImpostaCampi507(value As String)
    Txt(PAG1_PERCDETRASPAZZAMENTO_ID).Text = value
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
    LetParamValue mAnnoRiferimento, mCodComune, "PERCDETRAZIONE", True, Txt(PAG1_PERCDETRASPAZZAMENTO_ID).Text
End Sub

Public Sub ImpostaCampiNormalizzato(ValueIP As String, ValueRP As String, ValuePERCINIUD As String, ValueCUA As String, ValueCUD As String)
    
    Txt(PAG1_N_IP_ID).Text = ValueIP
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
    Txt(PAG1_N_RP_ID).Text = ValueRP
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
    Txt(PAG2_N_PERCINIUD_ID).Text = ValuePERCINIUD
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
    Txt(PAG1_N_CUA_ID).Text = ValueCUA
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
    Txt(PAG1_N_CUD_ID).Text = ValueCUD
    txt_LostFocus PAG1_PERCDETRASPAZZAMENTO_ID
End Sub


Public Sub VaiAllaFine()
    Me.Visible = False
    
    Command1_Click (ALLINIZIO)
    Command1_Click (ALLAFINE)
    
    Me.Visible = True
    
    If Not rsRIDUA.EOF Or Not rsRIDUD.EOF Then Command1_Click (MOSTRARIDUZIONI)
    
End Sub

Function Avvio(Optional sCodComune As String = vbNullString, Optional AnnoRiferimento As String, Optional tipoCalcolo As tpoTariffa = TariffaNonImpostata) As Boolean
    Dim rs As Recordset
    Info ""
    
    Avvio = False
    
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        'DataIns = DateAdd("yyyy", -1, Now)
        AnnoRiferimento = IIf(AnnoRiferimento = vbNullString, Year(Now), AnnoRiferimento)

        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
   
    
    mOptMetodoTariffa = tipoCalcolo
    mAnnoRiferimento = AnnoRiferimento
    mCodComune = sCodComune
    mDesComune = getDesComune(mCodComune)
        
    Set rsCostiComune = DatiCostiComune(mCodComune, mAnnoRiferimento)
    If rsCostiComune.EOF Then Info ("Occorre inserire i Costi comunali del " & mAnnoRiferimento & " - Prima di effettuale il calcolo della tariffa")
    
    Set rsDatiTecniciComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    If rsDatiTecniciComune.EOF Then
        Info ("Occorre inserire i Dati Tecnici Comunali del " & mAnnoRiferimento & " - Prima di effettuale il calcolo della tariffa")
    Else
        If rsDatiTecniciComune("QRifiutiTotali") = 0 Then
            Info ("Occorre inserire i Dati Tecnici Comunali del " & mAnnoRiferimento & " - Prima di effettuale il calcolo della tariffa")
        Else
            If rsDatiTecniciComune("QRifiutiCategorieAttivita") = 0 Then Info ("Verificare che i KC e KA e i dati relativi alle Categoire Attività siano stati inseriti" & " - Prima di effettuale il calcolo della tariffa")
        End If
    End If
    
    
    
    Set rsDatiAnagraficiComune = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
    If rsDatiAnagraficiComune.EOF Then Info ("Occorre inserire i Dati Anagrafici Comunali del " & mAnnoRiferimento & " - Prima di effettuale il calcolo della tariffa")
    
    If DTPAnno.ListCount <= 0 And mOptMetodoTariffa = TariffaNonImpostata Then
        CaricaComboAnniTariffazione
    End If
    
    
    If tipoCalcolo = TariffaNonImpostata Then Me.Show
    
    picButton.Visible = tipoCalcolo = TariffaNonImpostata

    
    'Mi posiziono sulla prima Picture
    Command1_Click (1)

    If tipoCalcolo <> TariffaNonImpostata Then opt_Click (tipoCalcolo)

    mCoeffMedioSvuotamentoLT = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFMEDSVULT")
    chkTarsu507perNComponenti.value = GetParamValue(mAnnoRiferimento, mCodComune, "CHKCALPERSINGUD")


    Avvio = True

End Function

Private Sub MetodoTariffazione()
    Dim lastTop As Long
    
    mDisabilitaEvento = True
    If opt.UBound = 0 Then
        Dim rs As ADODB.Recordset
        Set rs = gDBC.Execute("select * from tipitariffa order by codTipoTariffa")
        lastTop = opt(opt.UBound).Top
        Do While Not rs.EOF
        
            If rs!codTipoTariffa > 0 Then
                Load opt(opt.UBound + 1)
                'Set opt(opt.UBound).Parent = opt(0).Parent
                If rs!visibile Then
                    opt(opt.UBound).Top = lastTop + opt(opt.UBound).Height + 15
                    lastTop = opt(opt.UBound).Top
                End If
            End If
            opt(opt.UBound).Caption = rs!Descrizione
            opt(opt.UBound).value = rs!codTipoTariffa
            opt(opt.UBound).Visible = rs!visibile
            
            rs.MoveNext
        Loop
    End If
    Dim i As Integer
    If CDbl(Txt(PAG0_TOTSERVIZIO_ID)) = 0 Or CDbl(Txt(PAG0_GETTITO_ID)) = 0 Then
        Command1(AVANTI).Enabled = False
        Command1(ALLAFINE).Enabled = False
        For i = opt.LBound To opt.UBound
            opt(i).Enabled = False
        Next

        Exit Sub
    Else
        For i = opt.LBound To opt.UBound
            opt(i).Enabled = True
        Next
    End If
    
    Txt(PAG0_TIPOTARIFFA_ID).Text = 100
    Txt(PAG0_TIPOTARIFFA_ID).Text = Format$(CDbl(GetParamValue(mAnnoRiferimento, mCodComune, "PERCGETTITO")), "##0.00")
    mOptMetodoTariffa = GetParamValue(mAnnoRiferimento, mCodComune, "TIPOMETODO")
    
    mDisabilitaEvento = False
    
    opt(mOptMetodoTariffa).value = True
    Command1(AVANTI).Enabled = True
    opt_Click (mOptMetodoTariffa)
End Sub

Private Sub cmb_Click()
    LetParamValue mAnnoRiferimento, mCodComune, "TIPOCORREZIONE", True, cmb.ListIndex
    CacoloTipoAdeguamento cmb.ListIndex, vbNullString
End Sub

Private Sub DTPAnno_click()
    Dim Anno As String
    
    Anno = mAnnoRiferimento
    mAnnoRiferimento = DTPAnno.List(DTPAnno.ListIndex)
    
    If Not Avvio(mCodComune, mAnnoRiferimento) Then
        mAnnoRiferimento = Anno
        DTPAnno.Text = mAnnoRiferimento
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
    Info ""
End Sub



Private Sub opt_Click(index As Integer)
    If mDisabilitaEvento Then Exit Sub
    Select Case index
        Case METODOINTEGRALE_ID, METODOPORTAAPORTA_ID, METODOCOSTIINDIVISIBILI_ID
            Set PicPage = Picture3
            mPicName = "Picture3"
            Txt(PAG0_TIPOTARIFFA_ID).Locked = True
            Txt(PAG0_TIPOTARIFFA_ID).Text = 100
        Case METODOTRANSITORIO_ID
            Set PicPage = Picture2
            mPicName = "Picture2"
            Txt(PAG0_TIPOTARIFFA_ID).Locked = False
        Case METODO507_ID
            Txt(PAG0_TIPOTARIFFA_ID).Locked = False
            Set PicPage = Picture4
            mPicName = "Picture4"
    End Select
    
    'salvataggio TIPO METODO e PECENTUALE RIPARTIZIONE
    LetParamValue mAnnoRiferimento, mCodComune, "TIPOMETODO", True, index
    mOptMetodoTariffa = index
    
    LetParamValue mAnnoRiferimento, mCodComune, "PERCGETTITO", True, Txt(PAG0_TIPOTARIFFA_ID).Text
    
    
    'Carico le riduzioni se vi sono per UA e UD
    Set rsRIDUA = SelezionaBonus(BonusUtenzaAttivita, mCodComune, mAnnoRiferimento, mOptMetodoTariffa)
    
    Set rsRIDUD = SelezionaBonus(BonusUtenzaDomestica, mCodComune, mAnnoRiferimento, mOptMetodoTariffa)
    
    
    
End Sub

Public Sub SalvaTariffa()
    If mOptMetodoTariffa = mOptMetodoTariffa = TariffaNormalizzata _
                        Or mOptMetodoTariffa = TariffaCostiIndivisibili _
                        Or mOptMetodoTariffa = TariffaPortaAPorta Then
        'scrivo tariffe sul MDB
        SalvaTariffeMetodoIntegrale
    ElseIf mOptMetodoTariffa = Tariffa507 Then
        SalvaTariffeMetodo507
    End If

End Sub

Private Sub Command1_Click(index As Integer)
    Dim i As Integer
    Dim ii As Integer
    Dim pic  As PictureBox
    Dim Cancel As Boolean
    'valuta Corrente
    Dim valutaCorrente As String
    
    
    Select Case index
        Case MOSTRARIDUZIONI
            If Command1(MOSTRARIDUZIONI).Caption = "Applica &Riduzioni" Then
                If mOptMetodoTariffa = Tariffa507 Then
                    'RicalcolaTariffeConLeRiduzioni507
                ElseIf mOptMetodoTariffa = TariffaNormalizzata Or _
                    mOptMetodoTariffa = TariffaCostiIndivisibili Then
                    
                    'applico in un colpo solo RicalcolaTariffeConLeRiduzioni
                
                        TotMancatoIntroito = 0
                        AlfaPrec = 0
                                        
                        MostraMancatoIntroitoConLeRiduzioni
                        RicalcolaTariffeConLeRiduzioni
                        MostraMancatoIntroitoConLeRiduzioni
                        RicalcolaTariffeConLeRiduzioni
                
                
                End If
                Command1(MOSTRARIDUZIONI).Visible = False
            Else
                Select Case mPagina
                
'ANTONIO 26-02-02 Case PAG3_TARIFFAUDNORMALIZZATE
'                    TariffeUtenzeDomesticheNormalizzato True
'                    Command1(MOSTRARIDUZIONI).Visible = False
'                    mRiduzione = True
                    
                Case Else
                    If mOptMetodoTariffa = Tariffa507 Then
                        'AS07042003
                        'MostraMancatoIntroitoConLeRiduzioni507
                    ElseIf mOptMetodoTariffa = TariffaNormalizzata Or _
                           mOptMetodoTariffa = TariffaCostiIndivisibili Then
                        TotMancatoIntroito = 0
                        AlfaPrec = 0
                                        
                        MostraMancatoIntroitoConLeRiduzioni
                        RicalcolaTariffeConLeRiduzioni
                        MostraMancatoIntroitoConLeRiduzioni
                        'Command1(MOSTRARIDUZIONI).Visible = False
                        'RicalcolaTariffeConLeRiduzioni
                    End If
                End Select
            End If
            
            Command1(AVANTI).Default = True
            If Command1(AVANTI).Visible Then Command1(AVANTI).SetFocus
            Exit Sub
    
        Case MOSTRAINVALUTA
            
            valutaCorrente = IIf(Command1(MOSTRAINVALUTA).Caption = "In Lire", VALUTA_LIRE, VALUTA_EURO)
            
            Command1(MOSTRARIDUZIONI).Enabled = gValutaDefault = valutaCorrente
            
            If mOptMetodoTariffa = Tariffa507 Then
                MettiInValuta lvw10, valutaCorrente, 5
                lvw10.ColumnHeaders(2).Text = IIf(valutaCorrente = VALUTA_EURO, "Tari €/mq anno", "Tari £/mq anno")
                lvw10.ColumnHeaders(5).Text = IIf(valutaCorrente = VALUTA_EURO, "Tari €/mq anno", "Tari £/mq anno")
                
            Else
                Txt(PAG7_GETTITOUD_ID) = InValutaEst(Txt(PAG7_GETTITOUD_ID), valutaCorrente) 'InValutaEst(GettitoCalcUDPF + GettitoCalcUDPV, valuta)
                Txt(PAG7_GETTITOUA_ID) = InValutaEst(Txt(PAG7_GETTITOUA_ID), valutaCorrente)  'InValutaEst(GettitoCalcUAPF + GettitoCalcUAPV, valuta)
                Txt(PAG7_GETTITODICHIARATO_ID) = InValutaEst(Txt(PAG7_GETTITODICHIARATO_ID), valutaCorrente) 'InValutaEst(txt(PAG1_N_ET_ID), valuta)
                
                MettiInValuta lvw9, valutaCorrente, 5
            End If
            
            Command1(MOSTRAINVALUTA).Caption = IIf(Command1(MOSTRAINVALUTA).Caption = "In Lire", "In Euro", "In Lire")
            Command1(AVANTI).Enabled = IIf(Command1(MOSTRAINVALUTA).Caption = "In Lire", VALUTA_EURO, VALUTA_LIRE) = gValutaDefault
            Exit Sub
    End Select

    If index = AVANTI Then
        If Not Me.ActiveControl Is Nothing Then txt_Validate Me.ActiveControl.index, Cancel
        If Cancel = True Then Exit Sub
    End If

    For i = Picture1.LBound To Picture1.UBound
        Picture1(i).Visible = False
    Next
    
    For i = Picture2.LBound To Picture2.UBound
        Picture2(i).Visible = False
    Next
    
    For i = Picture3.LBound To Picture3.UBound
        Picture3(i).Visible = False
    Next
    
   
    mUscitaLoopAllFine = False
    For i = Picture4.LBound To Picture4.UBound
        Picture4(i).Visible = False
    Next
    
    Select Case index
        Case AVANTI
            If Command1(AVANTI).Caption = "&Salva" Then
                If mOptMetodoTariffa = TariffaNormalizzata _
                    Or mOptMetodoTariffa = TariffaCostiIndivisibili _
                    Or mOptMetodoTariffa = TariffaPortaAPorta Then
                    'scrivo tariffe sul MDB
                    SalvaTariffeMetodoIntegrale
                ElseIf mOptMetodoTariffa = Tariffa507 Then
                    SalvaTariffeMetodo507
                End If
                
                idPicture = 0
                Unload Me
                Exit Sub
            Else
                
                Command1(INDIETRO).Enabled = True
                If idPicture + 1 < PicPage.Count Then
                    idPicture = idPicture + 1
                    Command1(INDIETRO).Enabled = True
                    Command1(ALLINIZIO).Enabled = True
                Else
                    idPicture = PicPage.Count
                    Command1(AVANTI).Caption = "&Salva"
                    Command1(ALLAFINE).Enabled = False
                End If
            End If
        Case INDIETRO
            Command1(AVANTI).Enabled = True
            Command1(ALLAFINE).Enabled = True
            Command1(AVANTI).Caption = "&Avanti >>"
            If idPicture - 1 > PAG0_CALCOLOMETODO Then
                idPicture = idPicture - 1
            Else
                idPicture = PAG0_CALCOLOMETODO
                Set PicPage = Picture1
                idPicture = 0
                mPicName = "Picture1"
                Command1(INDIETRO).Enabled = False
                Command1(ALLINIZIO).Enabled = False
            End If
        Case ALLAFINE
            Do While Command1(AVANTI).Caption <> "&Salva" And mUscitaLoopAllFine = False
                Command1_Click AVANTI
               
            Loop
            Exit Sub
        Case ALLINIZIO
            Do While Command1(INDIETRO).Enabled
                Command1_Click INDIETRO
            Loop
            Exit Sub
        Case ANNULLA
            idPicture = 0
            Unload Me
            Exit Sub

    End Select
    '
    mPagina = idPicture
    
    Set pic = PicPage(idPicture)
    If Not CaricaValori Then
        If Me.Visible = True Then Command1_Click (ANNULLA)
    ElseIf Not FaiICalcoli(index) Then
            If Me.Visible = True Then
                    Command1_Click (INDIETRO)
                    Command1(INDIETRO).Enabled = False
                    Command1(AVANTI).Enabled = False
                    Command1(ALLAFINE).Enabled = False
                    Command1(ALLINIZIO).Enabled = False
                    mUscitaLoopAllFine = True
            End If
        Else
            pic.Visible = True
    End If
End Sub


Private Sub CaricaComboAnniTariffazione()
    Dim rs As Recordset
    
    Set rs = gDBC.Execute("SELECT DISTINCT AnagraficaComunale.AnnoRiferimento, AnagraficaComunale.flgSimulazione " & _
                         " From AnagraficaComunale " & _
                         " WHERE AnagraficaComunale.CodComune='" & mCodComune & "' AND " & _
                         " AnagraficaComunale.NumAbitantiResidenti>0 AND AnagraficaComunale.flgSimulazione=False ")
                         
    DTPAnno.Clear
    Do While Not rs.EOF
        DTPAnno.AddItem rs(0)
        rs.MoveNext
    Loop
    
    If DTPAnno.ListCount > 0 Then
        DTPAnno.ListIndex = DTPAnno.ListCount - 1
    End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    idPicture = 0
 
    
    MDIMain.mnuComprimi.Enabled = False
    'centra form
    CentraME Me
    
    
    Set PicPage = Picture1
    mPicName = "Picture1"
    
    keyColUDomestica = vbNullString
    keyColUAttivita = vbNullString
    keyColUDA = vbNullString

        
    lvw.ColumnHeaderIcons = MDIMain.imglist
    lvw.Icons = MDIMain.imglist
    lvw.SmallIcons = MDIMain.imglist

    lvw2.ColumnHeaderIcons = MDIMain.imglist
    lvw2.Icons = MDIMain.imglist
    lvw2.SmallIcons = MDIMain.imglist

    lvw3.ColumnHeaderIcons = MDIMain.imglist
    lvw3.Icons = MDIMain.imglist
    lvw3.SmallIcons = MDIMain.imglist
    
    lvw4.ColumnHeaderIcons = MDIMain.imglist
    lvw4.Icons = MDIMain.imglist
    lvw4.SmallIcons = MDIMain.imglist

    lvw5.ColumnHeaderIcons = MDIMain.imglist
    lvw5.Icons = MDIMain.imglist
    lvw5.SmallIcons = MDIMain.imglist

    lvw6.ColumnHeaderIcons = MDIMain.imglist
    lvw6.Icons = MDIMain.imglist
    lvw6.SmallIcons = MDIMain.imglist
    
    lvw7.ColumnHeaderIcons = MDIMain.imglist
    lvw7.Icons = MDIMain.imglist
    lvw7.SmallIcons = MDIMain.imglist
    
    lvw8.ColumnHeaderIcons = MDIMain.imglist
    lvw8.Icons = MDIMain.imglist
    lvw8.SmallIcons = MDIMain.imglist
    
    lvw9.ColumnHeaderIcons = MDIMain.imglist
    lvw9.Icons = MDIMain.imglist
    lvw9.SmallIcons = MDIMain.imglist
    
    lvw10.ColumnHeaderIcons = MDIMain.imglist
    lvw10.Icons = MDIMain.imglist
    lvw10.SmallIcons = MDIMain.imglist
    
    
    For i = Picture1.LBound To Picture1.UBound
        Picture1(i).Visible = False
        Picture1(i).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picButton.Height - 250
    Next
    
    For i = Picture2.LBound To Picture2.UBound
        Picture2(i).Visible = False
        Picture2(i).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picButton.Height - 250
    Next
    
    For i = Picture3.LBound To Picture3.UBound
        Picture3(i).Visible = False
        Picture3(i).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picButton.Height - 250
    Next

    For i = Picture4.LBound To Picture4.UBound
        Picture4(i).Visible = False
        Picture4(i).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - picButton.Height - 250
    Next
    
     
    For i = Frame1.LBound To Frame1.UBound
        Frame1(i).Move 0, 0, Picture1(0).ScaleWidth - 30, Picture1(0).ScaleHeight - 30
    Next


    'sistema max_len
    For i = Txt.LBound To Txt.UBound
        txt_GotFocus (i)
    Next



    Me.Caption = "Calcolo Tari: " & mDesComune
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

        
    If optUtenza(optColonnaDomestica) Then
        If keyColUDomestica = keyColUAttivita And keyColUAttivita <> vbNullString Then
                lvw.ColumnHeaders(keyColUAttivita).Icon = ICON_ATTIVITA
        ElseIf keyColUDomestica <> vbNullString Then
            lvw.ColumnHeaders(keyColUDomestica).Icon = ICON_NULL
        End If
        ColumnHeader.Icon = IIf(ColumnHeader.Key = keyColUAttivita, ICON_DOMESTICA_ATTIVITA, ICON_DOMESTICA)
        keyColUDomestica = ColumnHeader.Key
    Else
        If keyColUDomestica = keyColUAttivita And keyColUAttivita <> vbNullString Then
                lvw.ColumnHeaders(keyColUDomestica).Icon = ICON_DOMESTICA
        ElseIf keyColUAttivita <> vbNullString Then
            lvw.ColumnHeaders(keyColUAttivita).Icon = ICON_NULL
        End If
        ColumnHeader.Icon = IIf(ColumnHeader.Key = keyColUDomestica, ICON_DOMESTICA_ATTIVITA, ICON_ATTIVITA)
        keyColUAttivita = ColumnHeader.Key
    End If

End Sub



Private Sub lvw4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If ColumnHeader.index = 1 Then Exit Sub
    If optUtenzaDA(0) Then
        If keyColUDA <> vbNullString Then
            On Error Resume Next
            lvw4.ColumnHeaders(keyColUDA).Icon = ICON_NULL
            lvw4.ColumnHeaders(keyColUDA).Alignment = lvwColumnLeft
            lvw4.ColumnHeaders(keyColUDA).Alignment = lvwColumnRight
            On Error GoTo 0
        End If
        ColumnHeader.Icon = ICON_DOMESTICA_ATTIVITA
        
        keyColUDA = ColumnHeader.Key
        
        LetParamValue mAnnoRiferimento, mCodComune, "COLUDA", True, keyColUDA
        
    End If

 End Sub




Private Sub txt_Change(index As Integer)
        
    If index = PAG0_TIPOTARIFFA_ID Then
        
        If Txt(index) = vbNullString Then Txt(index) = 0: Txt(index).SetFocus: Exit Sub
        'Il valore deve essere numerico
        If Not IsNumeric(Txt(index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Txt(index) = OldTxtValue: Txt(index).SetFocus: Exit Sub
        
        Txt(PAG0_GETTITO_ID) = Format$((CDbl(Txt(PAG0_TOTSERVIZIO_ID)) * CDbl(Txt(PAG0_TIPOTARIFFA_ID))) / 100, FORMATVALUTA)
    End If

End Sub

Private Sub txt_GotFocus(index As Integer)
    If Txt(index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub
    
    'Imposto il numero MIN di caratteri in EDIT MODE
    If Txt(index).Text <> vbNullString Then Txt(index) = CDbl(Txt(index))
    'se Percentuale
    Select Case index
        Case PAG1_T_IP_ID, PAG1_T_X_ID, _
               PAG1_T_Y_ID, PAG1_T_CCON_ID, PAG1_T_K_ID, _
               PAG1_N_IP_ID, PAG1_N_RP_ID
                
                Txt(index).MaxLength = 4
        Case Else
                Txt(index).MaxLength = 15
    End Select
    
    DoEvents
    Txt(index).SelStart = 0
    Txt(index).SelLength = Len(Txt(index))
    OldTxtValue = Txt(index)
End Sub


Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_LostFocus(index As Integer)
            
    If Txt(index).Locked Then Exit Sub
    
    'Imposto il numero MIN di caratteri in EDIT MODE
    Txt(index) = CDbl(Txt(index))
    Txt(index).MaxLength = 15
    

    'se Percentuale
    Select Case index
        Case PAG1_T_IP_ID, PAG1_T_X_ID, _
              PAG1_T_Y_ID, PAG1_T_CCON_ID, _
              PAG1_T_K_ID, _
              PAG1_N_IP_ID, PAG1_N_RP_ID, _
              PAG2_N_PASSO_ID, PAG2_N_PERCINIUD_ID, _
              PAG2_N_PERCINIUA_ID, _
              PAG0_TIPOTARIFFA_ID, _
              PAG1_PERCDETRASPAZZAMENTO_ID, _
              PAG1_N_CUD_ID, _
              PAG1_N_CUA_ID

              
            If index = PAG2_N_PASSO_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PASSOPF", True, Txt(index).Text
            End If
            
            
            If index = PAG2_N_PERCINIUD_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PERCINIUD", True, Txt(index).Text
                Txt(index) = Format$(CDbl(Txt(index)), "##0.000000")
            End If
              
            If index = PAG2_N_PERCINIUA_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PERCINIUA", True, Txt(index)
                Txt(index) = Format$(CDbl(Txt(index)), "##0.000000")
            End If
                
            
            If index = PAG1_T_X_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "TRASLMIN", False, Txt(index)
            End If
            
            If index = PAG1_T_Y_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "RIDIMB23", True, Txt(index)
            End If
            
            If index = PAG1_T_CCON_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "RIDIMB1", True, Txt(index)
            End If
        
            If index = PAG1_T_K_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "COEFFINC", True, Txt(index)
            End If
        
            
            If index = PAG1_T_IP_ID Or index = PAG1_N_IP_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "INFLAZIO", False, Txt(index)
            End If
        
            If index = PAG1_N_RP_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "COEFFRIP", True, Txt(index)
            End If

            If index = PAG0_TIPOTARIFFA_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PERCGETTITO", True, Txt(PAG0_TIPOTARIFFA_ID).Text
            End If
            
            If index = PAG1_PERCDETRASPAZZAMENTO_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PERCDETRAZIONE", True, Txt(PAG1_PERCDETRASPAZZAMENTO_ID).Text
            End If
        Case Else
        
            Txt(index) = Format$(CDbl(Txt(index)), FORMATVALUTA)
    End Select
    
    
    'Calcolo
    If Txt(index).BackColor <> BACKCOLOR_CAMPI_CALC Then
        Txt(PAG2_N_PERCINIUD_ID).Enabled = False
        Txt(PAG2_N_PASSO_ID).Enabled = False

        FaiICalcoli index
        
        Txt(PAG2_N_PERCINIUD_ID).Enabled = True
        Txt(PAG2_N_PASSO_ID).Enabled = True
    
    End If
    
    If index = PAG1_N_CUD_ID Or index = PAG1_N_CUA_ID Then
        LetParamValue mAnnoRiferimento, mCodComune, "PERCCUD", True, Txt(PAG1_N_CUD_ID).Text
        Txt(PAG1_N_CUD_ID) = Format$(CDbl(Txt(PAG1_N_CUD_ID)), FORMAT6DEC)
    
        LetParamValue mAnnoRiferimento, mCodComune, "PERCCUA", True, Txt(PAG1_N_CUA_ID).Text
        Txt(PAG1_N_CUA_ID) = Format$(CDbl(Txt(PAG1_N_CUA_ID)), FORMAT6DEC)
    End If
    
End Sub



Private Sub txt_Validate(index As Integer, Cancel As Boolean)

    'Il valore deve essere numerico
    If Not IsNumeric(Txt(index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: Txt(index) = OldTxtValue: Exit Sub
    
    Select Case index
         
         'Il valore deve essere una Percentuale
         Case PAG2_N_PASSO_ID
                
            If Txt(index) > 100 Or Txt(index) <= 0 Then
                MsgBox "Attenzione il valore deve essere una percentuale", vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
                
         'Il valore deve essere una Percentuale tra 1 - 99
         Case PAG2_N_PASSO_ID, PAG2_N_PERCINIUD_ID, PAG2_N_PERCINIUA_ID
                
            If Txt(index) > 99 Or Txt(index) <= 0 Then
                MsgBox "Attenzione il valore deve essere una percentuale", vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
                
                 
         Case PAG1_T_IP_ID, PAG1_T_X_ID, _
              PAG1_T_Y_ID, PAG1_T_CCON_ID, PAG1_T_K_ID, _
              PAG1_N_IP_ID, PAG1_N_RP_ID, PAG1_PERCDETRASPAZZAMENTO_ID
              
            If Txt(index) > 100 Or Txt(index) < 0 Then
                MsgBox "Attenzione il valore deve essere una percentuale", vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            
            If index = PAG1_T_X_ID And CDbl(Txt(PAG1_T_X_ID)) < mMINTRASLMIN Then
                MsgBox "Attenzione il valore deve essere almeno " & mMINTRASLMIN, vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            If index = PAG1_T_Y_ID And CDbl(Txt(PAG1_T_Y_ID)) < mMINRIDIMB23 Then
                MsgBox "Attenzione il valore deve essere almeno " & mMINRIDIMB23, vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            If index = PAG1_T_CCON_ID And CDbl(Txt(PAG1_T_CCON_ID)) < mMINRIDIMB1 Then
                MsgBox "Attenzione il valore deve essere almeno " & mMINRIDIMB1, vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            If index = PAG1_PERCDETRASPAZZAMENTO_ID And (CDbl(Txt(PAG1_PERCDETRASPAZZAMENTO_ID)) > mMAXPERCDETRAZIONE Or CDbl(Txt(PAG1_PERCDETRASPAZZAMENTO_ID)) < mMINPERCDETRAZIONE) Then
                MsgBox "Attenzione il valore deve essere compreso fra " & mMINPERCDETRAZIONE & " e " & mMAXPERCDETRAZIONE, vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            
            
    End Select
End Sub

Public Function getValueTxt(index As Integer) As String
    getValueTxt = Txt(index)
End Function

Public Sub allineaCampoCollegato(index As Integer, Valore As String)
    Dim pPagina As Integer
    
    pPagina = mPagina
    
    Txt(index) = Valore
    Select Case index
        Case PAG1_PERCDETRASPAZZAMENTO_ID
            mPagina = 1
            FaiICalcoli PAG1_PERCDETRASPAZZAMENTO_ID
        Case PAG1_N_IP_ID
            mPagina = 1
            FaiICalcoli PAG1_N_IP_ID
        Case PAG1_N_RP_ID
            mPagina = 1
            FaiICalcoli PAG1_N_RP_ID
        Case PAG2_N_PERCINIUD_ID
            mPagina = 2
            FaiICalcoli PAG2_N_PERCINIUD_ID
        Case PAG1_N_CUD_ID
            mPagina = 1
            FaiICalcoli PAG1_N_CUD_ID
        Case PAG1_N_CUA_ID
            mPagina = 1
            FaiICalcoli PAG1_N_CUA_ID
    End Select
    
    mPagina = pPagina
End Sub



Private Function FaiICalcoli(ByVal ID As Integer) As Boolean
    'fai calcoli della pagina corrente
    Dim ETF As Double
    Dim ETV As Double
        
    
    FaiICalcoli = True
    Select Case mPicName
        Case "Picture1"
            Select Case mPagina
                Case PAG0_CALCOLOMETODO
                    'MetodoTariffazione
            End Select
        Case "Picture2"
            Select Case mPagina
                Case PAG1_CALCOLOET1
                    Txt(PAG1_T_ET1_ID) = Format$(CalcoloET1Transitorio(CDbl(Txt(PAG1_T_ET0_ID)), CDbl(Txt(PAG1_T_IP_ID)), CDbl(Txt(PAG1_T_X_ID)), CDbl(Txt(PAG1_T_Y_ID)), CDbl(Txt(PAG1_T_CCON_ID)), CDbl(Txt(PAG1_T_K_ID))), FORMATVALUTA)
                Case PAG2_DISTRIBUZIONEUTENZE
                    Txt(PAG2_T_ET1D_ID) = Format$(CDbl(Txt(PAG2_T_UDOMESTICA_ID) / CDbl(Txt(PAG2_T_GETTITO_ID)) * CDbl(Txt(PAG2_T_ET1_ID))), FORMATVALUTA)
                    Txt(PAG2_T_ET1A_ID) = Format$(CDbl(Txt(PAG2_T_UATTIVITA_ID) / CDbl(Txt(PAG2_T_GETTITO_ID)) * CDbl(Txt(PAG2_T_ET1_ID))), FORMATVALUTA)
                    'Suddividi le ET1 domestiche da 20 -> 70
                    FaiICalcoli = SuddividiLeUtenze
            End Select
        Case "Picture3"
            
            Select Case mPagina
                Case PAG1_CALCOLOINTEGRALE
                    ETF = CDbl(Txt(PAG1_N_PF_ID))
                    ETV = CDbl(Txt(PAG1_N_PV_ID))
                    'Calcolo ET
                    Txt(PAG1_N_ET_ID) = Format$(CalcoloET1Integrale(CDbl(Txt(PAG1_N_IP_ID)), CDbl(Txt(PAG1_N_RP_ID)), CDbl(Txt(PAG1_N_CG_ID)), CDbl(Txt(PAG1_N_CC_ID)), CDbl(Txt(PAG1_N_CK_ID)), CDbl(Txt(PAG1_N_CT_ID)), ETF, ETV), FORMATVALUTA)

                    Txt(PAG1_N_ETF_ID) = Format$(ETF, FORMATVALUTA)
                    Txt(PAG1_N_ETV_ID) = Format$(ETV, FORMATVALUTA)
                    
                    
                    
                    '
                    Select Case ID
                        Case PAG1_N_CUD_ID
                            Txt(PAG1_N_CUA_ID) = 100 - CDbl(Txt(PAG1_N_CUD_ID))
                        Case PAG1_N_CUA_ID
                            Txt(PAG1_N_CUD_ID) = 100 - CDbl(Txt(PAG1_N_CUA_ID))
                        Case PAG1_N_RP_ID
                            If Txt(PAG1_N_RP_ID).Tag <> Txt(PAG1_N_RP_ID) Then
                                Txt(PAG1_N_CUA_ID) = 0
                                Txt(PAG1_N_RP_ID).Tag = Txt(PAG1_N_RP_ID)
                            End If
                        Case PAG1_N_IP_ID
                            If Txt(PAG1_N_IP_ID).Tag <> Txt(PAG1_N_IP_ID) Then
                                Txt(PAG1_N_CUA_ID) = 0
                                Txt(PAG1_N_IP_ID).Tag = Txt(PAG1_N_IP_ID)
                            End If
                    
                    End Select


                    CUAPerc = CDbl(Txt(PAG1_N_CUA_ID))
                    CUDPerc = CDbl(Txt(PAG1_N_CUD_ID))
                    
                    If mOptMetodoTariffa = tpoTariffa.TariffaNormalizzata Or _
                        mOptMetodoTariffa = TariffaCostiIndivisibili Then
                        FaiICalcoli = CacolaCUDomesticoEAttivita(CUAPerc, CUDPerc)
                    Else
                        'azzero il CU nel caso di porta a porta
                        CUAPerc = 0
                        CUDPerc = 0
                    End If
                    
                    CUASave = CUA
                    Txt(PAG1_N_CUD_ID) = Format$(CUDPerc, FORMAT6DEC)
                    Txt(PAG1_N_CUA_ID) = Format$(CUAPerc, FORMAT6DEC)
                    lbl(PAG1_N_LBL_CUD) = Format$(CUD, FORMAT6DEC)
                    lbl(PAG1_N_LBL_CUA) = Format$(CUA, FORMAT6DEC)
                    
                    Txt(PAG1_N_CUD_ID).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    Txt(PAG1_N_CUA_ID).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(PAG1_N_LBL_CUD).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(PAG1_N_LBL_CUA).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(45).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(46).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(47).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    lbl(48).Visible = mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili
                    
                    
                Case PAG2_SUDDIVISIONEPARTEFISSA
                    Txt(PAG2_N_PERCINIUA_ID) = 100 - CDbl(Txt(PAG2_N_PERCINIUD_ID))
                    FaiICalcoli = SuddividiParteFissa
                    
                Case PAG3_TARIFFAUDNORMALIZZATE
                    
                    'ANTONIO 26-02-02mRiduzione = False
                    'ANTONIO 26-02-02TariffeUtenzeDomesticheNormalizzato mRiduzione
                    FaiICalcoli = TariffeUtenzeDomesticheNormalizzato
                    'ANTONIO 26-02-02Command1(MOSTRARIDUZIONI).Visible = Not rsRIDUA.EOF Or Not rsRIDUD.EOF
                Case PAG4_TARIFFAUDNORMALIZZATE_END
                    'ANTONIO 26-02-02PaginaRiassuntivaTariffaUD mRiduzione
                    FaiICalcoli = PaginaRiassuntivaTariffaUD
                Case PAG5_TARIFFAUANORMALIZZATE
                    'ANTONIO 26-02-02CUA = CUASave
                    'ANTONIO 26-02-02TariffeUtenzeAttivitaNormalizzato mRiduzione
                    FaiICalcoli = TariffeUtenzeAttivitaNormalizzato
                Case PAG6_TARIFFAUANORMALIZZATE_END
                    'ANTONIO 26-02-02PaginaRiassuntivaTariffaUA mRiduzione
                    FaiICalcoli = PaginaRiassuntivaTariffaUA
                Case PAG7_CORREZIONETARIFFE
                    FaiICalcoli = CacoloTipoAdeguamento(cmb.ListIndex, vbNullString)
                    'ANTONIO 26-02-02If mRiduzione Then Command1_Click (MOSTRARIDUZIONI)
            End Select
    
        Case "Picture4"
            Select Case mPagina
                Case PAG1_CALCOLO507
                
                    Dim TotUtenzeInRiduzioniDAUD_PV As Double
                    Dim TotSuperficieInRiduzioniDAUD_PF As Double
                    Dim TotSuperficieInRiduzioniDAUA_Pv As Double
                    Dim TotSuperficieInRiduzioniDAUA_PF As Double
                    
                    CostoTotAlNettoSpazzamento = CDbl(Txt(PAG0_TOTSERVIZIO_ID)) * (1 - (CDbl(Txt(PAG1_PERCDETRASPAZZAMENTO_ID)) / 100))
                    CalcoloCostoCopertura = CostoTotAlNettoSpazzamento * (CDbl(Txt(PAG0_TIPOTARIFFA_ID).Text) / 100)
                    Txt(PAG1_COSTOCOPERTURA_ID).Text = Format$(CalcoloCostoCopertura, FORMATVALUTA)
                    Txt(PAG1_COSTOTOTALNETTO_ID).Text = Format$(CostoTotAlNettoSpazzamento, FORMATVALUTA)
                    CMG = CalcoloCostoCopertura / (rsDatiTecniciComune("TotSuperficieIscrRuolo"))
                    Txt(PAG1_CMG_ID).Text = Format$(CMG, FORMAT6DEC)
                    TotaliDaRiduzioniUA "", rsRIDUA, , TotSuperficieInRiduzioniDAUA_PF
                    TotaliDaRiduzioniUD "", rsRIDUD, , TotSuperficieInRiduzioniDAUD_PF
                    SuperficiCorrette = TotSuperficieInRiduzioniDAUA_PF + TotSuperficieInRiduzioniDAUD_PF
                    CMG = CalcoloCostoCopertura / (rsDatiTecniciComune("TotSuperficieIscrRuolo") - SuperficiCorrette)
                    Txt(PAG1_CMGCONRIDUZIONI_ID).Text = Format$(CMG, FORMAT6DEC)
                    Txt(PAG1_COSTOTOTRIFIUTI_ID).Text = Format$(CDbl(Txt(PAG0_TOTSERVIZIO_ID)) / rsDatiTecniciComune("QRifiutiTotali"), FORMAT6DEC)
                    KM = rsDatiTecniciComune("QRifiutiTotali") / (rsDatiTecniciComune("TotSuperficieIscrRuolo") - SuperficiCorrette)
                    Txt(PAG1_KM_ID).Text = Format$(KM, "###,###,###,##0.00")
                    
                Case PAG2_CALCOLO507RISULTATO
                    FaiICalcoli = Risultato507NEW(vbNullString)
            End Select
    End Select
     
End Function



Private Function CaricaValori() As Boolean
    'Carica i valori nella pagina corrente
    
    On Error Resume Next
    CaricaValori = True
    Select Case mPicName
        Case "Picture1"
            Select Case mPagina
                Case PAG0_CALCOLOMETODO
                    
                    '
                    'Imposto il Gettito e Il totale servizio
                    Txt(PAG0_TOTSERVIZIO_ID) = Format$(0, FORMATVALUTA)
                    Txt(PAG0_GETTITO_ID) = Format$(0, FORMATVALUTA)
                    
                    Txt(PAG0_TOTSERVIZIO_ID) = Format$(CDbl(rsCostiComune("Tot_Servizio")), FORMATVALUTA)
                    Txt(PAG0_GETTITO_ID) = Format$(CDbl(rsCostiComune("Gettito")), FORMATVALUTA)
                    Txt(PAG0_GETTITO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG0_TOTSERVIZIO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    
                    'Calcolo Metodo
                    MetodoTariffazione
                    
            End Select
        
        Case "Picture2"
        
            Select Case mPagina
        
                Case PAG1_CALCOLOET1
                    
                    '
                    Txt(PAG1_T_ET0_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_T_ET1_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    '
                    Txt(PAG1_T_ET0_ID) = Txt(PAG0_GETTITO_ID)
                    Txt(PAG1_T_IP_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "INFLAZIO")
                    Txt(PAG1_T_X_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "TRASLMIN", , mMINTRASLMIN)
                    lbl(MINTRASLMIN_ID) = "(MIN " & Format$(mMINTRASLMIN, "##0.0") & ")"
                    Txt(PAG1_T_Y_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "RIDIMB23", , mMINRIDIMB23)
                    lbl(MINRIDIMB23_ID) = "(MIN " & Format$(mMINRIDIMB23, "##0.0") & ")"
                    Txt(PAG1_T_CCON_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "RIDIMB1", , mMINRIDIMB1)
                    lbl(MINRIDIMB1_ID) = "(MIN " & Format$(mMINRIDIMB1, "##0.0") & ")"
                    Txt(PAG1_T_K_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFINC")
                    
                Case PAG2_DISTRIBUZIONEUTENZE
                    
                    Txt(PAG2_T_GETTITO_ID) = Txt(PAG0_GETTITO_ID)
                    Set rsDatiTecniciComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
                    Set rsDatiAnagraficiComune = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
                    
                    Txt(PAG2_T_UDOMESTICA_ID) = Format$(CDbl(rsDatiTecniciComune("TotGettitoDaUtenzaDomestica")), FORMATVALUTA)
                    Txt(PAG2_T_UATTIVITA_ID) = Format$(CDbl(rsDatiTecniciComune("TotGettitoDaUtenzaAttivita")), FORMATVALUTA)
                    
                    Txt(PAG2_T_ET1_ID) = Txt(PAG1_T_ET1_ID)
                    Txt(PAG2_T_ET1D_ID) = 0
                    Txt(PAG2_T_ET1A_ID) = 0
        
                    Txt(PAG2_T_UDOMESTICA_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_UDOMESTICA_ID).Locked = True
                    Txt(PAG2_T_UATTIVITA_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_UATTIVITA_ID).Locked = True
                    Txt(PAG2_T_GETTITO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_GETTITO_ID).Locked = True
                    Txt(PAG2_T_ET1_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_ET1_ID).Locked = True
                    Txt(PAG2_T_ET1D_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_ET1D_ID).Locked = True
                    Txt(PAG2_T_ET1A_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_T_ET1A_ID).Locked = True
                    
            
                Case PAG3_TARIFFACATEGORIE
                    
                    CaricaValori = TariffeCategorieAttivita
                Case PAG4_TARIFFAUTENZEDOMESTICHE
                    
                    CaricaValori = TariffeUtenzeDomestiche
            End Select
        
        
        Case "Picture3"
            Select Case mPagina
                
                Case PAG1_CALCOLOINTEGRALE
                
                    
                    Txt(PAG1_N_ET_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_ET_ID).Locked = True
                    Txt(PAG1_N_ETF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_ETF_ID).Locked = True
                    Txt(PAG1_N_ETV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_ETV_ID).Locked = True
                    Txt(PAG1_N_CK_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_CK_ID).Locked = True
                    Txt(PAG1_N_CC_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_CC_ID).Locked = True
                    Txt(PAG1_N_CG_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_CG_ID).Locked = True
                    Txt(PAG1_N_CT_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_CT_ID).Locked = True
                    Txt(PAG1_N_PF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_PF_ID).Locked = True
                    Txt(PAG1_N_PV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_N_PV_ID).Locked = True

                    lbl(PAG1_N_LBL_ETF).ForeColor = FORECOLOR_PARTE_FISSA
                    lbl(PAG1_N_LBL_ETV).ForeColor = FORECOLOR_PARTE_VARIABILE
                    lbl(PAG1_N_LBL_PF).ForeColor = FORECOLOR_PARTE_FISSA
                    lbl(PAG1_N_LBL_PV).ForeColor = FORECOLOR_PARTE_VARIABILE
                    
                    
                    'Carica valore da Parametri
                    Txt(PAG1_N_IP_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "INFLAZIO")
                    Txt(PAG1_N_IP_ID).Tag = Txt(PAG1_N_IP_ID)
                    Txt(PAG1_N_RP_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFRIP")
                    Txt(PAG1_N_RP_ID).Tag = Txt(PAG1_N_RP_ID)
                    
                    'Carica valori da DB
                    Txt(PAG1_N_CK_ID) = Format$(CDbl(rsCostiComune("CK")), FORMATVALUTA)
                    Txt(PAG1_N_CC_ID) = Format$(CDbl(rsCostiComune("CC")), FORMATVALUTA)
                    Txt(PAG1_N_CG_ID) = Format$(CDbl(rsCostiComune("CG")), FORMATVALUTA)
                    Txt(PAG1_N_CT_ID) = Format$(CDbl(rsCostiComune("Tot_Servizio")), FORMATVALUTA)
                    Txt(PAG1_N_PF_ID) = Format$(CDbl(rsCostiComune("ParteFissa")), FORMATVALUTA)
                    Txt(PAG1_N_PV_ID) = Format$(CDbl(rsCostiComune("ParteVariabile")), FORMATVALUTA)
                    Txt(PAG1_N_CUA_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCCUA")
                    Txt(PAG1_N_CUD_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCCUD")
                    
                Case PAG2_SUDDIVISIONEPARTEFISSA
            
                    'Colonna selezionata per il calcolo tariffa
                    keyColUDA = GetParamValue(mAnnoRiferimento, mCodComune, "COLUDA")
                    
                    Txt(PAG2_N_ET_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_N_ET_ID).Locked = True
                    Txt(PAG2_N_ET_ID) = Txt(PAG1_N_ET_ID)
                    Txt(PAG2_N_ETF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_N_ETF_ID).Locked = True
                    Txt(PAG2_N_ETF_ID) = Txt(PAG1_N_ETF_ID)
                    Txt(PAG2_N_ETV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_N_ETV_ID).Locked = True
                    Txt(PAG2_N_ETV_ID) = Txt(PAG1_N_ETV_ID)
                    Txt(PAG2_N_PERCINIUA_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG2_N_PERCINIUA_ID).Locked = True
                    
                    
                    lbl(PAG2_N_LBL_ETF).ForeColor = FORECOLOR_PARTE_FISSA
                    lbl(PAG2_N_LBL_ETV).ForeColor = FORECOLOR_PARTE_VARIABILE
            
                    Txt(PAG2_N_PASSO_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PASSOPF")
                    Txt(PAG2_N_PERCINIUD_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCINIUD")
                    Txt(PAG2_N_PERCINIUA_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCINIUA")
                    Command1(MOSTRARIDUZIONI).Visible = False
                    Command1(MOSTRAINVALUTA).Visible = False

                Case PAG3_TARIFFAUDNORMALIZZATE
                Case PAG4_TARIFFAUDNORMALIZZATE_END
                
                    Command1(MOSTRARIDUZIONI).Visible = False
                    Command1(MOSTRAINVALUTA).Visible = False
    
                    Txt(PAG4_QUF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG4_QUF_ID).Locked = True
                    Txt(PAG4_QUV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG4_QUV_ID).Locked = True
                    Txt(PAG4_TotQF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG4_TotQF_ID).Locked = True
                    Txt(PAG4_TotQV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG4_TotQV_ID).Locked = True
                    Txt(PAG4_CU_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG4_CU_ID).Locked = True
                    TotQFUD = CDbl(lvw4.ListItems("DOMESTICA").SubItems(idxColonnaSelezionata(keyColUDA) + 1))
                    Txt(PAG4_TotQF_ID) = Format$(TotQFUD, FORMATVALUTA)
                    Txt(PAG4_TotQV_ID) = Format$(ToTQVUD, FORMATVALUTA)
                    Txt(PAG4_QUF_ID) = Format$(QUFUD, FORMAT6DEC)
                    Txt(PAG4_QUV_ID) = Format$(QUV, FORMAT6DEC)
                    Txt(PAG4_CU_ID) = Format$(CUD, FORMAT6DEC)

                Case PAG5_TARIFFAUANORMALIZZATE
                
                Case PAG6_TARIFFAUANORMALIZZATE_END

                    QUFUD = QUFUDsav
                    CUD = CUDsav
                    CUA = CUAsav
                    QUAPUA = QUAPUAsav


                    Txt(PAG6_QUAP_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG6_QUAP_ID).Locked = True
                    Txt(PAG6_TOTQF_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG6_TOTQF_ID).Locked = True
                    Txt(PAG6_TOTQV_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG6_TOTQV_ID).Locked = True
                    Txt(PAG6_CU_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG6_CU_ID).Locked = True

                    ToTQFUA = CDbl(lvw4.ListItems("ATTIVITA").SubItems(idxColonnaSelezionata(keyColUDA) + 1))
                    Txt(PAG6_TOTQF_ID) = Format$(ToTQFUA, FORMATVALUTA)
                    Txt(PAG6_TOTQV_ID) = Format$(ToTQVUA, FORMATVALUTA)
                    Txt(PAG6_QUAP_ID) = Format$(QUAPUA, "###,###,###,##0.00000")
                    Txt(PAG6_CU_ID) = Format$(CUA, "###,###,###,##0.00000")
                    Command1(MOSTRARIDUZIONI).Visible = False
                    Command1(MOSTRAINVALUTA).Visible = False
                    Command1(MOSTRAINVALUTA).Caption = "In Lire"
                    Command1(MOSTRARIDUZIONI).Caption = "Applica &Riduzioni"
                    
                Case PAG7_CORREZIONETARIFFE
                    Txt(PAG7_GETTITOUD_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG7_GETTITOUD_ID).Locked = True
                    Txt(PAG7_GETTITOUA_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG7_GETTITOUA_ID).Locked = True
                    Txt(PAG7_GETTITODICHIARATO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG7_GETTITODICHIARATO_ID).Locked = True
                    Txt(PAG7_FATTORECORRETTIVO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG7_FATTORECORRETTIVO_ID).Locked = True
                    Txt(PAG7_GUADAGNO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG7_GUADAGNO_ID).Locked = True

                    Txt(PAG7_GETTITOUD_ID) = Format$(GettitoCalcUDPF + GettitoCalcUDPV, FORMATVALUTA)
                    Txt(PAG7_GETTITOUA_ID) = Format$(GettitoCalcUAPF + GettitoCalcUAPV, FORMATVALUTA)
                    Txt(PAG7_GETTITODICHIARATO_ID) = Txt(PAG1_N_ET_ID)
                    Txt(PAG7_GUADAGNO_ID) = Format$((GettitoCalcUDPF + GettitoCalcUDPV + GettitoCalcUAPF + GettitoCalcUAPV) / CDbl(Txt(PAG1_N_ET_ID)), "##0.00")
                    
                    If CDbl(Txt(PAG7_GUADAGNO_ID)) = 1 Then
                        LetParamValue mAnnoRiferimento, mCodComune, "TIPOCORREZIONE", True, 0
                        cmb.Enabled = False
                    Else
                        cmb.Enabled = True
                    End If
                    
                    cmb.ListIndex = GetParamValue(mAnnoRiferimento, mCodComune, "TIPOCORREZIONE")
                    
                    
                    Command1(MOSTRARIDUZIONI).Visible = rsRIDUA.EOF = False Or rsRIDUD.EOF = False
                    
                    
                    Command1(MOSTRAINVALUTA).Visible = True
                    Command1(MOSTRARIDUZIONI).Enabled = True
                    
                    Command1(MOSTRARIDUZIONI).Default = Command1(MOSTRARIDUZIONI).Visible
                    Command1(MOSTRARIDUZIONI).SetFocus
                    Command1(AVANTI).Default = Not Command1(MOSTRARIDUZIONI).Default

            End Select
        
        Case "Picture4"
            
            
            Select Case mPagina
                Case PAG1_CALCOLO507
                    Txt(PAG1_PERCDETRASPAZZAMENTO_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCDETRAZIONE", mMAXPERCDETRAZIONE, mMINPERCDETRAZIONE)
                    lbl(MINMAXPERCDETRAZIONE_ID) = "(" & Format$(mMINPERCDETRAZIONE, "##0.0") & " -  " & Format$(mMAXPERCDETRAZIONE, "##0.0") & " )"
            
                    Txt(PAG1_COSTOTOTALNETTO_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_COSTOTOTALNETTO_ID).Locked = True
                    Txt(PAG1_CMG_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_CMG_ID).Locked = True
                    Txt(PAG1_CMGCONRIDUZIONI_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_CMGCONRIDUZIONI_ID).Locked = True
                    Txt(PAG1_COSTOTOTRIFIUTI_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_COSTOTOTRIFIUTI_ID).Locked = True
                    Txt(PAG1_COSTOCOPERTURA_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_COSTOCOPERTURA_ID).Locked = True
                    Txt(PAG1_KM_ID).BackColor = BACKCOLOR_CAMPI_CALC
                    Txt(PAG1_KM_ID).Locked = True
                    
                    Command1(MOSTRARIDUZIONI).Visible = False
                    Command1(MOSTRAINVALUTA).Visible = False
                    Command1(MOSTRAINVALUTA).Caption = "In Lire"
                    Command1(MOSTRARIDUZIONI).Caption = "Mostra &Riduzioni"

                    
                Case PAG2_CALCOLO507RISULTATO
                    Command1(MOSTRARIDUZIONI).Caption = "Applica &Riduzioni"
                    Command1(MOSTRARIDUZIONI).Visible = SuperficiCorrette > 0
                    
                    Command1(MOSTRARIDUZIONI).Enabled = True
                    Command1(MOSTRAINVALUTA).Visible = True
            End Select
    
    End Select
    
    
End Function


Private Function SuddividiLeUtenze() As Boolean
    Dim i As Long
    Dim ii As Long
    Dim Perc As Long
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim inc As Integer
    
    
    SuddividiLeUtenze = False
    With lvw
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        If CDbl(Txt(PAG2_T_ET1D_ID)) = 0 Then Exit Function
        
        .ColumnHeaders.Add , , "% di Ripartizione"
        For i = 0 To 50 \ 5 + 2
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(20 + 5 * i) & " - %" & CStr(100 - (20 + 5 * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDomestica And ColItem.Key = keyColUAttivita Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            ElseIf ColItem.Key = keyColUDomestica Then
                ColItem.Icon = ICON_DOMESTICA
            ElseIf ColItem.Key = keyColUAttivita Then
                ColItem.Icon = ICON_ATTIVITA
            End If
        Next
        
        .ListItems.Clear
        'Utenze Domestiche
        Set Item = .ListItems.Add(, , "DOMESTICHE")
        For ii = 1 To 2
            If ii = 1 Then
                Set Item = .ListItems.Add(, "DOMESTICHENCOMPFAMIGLIA", "N.Comp.Famiglia")
                Perc = 20
                inc = 1
            ElseIf ii = 2 Then
                Set Item = .ListItems.Add(, "DOMESTICHESUPERFICIE", "Superficie")
                Perc = 80
                inc = -1
            End If
            
            For i = 0 To 50 \ 5 + 2
                Item.SubItems(i + 1) = Format$(CDbl(Txt(PAG2_T_ET1D_ID)) * ((Perc + inc * 5 * i) / 100), FORMATVALUTA)
            Next
        Next
        
        'Utenze Attivita
        Set Item = .ListItems.Add(, , "")
        Set Item = .ListItems.Add(, , "ATTIVITA'")
        For ii = 1 To 2
            If ii = 1 Then
                Set Item = .ListItems.Add(, "ATTIVITAPRODUTTIVITA", "Produttività")
                Perc = 20
                inc = 1
            ElseIf ii = 2 Then
                Set Item = .ListItems.Add(, "ATTIVITASUPERFICIE", "Superficie")
                Perc = 80
                inc = -1
            End If
            
            For i = 0 To 50 \ 5 + 2
                Item.SubItems(i + 1) = Format$(CDbl(Txt(PAG2_T_ET1A_ID)) * ((Perc + inc * 5 * i) / 100), FORMATVALUTA)
            Next
        Next
        
        AllargaLeColonneLVW lvw, True, 3
        
        .Visible = True
    End With
    SuddividiLeUtenze = True
End Function


Private Function TariffeCategorieAttivita() As Boolean
    'Calcolo Tariffa  Utenze attivita con il metodo transitorio
    
    Dim i As Long
    Dim ii As Long
    Dim iTS As Long
    Dim iTA As Long
    Dim iALFA As Long
    Dim iBETA As Long
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim qry As ADODB.Command
    Dim rs As Recordset
    Dim TariffaMedia As Double
    Dim GettitoVoluto As Double
    Dim GettitoVolutoSuperficie As Double
    Dim TariffaSperimentale() As Double
    Dim GettitoTotaleSperimentale As Double
    Dim GettitoTotaleSperimentaleSuperficie As Double
    Dim TariffaAttuale() As Double
    Dim Alfa() As Double
    Dim Beta() As Double
    Dim Tariffa As Double
    
    TariffeCategorieAttivita = False
    With lvw2
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        If CDbl(Txt(PAG2_T_ET1D_ID)) = 0 Then Exit Function
        
        .ColumnHeaders.Add , , "Categorie Attività"
        For i = 0 To 50 \ 5 + 2
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(20 + 5 * i) & " - %" & CStr(100 - (20 + 5 * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDomestica And ColItem.Key = keyColUAttivita Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            ElseIf ColItem.Key = keyColUAttivita Then
                ColItem.Icon = ICON_ATTIVITA
            End If
        Next
        
        .ListItems.Clear
            
        'KC
        Set rs = SelezionaCoeffK(CoeffKC, mCodComune, mAnnoRiferimento, rsDatiAnagraficiComune("AreaGeografica"), mParAbitanti)
        For i = 1 To lvw.ColumnHeaders.Count - 1
            rs.MoveFirst
            
            iTS = -1
            On Error Resume Next
            iTS = UBound(TariffaSperimentale, 2)
            On Error GoTo 0
            ReDim Preserve TariffaSperimentale(100, iTS + 1)
            
            GettitoVoluto = CDbl(lvw.ListItems("ATTIVITAPRODUTTIVITA").SubItems(i))
            TariffaMedia = GettitoVoluto / rsDatiTecniciComune("TotSuperficieUtenzaAttivita")
            GettitoTotaleSperimentale = 0
            
            iTA = -1
            On Error Resume Next
            iTA = UBound(TariffaAttuale, 2)
            On Error GoTo 0
            ReDim Preserve TariffaAttuale(100, iTA + 1)
            
            
            GettitoVolutoSuperficie = CDbl(lvw.ListItems("ATTIVITASUPERFICIE").SubItems(i))
            GettitoTotaleSperimentaleSuperficie = 0

            ii = 0
            
            Do While Not rs.EOF
                'Produttivita
                TariffaSperimentale(ii, iTS + 1) = TariffaMedia * rs("Valore")
                GettitoTotaleSperimentale = GettitoTotaleSperimentale + (TariffaSperimentale(ii, iTS + 1) * rs("Superficie"))
                
                'Superficie
                TariffaAttuale(ii, iTA + 1) = rs("TariffaAttuale")
                GettitoTotaleSperimentaleSuperficie = GettitoTotaleSperimentaleSuperficie + (TariffaAttuale(ii, iTA + 1) * rs("Superficie"))
                
                ii = ii + 1
                rs.MoveNext
            Loop
            
            iALFA = -1
            On Error Resume Next
            iALFA = UBound(Alfa, 1)
            On Error GoTo 0
            ReDim Preserve Alfa(iALFA + 1)
            Alfa(iALFA + 1) = GettitoTotaleSperimentale / GettitoVoluto
        
            iBETA = -1
            On Error Resume Next
            iBETA = UBound(Beta, 1)
            On Error GoTo 0
            ReDim Preserve Beta(iBETA + 1)
            Beta(iBETA + 1) = GettitoTotaleSperimentaleSuperficie / GettitoVolutoSuperficie
            
        Next
        
        
        'Utenze Attivita
        ii = 0
        rs.MoveFirst
        Do While Not rs.EOF
            Set Item = .ListItems.Add(, "%" & rs("CodTipologiaAttivita"), rs("Descrizione"))
            For i = 1 To lvw.ColumnHeaders.Count - 1
                Tariffa = (TariffaSperimentale(ii, i - 1) / Alfa(i - 1)) + (TariffaAttuale(ii, i - 1) / Beta(i - 1))
                Item.SubItems(i) = Format$(Tariffa, FORMATVALUTA)
            Next
            ii = ii + 1
            rs.MoveNext
        Loop
        
        AllargaLeColonneLVW lvw2, 3, 2
        
        .Visible = True
    
    
    End With
    TariffeCategorieAttivita = True
    

End Function

Private Function TariffeUtenzeDomestiche() As Boolean
    'Calcolo Tariffa Utenze Domestica
    Dim i As Long
    Dim ii As Long
    Dim iTS As Long
    Dim iTA As Long
    Dim iALFA As Long
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim qry As ADODB.Command
    Dim rs As Recordset
    
    
    Dim TariffaMedia As Double
    Dim GettitoVoluto As Double
    Dim GettitoVolutoSuperficie As Double
    Dim TariffaSperimentale() As Double
    Dim TariffaSuperficie() As Double
    Dim GettitoTotaleSperimentale As Double
    Dim Alfa() As Double
    Dim NComFamiglia As Long
    Dim Tariffa   As Double
    
    TariffeUtenzeDomestiche = False
    With lvw3
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        If CDbl(Txt(PAG2_T_ET1D_ID)) = 0 Then Exit Function
        
        .ColumnHeaders.Add , , "Utenza Domestica"
        For i = 0 To 50 \ 5 + 2
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(20 + 5 * i) & " - %" & CStr(100 - (20 + 5 * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDomestica And ColItem.Key = keyColUAttivita Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            ElseIf ColItem.Key = keyColUDomestica Then
                ColItem.Icon = ICON_DOMESTICA
            End If
        Next
        
        .ListItems.Clear
        
        'KB
        Set rs = SelezionaCoeffK(CoeffKB, mCodComune, mAnnoRiferimento, 0, 0)
        If rs.EOF Then Info "Non Trovati Coefficienti KB": Exit Function
        
        For i = 1 To lvw.ColumnHeaders.Count - 1
            rs.MoveFirst
            
            iTS = -1
            On Error Resume Next
            iTS = UBound(TariffaSperimentale, 2)
            On Error GoTo 0
            ReDim Preserve TariffaSperimentale(100, iTS + 1)
            
            iTA = -1
            On Error Resume Next
            iTA = UBound(TariffaSuperficie)
            On Error GoTo 0
            ReDim Preserve TariffaSuperficie(iTA + 1)
            
            GettitoVoluto = CDbl(lvw.ListItems("DOMESTICHENCOMPFAMIGLIA").SubItems(i))
            TariffaMedia = GettitoVoluto / (rsDatiTecniciComune("QRifiutiTotali") - rsDatiTecniciComune("QRifiutiCategorieAttivita"))
            GettitoTotaleSperimentale = 0
            
            GettitoVolutoSuperficie = CDbl(lvw.ListItems("DOMESTICHESUPERFICIE").SubItems(i))
            
            'Tariffa Superficie
            TariffaSuperficie(iTA + 1) = GettitoVolutoSuperficie / rsDatiTecniciComune("TotSuperficieUtenzaDomestica")

            ii = 0
            
            Do While Not rs.EOF
                'Numero Componenti
                Select Case rs("IDNumeroComponenti")
                    Case 1
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componente")
                    Case Is >= 2
                        If rs("IDNumeroComponenti") = 6 Then
                            NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                        Else
                            NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("IDNumeroComponenti")) & "Componenti")
                        End If
                End Select
                
                'Aggiundo i numero ruoli seconda casa se vi sono
                If rsDatiTecniciComune("Id_NumCompSecondaCasa") = rs("IDNumeroComponenti") Then
                    NComFamiglia = NComFamiglia + rsDatiTecniciComune("NRuoliSecondaCasa")
                End If
                
                
                TariffaSperimentale(ii, iTS + 1) = TariffaMedia * rs("Valore") * 360
                GettitoTotaleSperimentale = GettitoTotaleSperimentale + (TariffaSperimentale(ii, iTS + 1) * NComFamiglia)
                
                ii = ii + 1
                rs.MoveNext
            Loop
            
            iALFA = -1
            On Error Resume Next
            iALFA = UBound(Alfa, 1)
            On Error GoTo 0
            ReDim Preserve Alfa(iALFA + 1)
            Alfa(iALFA + 1) = GettitoTotaleSperimentale / GettitoVoluto
        
        Next
        
        
        'Utenze Domestiche NumeroComponenti
        ii = 0
        rs.MoveFirst
        Do While Not rs.EOF
            Set Item = .ListItems.Add(, "%" & rs("IdNumeroComponenti"), rs("Descrizione"))

            For i = 1 To lvw.ColumnHeaders.Count - 1
                Tariffa = TariffaSperimentale(ii, i - 1) / Alfa(i - 1)
                Item.SubItems(i) = Format$(Tariffa, FORMATVALUTA)
            Next
            ii = ii + 1
            rs.MoveNext
        Loop
        
        'Utenze Superficie
        Set Item = .ListItems.Add(, "%UDSuperficie", "Superficie")
        For i = 1 To lvw.ColumnHeaders.Count - 1
            Item.SubItems(i) = Format$(TariffaSuperficie(i - 1), FORMATVALUTA)
        Next


        AllargaLeColonneLVW lvw3, True, 2
        
        .Visible = True
    End With
    TariffeUtenzeDomestiche = True


End Function


Private Function SuddividiParteFissa() As Boolean
    Dim i As Long
    Dim ii As Long
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim inc As Integer
    Dim Perc As Double
    Dim Estremo As Double
    Dim Passo As Double
    Dim PERCINIUD As Double
    Dim PercIniUA As Double
    Dim Valore As Double
    
    SuddividiParteFissa = False
    With lvw4
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        If CDbl(Txt(PAG2_N_ET_ID)) = 0 Then Exit Function
        
        .ColumnHeaders.Add , , "% di Ripartizione"
        
        Passo = Txt(PAG2_N_PASSO_ID)
        PERCINIUD = Txt(PAG2_N_PERCINIUD_ID)
        PercIniUA = Txt(PAG2_N_PERCINIUA_ID)
        
        If (PERCINIUD > PercIniUA) Then
            Estremo = (100 - Passo - PERCINIUD) / Passo
        Else
            Estremo = Abs((PercIniUA - PERCINIUD) / Passo)
        End If
        
        For i = 0 To Estremo
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(PERCINIUD + Passo * i) & " - %" & CStr(100 - (PERCINIUD + Passo * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Alignment = lvwColumnRight
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDA Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            End If
        Next
        
        .ListItems.Clear
        'Utenze Domestiche
        Set Item = .ListItems.Add(, , "PARTE FISSA")
        For ii = 1 To 2
            If ii = 1 Then
                Set Item = .ListItems.Add(, "DOMESTICA", "Domestica")
                Perc = PERCINIUD
                inc = 1
            ElseIf ii = 2 Then
                Set Item = .ListItems.Add(, "ATTIVITA", "Attivita")
                Perc = PercIniUA
                inc = -1
            End If
            
            For i = 0 To Estremo
                Valore = CDbl(Txt(PAG2_N_ETF_ID)) * ((Perc + inc * Passo * i) / 100)
                Item.SubItems(i + 1) = Format$(Valore, FORMATVALUTA)
            Next
        Next
        
        AllargaLeColonneLVW lvw4, True, 3
        
        
        .Visible = True

    End With
    
    SuddividiParteFissa = True
End Function




Private Function TariffeUtenzeDomesticheNormalizzato() As Boolean
    'Calcolo Tariffa Utenze Domestica con il metodo Normalizzato
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim qry As ADODB.Command
    Dim rs As Recordset
    Dim i As Integer, ii As Integer
    Dim Atot As Double
    Dim Passo As Double
    Dim PERCINIUD As Double
    Dim PercIniUA As Double
    Dim Estremo  As Double
    Dim Tariffa As Double
    Dim TotSuperficie As Double
    Dim QUF As Double
    
    
    With lvw5
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        TariffeUtenzeDomesticheNormalizzato = False
        Set ColItem = .ColumnHeaders.Add(, , "Utenza Domestica")
        Set ColItem = .ColumnHeaders.Add(, , "Parte Variabile")
        ColItem.Alignment = lvwColumnRight
        
        Passo = Txt(PAG2_N_PASSO_ID)
        PERCINIUD = Txt(PAG2_N_PERCINIUD_ID)
        PercIniUA = Txt(PAG2_N_PERCINIUA_ID)
        
        If (PERCINIUD > PercIniUA) Then
            Estremo = (100 - Passo - PERCINIUD) / Passo
        Else
            Estremo = Abs((PercIniUA - PERCINIUD) / Passo)
        End If
        
        For i = 0 To Estremo
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(PERCINIUD + Passo * i) & " - %" & CStr(100 - (PERCINIUD + Passo * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Alignment = lvwColumnRight
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDA Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            End If
        Next
        
        .ListItems.Clear
            
        'calcolo superfice con KA
        Erase SupCorrettaConRiduzioneUD
        Erase SupCorrettaSenzaRiduzioneUD
        
        If Not SuperficeCorrettaUD(mCodComune, mAnnoRiferimento, SupCorrettaConRiduzioneUD(), SupCorrettaSenzaRiduzioneUD(), rsRIDUD, True) Then
            Info "Attenzione Coefficienti KA mancanti"
            Exit Function
        End If
            
            
        'Parte variabile
        '--------------------------------------------------------------------------------
        'se il metodo tariffazione è il porta porta il calcolo della parte variabile non
        'è più in base ai coefficienti e il nuomero utenze, ma per svuotamente.
        If mOptMetodoTariffa = TariffaPortaAPorta Then
            
            Set rs = SelezionaTotSvuotamentiUD(mCodComune, mAnnoRiferimento)
            If rs.EOF Then Info "Non Trovati gli svuotamenti utenza domestica per il comune " & mCodComune & "-" & mDesComune: Exit Function
                        
            'Parte variabile
            ii = 0
            rs.MoveFirst
            Do While Not rs.EOF
                Set Item = .ListItems.Add(, "%" & rs("IdNumeriComponenti"), rs("Descrizione"))
                Tariffa = rsCostiComune("ParteVariabile").value / (rs("TotaleSvuotamentoPerComune").value * mCoeffMedioSvuotamentoLT)
                Item.SubItems(1) = Format$(Tariffa, FORMAT6DEC)
                Item.ListSubItems(1).Tag = SetValueInTag(TAG_TAR_PVUD_FLD, Tariffa, "")
                ii = ii + 1
                rs.MoveNext
            Loop
        Else
            'KB
            Set rs = SelezionaCoeffK(CoeffKB, mCodComune, mAnnoRiferimento, 0, 0)
            If rs.EOF Then Info "Non Trovati Coefficienti KB": Exit Function
            
            '
            Atot = NumeroUtenzeDomesticheCorreteKB(mCodComune, mAnnoRiferimento)
            
            
            QUV = rsDatiTecniciComune("QRifiutiUtenzaDomestica") / Atot
            QUVsav = QUV
        
            'Parte variabile
            ii = 0
            rs.MoveFirst
            Do While Not rs.EOF
                Set Item = .ListItems.Add(, "%" & rs("IdNumeroComponenti"), rs("Descrizione"))
                Tariffa = QUV * rs("valore") * CUD
                Item.SubItems(1) = Format$(Tariffa, FORMAT6DEC)
                Item.ListSubItems(1).Tag = SetValueInTag(TAG_TAR_PVUD_FLD, Tariffa, "")
                ii = ii + 1
                rs.MoveNext
            Loop
        End If
        
        'Parte Fissa
        '--------------------------------------------------------------
        'Seleziono i KA
        Set rs = SelezionaCoeffK(CoeffKA, mCodComune, mAnnoRiferimento, rsDatiAnagraficiComune("AreaGeografica"), mParAbitanti)

        TotSuperficie = 0
        i = 0
        If rs.EOF Then Info "Non Trovati Coefficienti KA": Exit Function
        
        rs.MoveFirst
        Do While Not rs.EOF
            TotSuperficie = TotSuperficie + (SupCorrettaSenzaRiduzioneUD(i) * rs("Valore"))
            rs.MoveNext
            i = i + 1
        Loop
        
       
        'Calcolo
        rs.MoveFirst
        Do While Not rs.EOF
            For i = 1 To lvw4.ColumnHeaders.Count - 1
                    
                QUF = (lvw4.ListItems("DOMESTICA").SubItems(i) / TotSuperficie)
                Tariffa = rs("Valore") * QUF
                lvw5.ListItems("%" & rs("IdNumeroComponenti")).SubItems(i + 1) = Format$(Tariffa, FORMAT6DEC)
                If lvw4.ColumnHeaders(i + 1).Key = keyColUDA Then
                    'mi salvo il QUF usato con la colonna selezionata come tariffa
                    QUFUD = QUF
                    QUFUDsav = QUFUD
                
                    Set Item = lvw5.ListItems("%" & rs("IdNumeroComponenti"))
                    Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 2).Tag = SetValueInTag(TAG_TAR_PFUD_FLD, Tariffa, Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 2).Tag)
                End If
                
                
            Next
            rs.MoveNext
        Loop

        AllargaLeColonneLVW lvw5, True, 2
        
        .Visible = True
    End With
    TariffeUtenzeDomesticheNormalizzato = True
    
End Function


Private Function TariffeUtenzeAttivitaNormalizzato(Optional ConRiduzione As Boolean = False) As Boolean
    'Calcolo Tariffa Utenze Attivita con il metodo Normalizzato
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim qry As ADODB.Command
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim rsSvuotamentiUA As ADODB.Recordset
    Dim i As Integer, ii As Integer
    Dim Passo As Double
    Dim PERCINIUD As Double
    Dim PercIniUA As Double
    Dim Estremo  As Double
    Dim tariffaV As Double
    Dim s As Long
    Dim TotSuperficie As Double
    Dim QUAP As Double
    Dim TotSuperficieInRiduzioniDAUA_Pv As Double
    Dim TotSuperficieInRiduzioniDAUA_PF As Double
    Dim iUbound As Long
    
    With lvw7
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        TariffeUtenzeAttivitaNormalizzato = False
        .ColumnHeaders.Add , , "Utenza Attività", 3500
        
        Passo = Txt(PAG2_N_PASSO_ID)
        PERCINIUD = Txt(PAG2_N_PERCINIUD_ID)
        PercIniUA = Txt(PAG2_N_PERCINIUA_ID)
        
        If (PERCINIUD > PercIniUA) Then
            Estremo = (100 - Passo - PERCINIUD) / Passo
        Else
            Estremo = Abs((PercIniUA - PERCINIUD) / Passo)
        End If

        
        For i = 0 To Estremo
            Set ColItem = .ColumnHeaders.Add(, , "%" & CStr(PERCINIUD + Passo * i) & " - %" & CStr(100 - (PERCINIUD + Passo * i)), , lvwColumnCenter)
            ColItem.Icon = ICON_NULL
            ColItem.Alignment = lvwColumnRight
            ColItem.Key = "COL" & CStr(i)
            If ColItem.Key = keyColUDA Then
                ColItem.Icon = ICON_DOMESTICA_ATTIVITA
            End If
        Next
        
        .ListItems.Clear
            
        'KD
        Set rs1 = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsDatiAnagraficiComune("AreaGeografica"), mParAbitanti)
        If rs1.EOF Then Info "Non Trovati Coefficienti KD": Exit Function
        
       
        s = 1
        'Seleziono i KC
        Set rs = SelezionaCoeffK(CoeffKC, mCodComune, mAnnoRiferimento, rsDatiAnagraficiComune("AreaGeografica"), mParAbitanti)
        
        If rs.EOF Then Info "Occorre inserire i coefficienti KC": Exit Function
        rs.MoveFirst
        
        
        TotSuperficie = 0
        Erase SupConRiduzioneUAPF
        Erase SupConRiduzioneUAPV
        Do While Not rs.EOF
            
            
            If rs("Superficie") > 0 Then
            
                iUbound = -1
                On Error Resume Next
                iUbound = UBound(SupConRiduzioneUAPF)
                On Error GoTo 0
                ReDim Preserve SupConRiduzioneUAPF(iUbound + 1)
                ReDim Preserve SupConRiduzioneUAPV(iUbound + 1)
                
                TotaliDaRiduzioniUA rs("CodTipologiaAttivita"), rsRIDUA, TotSuperficieInRiduzioniDAUA_Pv, TotSuperficieInRiduzioniDAUA_PF
                
                SupConRiduzioneUAPF(iUbound + 1) = rs("Superficie") - TotSuperficieInRiduzioniDAUA_PF
                SupConRiduzioneUAPV(iUbound + 1) = rs("Superficie") - TotSuperficieInRiduzioniDAUA_Pv
            End If
            
            TotSuperficie = TotSuperficie + (rs("valore") * rs("Superficie"))
            
            rs.MoveNext
            rs1.MoveNext
        Loop
        
        
        'Calcolo
        rs.MoveFirst
        rs1.MoveFirst
        QUAPUA = 0

        If mOptMetodoTariffa = TariffaPortaAPorta Then
            Set rsSvuotamentiUA = SelezionaTotSvuotamentiUA(mCodComune, mAnnoRiferimento)
        End If
        
        On Error GoTo TariffeUtenzeAttivitaNormalizzatoERR:
        Do While Not rs.EOF
            If rs("Superficie") > 0 Then
                Set Item = lvw7.ListItems.Add(, "%" & rs("CodTipologiaAttivita"), rs("Descrizione"))
                For i = 1 To lvw4.ColumnHeaders.Count - 1
                    QUAP = lvw4.ListItems("ATTIVITA").SubItems(i) / TotSuperficie
                    Item.SubItems(i) = Format$((CUA * rs1("valore") * s) + rs("Valore") * QUAP, FORMAT6DEC)
                    
                    If lvw4.ColumnHeaders(i + 1).Key = keyColUDA Then
                        'mi salvo il QUAP usato con la colonna selezionata come tariffa
                        QUAPUA = QUAP
                        QUAPUAsav = QUAPUA
                        'mi salvo la parte variabile della tariffa per mostrarla alla fine del calcolo
                        If mOptMetodoTariffa = TariffaPortaAPorta Then
                            Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag = SetValueInTag(TAG_TAR_PVUA_FLD, rsCostiComune("ParteVariabile").value / (rsSvuotamentiUA("TotaleSvuotamentoPerComune").value * mCoeffMedioSvuotamentoLT), "")
                        Else
                            Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag = SetValueInTag(TAG_TAR_PVUA_FLD, CUA * rs1("valore") * s, "")
                        End If
                        'mi salvo la parte fissa della tariffa per mostrarla alla fine del calcolo
                        Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag = SetValueInTag(TAG_TAR_PFUA_FLD, rs("Valore") * QUAP, Item.ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
                    End If
                Next
                If mOptMetodoTariffa = TariffaPortaAPorta Then
                    If Not rsSvuotamentiUA.EOF Then rsSvuotamentiUA.MoveNext
                End If
            End If
            rs1.MoveNext
            rs.MoveNext

        Loop

        AllargaLeColonneLVW lvw7, True, -2
        
        .Visible = True
    End With
    
    
    Set rs = Nothing
    Set rs1 = Nothing
    TariffeUtenzeAttivitaNormalizzato = True
    Exit Function
    
TariffeUtenzeAttivitaNormalizzatoERR:
    
MsgBox "Non Trovati Svuotamenti Utenza Attivita"


End Function


Function CacolaCUDomesticoEAttivita(CUAPerc As Double, CUDPerc As Double, Optional QRifiutiUDRiduzioni As Double = 0, Optional QRifiutiUARiduzioni As Double = 0) As Boolean
    Dim TotaleServizio As Double
    
    CacolaCUDomesticoEAttivita = False
    'percentuale rifiuti UD
    If CUAPerc = 0 Or CUDPerc = 0 Then
        CUDPerc = ((rsDatiTecniciComune("QRifiutiUtenzaDomestica") - QRifiutiUDRiduzioni) / (rsDatiTecniciComune("QRifiutiTotali") - QRifiutiUDRiduzioni - QRifiutiUARiduzioni)) * 100
        'percentuale rifiuti UA
        CUAPerc = ((rsDatiTecniciComune("QRifiutiCategorieAttivita") - QRifiutiUARiduzioni) / (rsDatiTecniciComune("QRifiutiTotali") - QRifiutiUDRiduzioni - QRifiutiUARiduzioni)) * 100
    End If
    
    'Calcolo Totali Quote
    TotaleServizio = CDbl(Txt(PAG1_N_ET_ID))
    ToTQVUD = (TotaleServizio - CDbl(Txt(PAG1_N_ETF_ID))) * (CUDPerc / 100)
    ToTQVUA = (TotaleServizio - CDbl(Txt(PAG1_N_ETF_ID))) * (CUAPerc / 100)
        
    CUD = (TotaleServizio - CDbl(Txt(PAG1_N_ETF_ID))) * (CUDPerc / 100) / (rsDatiTecniciComune("QRifiutiUtenzaDomestica") - QRifiutiUDRiduzioni)
    CUDsav = CUD
    CUA = (TotaleServizio - CDbl(Txt(PAG1_N_ETF_ID))) * (CUAPerc / 100) / (rsDatiTecniciComune("QRifiutiCategorieAttivita") - QRifiutiUARiduzioni)
    CUAsav = CUA
    
    CacolaCUDomesticoEAttivita = True
End Function


Private Function PaginaRiassuntivaTariffaUD(Optional ConRiduzione As Boolean = False) As Boolean
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim RsDAtiUD As ADODB.Recordset
    Dim rsTotSvuotamentiUD As ADODB.Recordset
    Dim NComFamiglia As Double
    Dim TotUtenzeInRiduzioniDAUD_PV As Double
    Dim TariffaUDV As Double
    Dim TariffaUDF As Double
    Dim iUbound As Long
    
    PaginaRiassuntivaTariffaUD = False
    With lvw6
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        'se il metodo calcolo è porta a porta cambiano le intestazioni di colonna
        Set ColItem = .ColumnHeaders.Add(, , "Utenza Domestica")
        Set ColItem = .ColumnHeaders.Add(, , "n. Ruoli")
        Set ColItem = .ColumnHeaders.Add(, , "Superficie Tot.")
        Set ColItem = .ColumnHeaders.Add(, , "Superf. Abitaz.")
        
        If mOptMetodoTariffa = TariffaPortaAPorta Then
            Set ColItem = .ColumnHeaders.Add(, , "Litri Svuotati")
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa PV")
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa PF")
            Set ColItem = .ColumnHeaders.Add(, , "C/Medio Tariff.")
        Else
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa")
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa (n-1)")
            Set ColItem = .ColumnHeaders.Add(, , "Differenza %")
        End If
        
        For i = 2 To .ColumnHeaders.Count
            .ColumnHeaders(i).Alignment = lvwColumnRight
        Next
        
        
        .ListItems.Clear
        
        Set rs = gDBC.Execute("Select * from NumeroComponentiFamiglia")
        Set RsDAtiUD = gDBC.Execute("select * from UtenzeDomesticheComunali where CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "' order by IdNumeriComponenti")
    
        Erase NumeroUtenzaDomesticaConRiduzione
        
        'Utenze domestiche
        i = 0
        GettitoCalcUDPF = 0
        GettitoCalcUDPV = 0
        Do While Not rs.EOF
            Set Item = .ListItems.Add(, "%" & rs("NumeroComponenti"), rs("Descrizione"))

            'Numero Componenti
            Select Case rs("NumeroComponenti")
                Case 1
                    NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("NumeroComponenti")) & "Componente")
                Case Is >= 2
                    If rs("NumeroComponenti") = 6 Then
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                    Else
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs("NumeroComponenti")) & "Componenti")
                    End If
            End Select
            
            'seconda casa
            If rsDatiTecniciComune("Id_NumCompSecondaCasa") = rs("NumeroComponenti") Then
                NComFamiglia = NComFamiglia + rsDatiTecniciComune("NRuoliSecondaCasa")
            End If
            
            Item.SubItems(1) = NComFamiglia
            Item.SubItems(2) = Format$(CDbl(SupCorrettaSenzaRiduzioneUD(i)), FORMATVALUTA)
            
            'salvataggio dei ruoli meno le riduzioni
            iUbound = -1
            On Error Resume Next
            iUbound = UBound(NumeroUtenzaDomesticaConRiduzione)
            On Error GoTo 0
            ReDim Preserve NumeroUtenzaDomesticaConRiduzione(iUbound + 1)
            
            TotaliDaRiduzioniUD rs("NumeroComponenti"), rsRIDUD, TotUtenzeInRiduzioniDAUD_PV
            NumeroUtenzaDomesticaConRiduzione(iUbound + 1) = NComFamiglia - TotUtenzeInRiduzioniDAUD_PV
            
            
            If NComFamiglia = 0 Then
                Item.SubItems(3) = 0
                Item.SubItems(4) = 0
            Else
                TariffaUDV = GetValueInTag(TAG_TAR_PVUD_FLD, lvw5.ListItems(i + 1).ListSubItems(1).Tag)
                TariffaUDF = GetValueInTag(TAG_TAR_PFUD_FLD, lvw5.ListItems(i + 1).ListSubItems(idxColonnaSelezionata(keyColUDA) + 2).Tag)
                
                If mOptMetodoTariffa = TariffaPortaAPorta Then
                
                    If rsTotSvuotamentiUD Is Nothing Then
                        Set rsTotSvuotamentiUD = SelezionaTotSvuotamentiUD(mCodComune, mAnnoRiferimento)
                    End If
                    
                    Item.SubItems(3) = Format$(CDbl(Item.SubItems(2)) / NComFamiglia, FORMATSTANDARD)
                    Item.SubItems(4) = rsTotSvuotamentiUD!totSvuotamenti * mCoeffMedioSvuotamentoLT
                    Item.SubItems(5) = Format$(TariffaUDV, FORMAT6DEC)
                    Item.SubItems(6) = Format$((TariffaUDF * CDbl(Item.SubItems(3))), FORMAT6DEC)
                    'TF + (Svuotamenti/Nruoli) * TV
                    Item.SubItems(7) = Format$(CDbl(Item.SubItems(6)) + (TariffaUDV * (CDbl(Item.SubItems(4)) / CDbl(Item.SubItems(1)))), FORMAT3DEC)
                    
                    
                    
                Else
                    Item.SubItems(3) = Format$(CDbl(Item.SubItems(2)) / NComFamiglia, FORMATSTANDARD)
                    Item.SubItems(4) = Format$((TariffaUDV + (TariffaUDF * CDbl(Item.SubItems(3)))), FORMAT6DEC)
                    
                    'calcolo anno precedente
                    If Not RsDAtiUD.EOF Then
                        Item.SubItems(5) = Format$(RsDAtiUD("TariffaNComp") + RsDAtiUD("TariffaSuperficie") * CDbl(Item.SubItems(3)), FORMAT6DEC)
                        If Val(Item.SubItems(5)) <> 0 Then
                            Item.SubItems(6) = Format$(((CDbl(Item.SubItems(4)) - CDbl(Item.SubItems(5))) / CDbl(Item.SubItems(5)) * 100), "##0.00")
                        Else
                            Item.SubItems(6) = "manca tariffa precedente"
                        End If
                    Else
                        Item.SubItems(5) = 0
                        Item.SubItems(6) = 0
                    End If
                End If
            End If
            
            
            Dim NComFamigliaR As Double, SuperfR As Double
            
            'ANTONIO 26-02-02 TotaliDaRiduzioniUD rs("NumeroComponenti"), rsRIDUD, NComFamigliaR, SuperfR
                                           
            'Totale Gettito UD Parte Fissa
            GettitoCalcUDPF = GettitoCalcUDPF + TariffaUDF * (CDbl(Item.SubItems(2)) - SuperfR)

            If mOptMetodoTariffa = TariffaPortaAPorta Then
                'Totale Gettito UD Parte Variabile
                GettitoCalcUDPV = GettitoCalcUDPV + (TariffaUDV * (rsTotSvuotamentiUD!totSvuotamenti * mCoeffMedioSvuotamentoLT))
            Else
                'Totale Gettito UD Parte Variabile
                GettitoCalcUDPV = GettitoCalcUDPV + (TariffaUDV * (NComFamiglia - NComFamigliaR))
            End If
            i = i + 1
            rs.MoveNext
            If Not RsDAtiUD.EOF Then RsDAtiUD.MoveNext
            If mOptMetodoTariffa = TariffaPortaAPorta Then
                If Not rsTotSvuotamentiUD.EOF Then rsTotSvuotamentiUD.MoveNext
            End If
        Loop
        
        
        AllargaLeColonneLVW lvw6, False, 2
        
        .Visible = True

    End With
    
    PaginaRiassuntivaTariffaUD = True
End Function


Private Function PaginaRiassuntivaTariffaUA(Optional ConRiduzione As Boolean = False) As Boolean
    Dim ColItem As ColumnHeader
    Dim Item As ListItem
    Dim i As Integer
    Dim RsDAtiUA As ADODB.Recordset
    Dim RsDAtiUAPrec As ADODB.Recordset
    Dim ParteFissa As Double
    Dim ParteVariabile As Double
    Dim rsSvuotamentiUA As ADODB.Recordset
    
    PaginaRiassuntivaTariffaUA = False
    With lvw8
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        
        If mOptMetodoTariffa = TariffaPortaAPorta Then
            Set rsSvuotamentiUA = SelezionaTotSvuotamentiUA(mCodComune, mAnnoRiferimento)
            
            Set ColItem = .ColumnHeaders.Add(, , "Categoria", 4500)
            Set ColItem = .ColumnHeaders.Add(, , "Superficie")
            Set ColItem = .ColumnHeaders.Add(, , "n. Ruoli")
            Set ColItem = .ColumnHeaders.Add(, , "Sup.Media")
            Set ColItem = .ColumnHeaders.Add(, , "Litri Svuotati")
            Set ColItem = .ColumnHeaders.Add(, , "Tar.PF")
            Set ColItem = .ColumnHeaders.Add(, , "Tar.PV")
            Set ColItem = .ColumnHeaders.Add(, , "C/Medio Tarif.")
            Set ColItem = .ColumnHeaders.Add(, , "")
            Set ColItem = .ColumnHeaders.Add(, , "")
            
        Else
            Set ColItem = .ColumnHeaders.Add(, , "Categoria", 4500)
            Set ColItem = .ColumnHeaders.Add(, , "Superficie")
            Set ColItem = .ColumnHeaders.Add(, , "n. Ruoli")
            Set ColItem = .ColumnHeaders.Add(, , "Sup.Media")
            Set ColItem = .ColumnHeaders.Add(, , "Tar.PF")
            Set ColItem = .ColumnHeaders.Add(, , "Tar.PV")
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa")
            Set ColItem = .ColumnHeaders.Add(, , "Tariffa(n-1)")
            Set ColItem = .ColumnHeaders.Add(, , "Differenza%")
        
            
        End If
        
        For i = 2 To .ColumnHeaders.Count
            .ColumnHeaders(i).Alignment = lvwColumnRight
        Next
        
        
        .ListItems.Clear
        
        Set RsDAtiUA = gDBC.Execute("SELECT CategorieAttivitaComunali.*, CategorieAttivitaComunali.Descrizione AS Categoria " & _
                         "FROM CategorieAttivitaComunali WHERE  Superficie> 0 and NumRuoli>0 and CategorieAttivitaComunali.CodComune='" & mCodComune & "' and CategorieAttivitaComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by IdCategoriaAttivita")
         
        
        GettitoCalcUAPF = 0
        GettitoCalcUAPV = 0
        Do While Not RsDAtiUA.EOF
            
            Set Item = .ListItems.Add(, "%" & RsDAtiUA("IdCategoriaAttivita"), RsDAtiUA("Categoria"))
    
            Item.SubItems(1) = Format$(RsDAtiUA("Superficie"), FORMATSTANDARD)
            Item.SubItems(2) = Format$(RsDAtiUA("NumRuoli"), FORMATSTANDARD)
            Item.SubItems(3) = Format$(RsDAtiUA("Superficie") / RsDAtiUA("NumRuoli"), "###,###,###,##0,00")
            ParteFissa = 0 + GetValueInTag(TAG_TAR_PFUA_FLD, lvw7.ListItems("%" & RsDAtiUA("IdCategoriaAttivita")).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
            ParteVariabile = 0 + GetValueInTag(TAG_TAR_PVUA_FLD, lvw7.ListItems("%" & RsDAtiUA("IdCategoriaAttivita")).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
            If mOptMetodoTariffa = TariffaPortaAPorta Then
                Item.SubItems(4) = rsSvuotamentiUA!totSvuotamenti * mCoeffMedioSvuotamentoLT
                Item.SubItems(5) = Format$(ParteFissa, FORMAT6DEC)
                Item.SubItems(6) = Format$(ParteVariabile, FORMAT6DEC)
                '(TF*SupMedia)+((svuot/nRuoli)*TV)
                Item.SubItems(7) = Format$((ParteFissa * CDbl(Item.SubItems(3))) + ((CDbl(Item.SubItems(4) / Item.SubItems(2))) * ParteVariabile), FORMAT3DEC)
                Item.SubItems(8) = ""
            Else
                Item.SubItems(4) = Format$(ParteFissa, FORMAT6DEC)
                Item.SubItems(5) = Format$(ParteVariabile, FORMAT6DEC)
                Item.SubItems(6) = Format$(ParteFissa + ParteVariabile, FORMAT6DEC)
            End If
            
            Dim SuperfRPF As Double
            Dim SuperfRPV As Double
            
            'ANTONIO 26-02-02TotaliDaRiduzioniUA RsDAtiUA("IdCategoriaAttivita"), rsRIDUA, SuperfRPV, SuperfRPF
            
            'Totale Gettito UA parte Fissa
            GettitoCalcUAPF = GettitoCalcUAPF + ParteFissa * (RsDAtiUA("Superficie") - SuperfRPF)
            If mOptMetodoTariffa = TariffaPortaAPorta Then
                'Totale Gettito UD Parte Variabile
                GettitoCalcUAPV = GettitoCalcUAPV + (ParteVariabile * (rsSvuotamentiUA!totSvuotamenti * mCoeffMedioSvuotamentoLT))
            Else
                'Totale Gettito UA parte Variabile
                GettitoCalcUAPV = GettitoCalcUAPV + ParteVariabile * (RsDAtiUA("Superficie"))
            End If
            
            Set RsDAtiUAPrec = gDBC.Execute("select * from CategorieAttivitaComunali where   CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "' and IdCAtegoriaAttivita =" & RsDAtiUA("IdCategoriaAttivita"))
            
            If Not RsDAtiUAPrec.EOF And (mOptMetodoTariffa = TariffaNormalizzata Or mOptMetodoTariffa = TariffaCostiIndivisibili) Then
                Item.SubItems(7) = Format$(RsDAtiUAPrec("Tariffa"), FORMAT6DEC)
                If RsDAtiUAPrec("Tariffa") > 0 Then
                    Item.SubItems(8) = Format$((CDbl(Item.SubItems(6)) - RsDAtiUAPrec("Tariffa")) / RsDAtiUAPrec("Tariffa") * 100, "##0.00")
                Else
                    Item.SubItems(8) = "manca tariffa precedente"
                End If
            ElseIf mOptMetodoTariffa = TariffaPortaAPorta Then
                Item.SubItems(8) = ""
                Item.SubItems(9) = ""
            End If
            RsDAtiUAPrec.Close
            RsDAtiUA.MoveNext
            If mOptMetodoTariffa = TariffaPortaAPorta Then
                If Not rsSvuotamentiUA.EOF Then rsSvuotamentiUA.MoveNext
            End If
        Loop
        
        AllargaLeColonneLVW lvw8, False, -2
        
        .Visible = True

    End With
    PaginaRiassuntivaTariffaUA = True
End Function


Function idxColonnaSelezionata(s As String) As Long
    Dim i As Long
    
    i = InStr(1, s, "COL")
    idxColonnaSelezionata = Val(Mid$(s, i + 3))
End Function



Private Function CacoloTipoAdeguamento(TipoAdeguamento As Integer, valuta As String) As Boolean
    Dim Item As ListItem
    Dim ColItem As ColumnHeader
    Dim i As Integer
    Dim sKey As String
    Dim numRowUD As Integer
    Dim numRowUA As Integer
    Dim sDesUA As String
    Dim sDesUD As String
    Dim sTariffaUDF As String
    Dim sTariffaUDV As String
    Dim sTariffaUA As String
    Dim sTariffaUA_PF As String
    Dim sTariffaUA_PV As String
    Dim FatCorrettivo As Double
    
    CacoloTipoAdeguamento = False
    Select Case TipoAdeguamento
        Case 0
            Txt(PAG7_FATTORECORRETTIVO_ID) = Format$(Txt(PAG7_GUADAGNO_ID), "##0.00")
        Case 1
            Txt(PAG7_FATTORECORRETTIVO_ID) = Format$(1 - (CDbl(Txt(PAG7_GETTITODICHIARATO_ID)) - (GettitoCalcUDPF + GettitoCalcUDPV + GettitoCalcUAPF + GettitoCalcUAPV)) / (GettitoCalcUAPF + GettitoCalcUAPV), "##0.000")
        Case 2
            Txt(PAG7_FATTORECORRETTIVO_ID) = Format$(1 - (CDbl(Txt(PAG7_GETTITODICHIARATO_ID)) - (GettitoCalcUDPF + GettitoCalcUDPV + GettitoCalcUAPF + GettitoCalcUAPV)) / GettitoCalcUDPV, "##0.000")
    End Select

    FatCorrettivo = CDbl(Txt(PAG7_FATTORECORRETTIVO_ID))
    
    With lvw9
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        Set ColItem = .ColumnHeaders.Add(, , "Utenza Domestica")
        ColItem.Alignment = lvwColumnLeft
        Set ColItem = .ColumnHeaders.Add(, , "TariffaF")
        ColItem.Alignment = lvwColumnRight
        Set ColItem = .ColumnHeaders.Add(, , "TariffaV")
        ColItem.Alignment = lvwColumnRight
        Set ColItem = .ColumnHeaders.Add(, , "   ", 200)
        ColItem.Alignment = lvwColumnLeft
        Set ColItem = .ColumnHeaders.Add(, , "Categoria", 4500)
        ColItem.Alignment = lvwColumnLeft
        Set ColItem = .ColumnHeaders.Add(, , "Tariffa")
        ColItem.Alignment = lvwColumnRight
        Set ColItem = .ColumnHeaders.Add(, , "TariffaF")
        ColItem.Alignment = lvwColumnRight
        Set ColItem = .ColumnHeaders.Add(, , "TariffaV")
        ColItem.Alignment = lvwColumnRight
        

        .ListItems.Clear

        'Lettura listview con tariffa UD e UA
        numRowUD = lvw6.ListItems.Count
        numRowUA = lvw8.ListItems.Count
        
        
        For i = 1 To IIf(numRowUD <= numRowUA, numRowUA, numRowUD)
            
            sTariffaUDF = vbNullString
            sTariffaUDV = vbNullString
            sDesUA = vbNullString
            sTariffaUA = vbNullString
            sTariffaUA_PF = vbNullString
            sTariffaUA_PV = vbNullString
            If i <= numRowUD Then
                sKey = SetValueInTag(TAG_UD_FLD, lvw6.ListItems(i).Key, sKey)
                sDesUD = lvw6.ListItems(i).Text
                sTariffaUDV = GetValueInTag(TAG_TAR_PVUD_FLD, lvw5.ListItems(lvw6.ListItems(i).Key).ListSubItems(1).Tag)
                sTariffaUDF = GetValueInTag(TAG_TAR_PFUD_FLD, lvw5.ListItems(lvw6.ListItems(i).Key).ListSubItems(idxColonnaSelezionata(keyColUDA) + 2).Tag)
                Select Case TipoAdeguamento
                    Case 0
                        sTariffaUDF = sTariffaUDF / FatCorrettivo
                        sTariffaUDV = sTariffaUDV / FatCorrettivo
                    Case 2
                        sTariffaUDV = sTariffaUDV / FatCorrettivo
                End Select
                
            Else
                sKey = SetValueInTag(TAG_UD_FLD, -1, sKey)
                sDesUD = " "
                sTariffaUDF = " "
                sTariffaUDV = " "
            End If
            
            If i <= numRowUA Then
                sKey = SetValueInTag(TAG_UA_FLD, lvw8.ListItems(i).Key, sKey)
                sDesUA = lvw8.ListItems(i).Text
                sTariffaUA_PF = GetValueInTag(TAG_TAR_PFUA_FLD, lvw7.ListItems(lvw8.ListItems(i).Key).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
                sTariffaUA_PV = GetValueInTag(TAG_TAR_PVUA_FLD, lvw7.ListItems(lvw8.ListItems(i).Key).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
                Select Case TipoAdeguamento
                    Case 0
                        sTariffaUA_PF = sTariffaUA_PF / FatCorrettivo
                        sTariffaUA_PV = sTariffaUA_PV / FatCorrettivo
                    Case 1
                        sTariffaUA_PF = sTariffaUA_PF / FatCorrettivo
                        sTariffaUA_PV = sTariffaUA_PV / FatCorrettivo
                End Select
                
                sTariffaUA = CDbl(sTariffaUA_PF) + CDbl(sTariffaUA_PV)
            Else
                sKey = SetValueInTag(TAG_UA_FLD, -1, sKey)
                sDesUA = " "
                sTariffaUA = " "
            End If
            
            Set Item = .ListItems.Add(, sKey, sDesUD)
            'mi salvo la tariffa UD
            Item.SubItems(1) = InValutaEst(sTariffaUDF, valuta, 6)
            Item.ListSubItems(1).Tag = SetValueInTag(TAG_TAR_PFUD_FLD, sTariffaUDF, "")
            Item.SubItems(2) = InValutaEst(sTariffaUDV, valuta, 6)
            Item.ListSubItems(2).Tag = SetValueInTag(TAG_TAR_PVUD_FLD, sTariffaUDV, "")
            Item.SubItems(4) = sDesUA
            Item.SubItems(5) = InValutaEst(sTariffaUA, valuta, 6)
            Item.ListSubItems(5).Tag = SetValueInTag(TAG_TAR_PFUA_FLD, sTariffaUA, "")
            
            'mi salvo la tariffa UA
            Item.SubItems(6) = InValutaEst(sTariffaUA_PF, valuta, 6)
            Item.ListSubItems(6).Tag = SetValueInTag(TAG_TAR_PFUA_FLD, sTariffaUA_PF, "")
            Item.SubItems(7) = InValutaEst(sTariffaUA_PV, valuta, 6)
            Item.ListSubItems(7).Tag = SetValueInTag(TAG_TAR_PVUA_FLD, sTariffaUA_PV, "")
        Next
        
        For i = 1 To numRowUD + 9 - lvw9.ListItems.Count
            Set Item = lvw9.ListItems.Add(, "NoData" & CStr(i), "")
            Item.ListSubItems.Add , , ""
        Next
        
        AllargaLeColonneLVW lvw9, False, 1, 1, 1, -1, -1
        
        .Visible = True
    End With
    
    CacoloTipoAdeguamento = True
End Function


Function SalvaTariffeMetodoIntegrale() As Boolean
        Dim Item As ListItem
        Dim sKey As String
        Dim TariffaPF As Double
        Dim TariffaPV As Double
        Dim NumRuoli As Long
        
        For Each Item In lvw9.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKey = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKey <> "-1" Then
                TariffaPF = GetValueInTag(TAG_TAR_PFUD_FLD, Item.ListSubItems(1).Tag)
                TariffaPV = GetValueInTag(TAG_TAR_PVUD_FLD, Item.ListSubItems(2).Tag)
                NumRuoli = lvw6.ListItems(sKey).ListSubItems(1)
                gDBC.Execute "Update UtenzeDomesticheComunali set NumeroUtenze=" & NumRuoli & ",TariffaSuperficie=" & FormatFLD(TariffaPF, adDouble) & ", TariffaNComp=" & FormatFLD(TariffaPV, adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "' and IdNumeriComponenti=" & Mid$(sKey, 2) & ""
            End If
            
            sKey = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKey <> "-1" Then
                TariffaPV = GetValueInTag(TAG_TAR_PVUA_FLD, Item.ListSubItems(7).Tag)
                TariffaPF = GetValueInTag(TAG_TAR_PFUA_FLD, Item.ListSubItems(6).Tag)
            
                gDBC.Execute "Update CategorieAttivitaComunali set TariffaPV=" & FormatFLD(TariffaPV, adDouble) & ",TariffaPF=" & FormatFLD(TariffaPF, adDouble) & ",Tariffa=" & FormatFLD(CDbl(Item.SubItems(5)), adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "' and IdCategoriaAttivita=" & Mid$(sKey, 2) & ""
            End If
        Next
        
        If mOptMetodoTariffa = TariffaPortaAPorta Then
            'salvo la tariffa x tipo contenitore
            gDBC.Execute "Update ContenitoriComunali set CostoUnitario=id*" & FormatFLD(TariffaPV, adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
        End If
        
        
        'Dati di Analsi
        
        QUFUD = Round(QUFUD, 6)
        CUD = Round(CUD, 6)
        CUA = Round(CUA, 6)
        QUAPUA = Round(QUAPUA, 6)
        QUV = Round(QUV, 6)
        
        If GettitoCalcUDPVRidotto > 0 Then ToTQVUD = GettitoCalcUDPVRidotto
        If GettitoCalcUDPFRidotto > 0 Then TotQFUD = GettitoCalcUDPFRidotto
        If GettitoCalcUAPFRidotto > 0 Then ToTQFUA = GettitoCalcUAPFRidotto
        If GettitoCalcUAPVRidotto > 0 Then ToTQVUA = GettitoCalcUAPVRidotto
        
        'Salvataggio dati finali calcolo tariffa
        On Error Resume Next
        gDBC.Execute "INSERT INTO DatiFinaliCalcoloTariffa ( CodComune,CodTipoTariffa, AnnoRiferimento, QVDOM, QFDOM, QUVDOM, QUFDOM, CUDOM, QVNDOM, QFNDOM, CUNDOM, QAPFNDOM,GettitoUD,GettitoUA, IntroitoDaUA_PF,IntroitoDaUA_PV,IntroitoDaUD_PV,IntroitoDaUD_PF) values(" & _
            FormatFLD(mCodComune, adVarChar) & "," & mOptMetodoTariffa & "," & FormatFLD(mAnnoRiferimento, adVarChar) & "," & FormatFLD(ToTQVUD, adDouble) & "," & FormatFLD(TotQFUD, adDouble) & "," & FormatFLD(QUV, adDouble) & "," & _
            FormatFLD(QUFUD, adDouble) & "," & FormatFLD(CUD, adDouble) & "," & FormatFLD(ToTQVUA, adDouble) & "," & FormatFLD(ToTQFUA, adDouble) & "," & FormatFLD(CUA, adDouble) & "," & FormatFLD(QUAPUA, adDouble) & "," & _
            FormatFLD(TotQFUD + ToTQVUD, adDouble) & "," & FormatFLD(ToTQFUA + ToTQVUA, adDouble) & "," & FormatFLD(TotMancatoIntroitoDAUA_PF, adDouble) & "," & FormatFLD(TotMancatoIntroitoDAUA_PV, adDouble) & "," & FormatFLD(TotMancatoIntroitoDAUD_PV, adDouble) & "," & FormatFLD(TotMancatoIntroitoDAUD_PF, adDouble) & ")"
        If gDBC.Errors.Count > 0 Then
            If gDBC.Errors(0).Number = -2147467259 Or gDBC.Errors(0).NativeError = -105121349 Then
                On Error GoTo 0
                gDBC.Execute "UPDATE DatiFinaliCalcoloTariffa SET " & _
                    "CodComune =" & FormatFLD(mCodComune, adVarChar) & ",AnnoRiferimento=" & FormatFLD(mAnnoRiferimento, adVarChar) & ",QVDOM=" & FormatFLD(ToTQVUD, adDouble) & ",QFDOM=" & FormatFLD(TotQFUD, adDouble) & ",QUVDOM=" & FormatFLD(QUV, adDouble) & _
                    ",QUFDOM=" & FormatFLD(QUFUD, adDouble) & ",CUDOM=" & FormatFLD(CUD, adDouble) & ",QVNDOM=" & FormatFLD(ToTQVUA, adDouble) & ",QFNDOM=" & FormatFLD(ToTQFUA, adDouble) & ",CUNDOM=" & FormatFLD(CUA, adDouble) & ",QAPFNDOM=" & FormatFLD(QUAPUA, adDouble) & _
                    ",GettitoUD= " & FormatFLD(TotQFUD + ToTQVUD, adDouble) & ",GettitoUA= " & FormatFLD(ToTQFUA + ToTQVUA, adDouble) & _
                    ",IntroitoDaUA_PF= " & FormatFLD(TotMancatoIntroitoDAUA_PF, adDouble) & ",IntroitoDaUA_PV= " & FormatFLD(TotMancatoIntroitoDAUA_PV, adDouble) & ",IntroitoDaUD_PV= " & FormatFLD(TotMancatoIntroitoDAUD_PV, adDouble) & ",IntroitoDaUD_PF= " & FormatFLD(TotMancatoIntroitoDAUD_PF, adDouble) & _
                    " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
            Else
                MsgBox "Non è stato possibile salvare i dati finali:" + vbCrLf + Error$, vbExclamation
            End If
        End If
End Function


Function SalvaTariffeMetodo507() As Boolean
        Dim Item As ListItem
        Dim sKey As String
        Dim whereCondition As String
        
        For Each Item In lvw10.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKey = GetValueInTag(TAG_UD_FLD, Item.Key)
            
            LetParamValue mAnnoRiferimento, mCodComune, "CHKCALPERSINGUD", True, chkTarsu507perNComponenti
            
            If chkTarsu507perNComponenti Then
                whereCondition = " and IdNumeriComponenti=" & Mid$(sKey, 2) & ""
            End If
            
            If sKey <> "-1" Then gDBC.Execute "Update UtenzeDomesticheComunali set TariffaSuperficie=" & FormatFLD(CDbl(Item.SubItems(4)), adDouble) & ", TariffaNComp=0 where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'" & whereCondition
            
            sKey = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKey <> "-1" Then gDBC.Execute "Update CategorieAttivitaComunali set TariffaPV=0,TariffaPF=0,Tariffa=" & FormatFLD(CDbl(Item.SubItems(1)), adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "' and IdCategoriaAttivita=" & Mid$(sKey, 2) & ""
        Next
        
        'Salvataggio dati finali calcolo tariffa
        On Error Resume Next
        gDBC.Execute "INSERT INTO DatiFinaliCalcoloTarsu507 ( CodComune, AnnoRiferimento, CostoTotale, Gettito, CMG, Copertura, " & _
                     "ProdMedia,CoeffCorrettivoIPS,IQSUtenzaDomestica,IPSUtenzaDomestica,CoeffCorrTariffa) values(" & _
                FormatFLD(mCodComune, adVarChar) & "," & FormatFLD(mAnnoRiferimento, adVarChar) & "," & FormatFLD(CDbl(Txt(PAG0_TOTSERVIZIO_ID)), adDouble) & "," & _
                FormatFLD(CDbl(Txt(PAG0_GETTITO_ID)), adDouble) & "," & FormatFLD(CMG, adDouble) & "," & _
                FormatFLD(CDbl(Txt(PAG0_TIPOTARIFFA_ID)), adDouble) & "," & FormatFLD(KM, adDouble) & "," & _
                FormatFLD(CoeffCorrettivo507IPS, adDouble) & "," & _
                FormatFLD(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), adDouble) & "," & _
                FormatFLD(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), adDouble) & "," & _
                FormatFLD(CoeffCorrTariffa507, adDouble) & ")"
             
             
        If gDBC.Errors.Count > 0 Then
            If gDBC.Errors(0).Number = -2147467259 Or gDBC.Errors(0).NativeError = -105121349 Then
                On Error GoTo 0
                gDBC.Execute "UPDATE DatiFinaliCalcoloTarsu507 SET " & _
                    "CodComune =" & FormatFLD(mCodComune, adVarChar) & ",AnnoRiferimento=" & FormatFLD(mAnnoRiferimento, adVarChar) & _
                    ",CostoTotale=" & FormatFLD(CDbl(Txt(PAG0_TOTSERVIZIO_ID)), adDouble) & ",Gettito=" & FormatFLD(CDbl(Txt(PAG0_GETTITO_ID)), adDouble) & ",CMG=" & FormatFLD(CMG, adDouble) & _
                    ",Copertura=" & FormatFLD(CDbl(Txt(PAG0_TIPOTARIFFA_ID)), adDouble) & ",ProdMedia=" & FormatFLD(KM, adDouble) & _
                    ",CoeffCorrettivoIPS=" & FormatFLD(CoeffCorrettivo507IPS, adDouble) & _
                    ",IQSUtenzaDomestica=" & FormatFLD(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), adDouble) & _
                    ",IPSUtenzaDomestica=" & FormatFLD(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), adDouble) & _
                    ",CoeffCorrTariffa=" & FormatFLD(CoeffCorrTariffa507, adDouble) & _
                    " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
            Else
                MsgBox "Non è stato possibile salvare i dati finali:" + vbCrLf + Error$, vbExclamation
            End If
        End If
        
       
        gDBC.Execute "Update ValoriCoefficientiIQSCAComunali set  IPSCalc= IPS*" & FormatFLD(CoeffCorrettivo507IPS, adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
        gDBC.Execute "Update ValoriCoefficienti507UDComunali set  ValoreCalc= Valore*" & FormatFLD(CoeffCorrettivo507IPS, adDouble) & " where AnnoRiferimento='" & mAnnoRiferimento & "' and CodComune='" & mCodComune & "'"
        
End Function


Sub MostraMancatoIntroitoConLeRiduzioni()
    'aggiorna le tariffe finali tendo conto delle riduzioni
    Dim TariffaPV As Double
    Dim TariffaPF As Double
    Dim UD As Integer
    Dim UA As Integer
    Dim Item As ListItem
    Dim sKeyNode As String
    Dim pKeyUA As String
    Dim pKeyUD As String
    Dim pKey As String
    Dim i As Integer
    Dim SiUA As Boolean
    Dim SiUD As Boolean
    
    TotMancatoIntroitoDAUA_PF = 0
    TotMancatoIntroitoDAUD_PF = 0
    TotMancatoIntroitoDAUA_PV = 0
    TotMancatoIntroitoDAUD_PV = 0
    
    pKeyUA = CreaLaPkey("BonusUtenzeAttivita")
    pKeyUD = CreaLaPkey("BonusUtenzeDomestiche")
    
    
    
    If Not rsRIDUA.EOF Then
        For i = 0 To rsRIDUA.fields.Count - 1
            If InStr(pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name) And rsRIDUA.fields(i).Name <> "Riduzione" And rsRIDUA.fields(i).Name <> "ParteFissa" And rsRIDUA.fields(i).Name <> "CodAttivita" Then
                MettivaloreInPKEY pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name, rsRIDUA.fields(i).value, rsRIDUA.fields(i).Type
            End If
        Next
        SiUA = True
    End If
    
    If Not rsRIDUD.EOF Then
        For i = 0 To rsRIDUD.fields.Count - 1
            If InStr(pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name) And rsRIDUD.fields(i).Name <> "Riduzione" And rsRIDUD.fields(i).Name <> "ParteFissa" And rsRIDUD.fields(i).Name <> "IdNumeriComponenti" Then
                MettivaloreInPKEY pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name, rsRIDUD.fields(i).value, rsRIDUD.fields(i).Type
            End If
        Next
        SiUD = True
    End If
    
    For Each Item In lvw9.ListItems
            If InStr(Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            
            If sKeyNode <> "-1" And SiUD Then
                'Utenze domestiche
                TariffaPF = GetValueInTag(TAG_TAR_PFUD_FLD, Item.ListSubItems(1).Tag)
                TariffaPV = GetValueInTag(TAG_TAR_PVUD_FLD, Item.ListSubItems(2).Tag)
                UD = Mid$(sKeyNode, 2)
                
                pKey = pKeyUD
                MettivaloreInPKEY pKey, SEPARATORE & "IdNumeriComponenti", UD, adNumeric
                rsRIDUD.Filter = adFilterNone
                rsRIDUD.Filter = GetPkeyForWhereCondition(pKey)

                Do While Not rsRIDUD.EOF
                    If rsRIDUD("ParteFissa") = True Then
                        TotMancatoIntroitoDAUD_PF = TotMancatoIntroitoDAUD_PF + TariffaPF * (rsRIDUD("Riduzione") / 100) * rsRIDUD("Valore")
                    Else
                        TotMancatoIntroitoDAUD_PV = TotMancatoIntroitoDAUD_PV + TariffaPV * (rsRIDUD("Riduzione") / 100) * rsRIDUD("Valore")
                    End If
                    rsRIDUD.MoveNext
                Loop
                            
            End If
            
            sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" And SiUA Then
                'Attivita
                TariffaPV = GetValueInTag(TAG_TAR_PVUA_FLD, Item.ListSubItems(7).Tag)
                TariffaPF = GetValueInTag(TAG_TAR_PFUA_FLD, Item.ListSubItems(6).Tag)
                UA = Mid$(sKeyNode, 2)
                
                
                pKey = pKeyUA
                MettivaloreInPKEY pKey, SEPARATORE & "CodAttivita", UA, adNumeric
                rsRIDUA.Filter = adFilterNone
                rsRIDUA.Filter = GetPkeyForWhereCondition(pKey)


                Do While Not rsRIDUA.EOF
                    If rsRIDUA("ParteFissa") = True Then
                        TotMancatoIntroitoDAUA_PF = TotMancatoIntroitoDAUA_PF + TariffaPF * (rsRIDUA("Riduzione") / 100) * rsRIDUA("Valore")
                    Else
                        TotMancatoIntroitoDAUA_PV = TotMancatoIntroitoDAUA_PV + TariffaPV * (rsRIDUA("Riduzione") / 100) * rsRIDUA("Valore")
                    End If
                    rsRIDUA.MoveNext
                Loop
            End If
    Next

    Command1(MOSTRARIDUZIONI).Caption = "Applica &Riduzioni"
    'carica nel LISTVIEW la perdita dalle riduzioni
    lvw9.ListItems(10).Text = "Introito da UA_PF"
    lvw9.ListItems(10).ForeColor = vbRed
    lvw9.ListItems(10).ListSubItems(1) = Format$(TotMancatoIntroitoDAUA_PF, FORMATVALUTA)
    lvw9.ListItems(10).ListSubItems(1).ForeColor = vbBlue
    
    lvw9.ListItems(11).Text = "Introito da UA_PV"
    lvw9.ListItems(11).ForeColor = vbRed
    lvw9.ListItems(11).ListSubItems(1) = Format$(TotMancatoIntroitoDAUA_PV, FORMATVALUTA)
    lvw9.ListItems(11).ListSubItems(1).ForeColor = vbBlue


    lvw9.ListItems(13).Text = "Introito da UD_PV"
    lvw9.ListItems(13).ForeColor = vbRed
    lvw9.ListItems(13).ListSubItems(1) = Format$(TotMancatoIntroitoDAUD_PV, FORMATVALUTA)
    lvw9.ListItems(13).ListSubItems(1).ForeColor = vbBlue
    lvw9.ListItems(14).Text = "Introito da UD_PF"
    lvw9.ListItems(14).ForeColor = vbRed
    lvw9.ListItems(14).ListSubItems(1) = Format$(TotMancatoIntroitoDAUD_PF, FORMATVALUTA)
    lvw9.ListItems(14).ListSubItems(1).ForeColor = vbBlue

    AllargaLeColonneLVW lvw9, False, 1, 1, 1, -1, -1
    
    
    If SiUA Then
        rsRIDUA.Filter = adFilterNone
        rsRIDUA.MoveFirst
    End If
    
    If SiUD Then
        rsRIDUD.Filter = adFilterNone
        rsRIDUD.MoveFirst
    End If
End Sub




'AS07042003
Sub MostraMancatoIntroitoConLeRiduzioni507OLD()
    'aggiorna le tariffe finali tendo conto delle riduzioni
    Dim Tariffa As Double
    Dim UD As Integer
    Dim UA As Integer
    Dim Item As ListItem
    Dim sKeyNode As String
    Dim pKeyUA As String
    Dim pKeyUD As String
    Dim pKey As String
    Dim i As Integer
    Dim SiUD As Boolean
    Dim SiUA As Boolean
    
    TotMancatoIntroito507 = 0
    
    pKeyUA = CreaLaPkey("BonusUtenzeAttivita")
    pKeyUD = CreaLaPkey("BonusUtenzeDomestiche")
    
    If Not rsRIDUA.EOF Then
        For i = 0 To rsRIDUA.fields.Count - 1
            If InStr(pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name) And rsRIDUA.fields(i).Name <> "Riduzione" And rsRIDUA.fields(i).Name <> "ParteFissa" And rsRIDUA.fields(i).Name <> "CodAttivita" Then
                MettivaloreInPKEY pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name, rsRIDUA.fields(i).value, rsRIDUA.fields(i).Type
            End If
            SiUA = True
        Next
    End If
    
    If Not rsRIDUD.EOF Then
        For i = 0 To rsRIDUD.fields.Count - 1
            If InStr(pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name) And rsRIDUD.fields(i).Name <> "Riduzione" And rsRIDUD.fields(i).Name <> "ParteFissa" And rsRIDUD.fields(i).Name <> "IdNumeriComponenti" Then
                MettivaloreInPKEY pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name, rsRIDUD.fields(i).value, rsRIDUD.fields(i).Type
            End If
        Next
        SiUD = True
    End If
    
    For Each Item In lvw10.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKeyNode <> "-1" And SiUD Then
                'Utenze domestiche
                Tariffa = CDbl(Item.SubItems(4))
                UD = Mid$(sKeyNode, 2)
                
                pKey = pKeyUD
                MettivaloreInPKEY pKey, SEPARATORE & "IdNumeriComponenti", UD, adNumeric
                rsRIDUD.Filter = adFilterNone
                rsRIDUD.Filter = GetPkeyForWhereCondition(pKey)

                Do While Not rsRIDUD.EOF
                    If rsRIDUD("ParteFissa") = True Then
                        TotMancatoIntroito507 = TotMancatoIntroito507 + Tariffa * (rsRIDUD("Riduzione") / 100) * rsRIDUD("Valore")
                    End If
                    rsRIDUD.MoveNext
                Loop
            
            End If
            
            sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" And SiUA Then
                'Attivita
                Tariffa = CDbl(Item.SubItems(1))
                UA = Mid$(sKeyNode, 2)
                
                
                pKey = pKeyUA
                MettivaloreInPKEY pKey, SEPARATORE & "CodAttivita", UA, adNumeric
                rsRIDUA.Filter = adFilterNone
                rsRIDUA.Filter = GetPkeyForWhereCondition(pKey)

                Do While Not rsRIDUA.EOF
                    If rsRIDUA("ParteFissa") = True Then
                        TotMancatoIntroito507 = TotMancatoIntroito507 + Tariffa * (rsRIDUA("Riduzione") / 100) * rsRIDUA("Valore")
                    End If
                    rsRIDUA.MoveNext
                Loop
            End If
    Next

    Command1(MOSTRARIDUZIONI).Caption = "Applica &Riduzioni"
    'carica nel LISTVIEW la perdita dalle riduzioni
    lvw10.ListItems(10).ListSubItems(3).Text = "- Introito"
    lvw10.ListItems(10).ListSubItems(3).ForeColor = vbRed
    lvw10.ListItems(10).ListSubItems(4) = Format$(TotMancatoIntroito507, FORMATVALUTA)
    lvw10.ListItems(10).ListSubItems(4).ForeColor = vbBlue
    



    AllargaLeColonneLVW lvw10, False, 1, 1, 1, -1, -1
    
End Sub


Sub RicalcolaTariffeConLeRiduzioni()
    'Dim AlfaUDPF As Double
    'Dim AlfaUDPV As Double
    'Dim AlfaUAPF As Double
    'Dim AlfaUAPV As Double
    Dim Item As ListItem
    Dim sKeyNode As String
    Dim TotMancatoIntroitoPrec As Double
    Dim Alfa As Double
    
    'AlfaUDPF = TotMancatoIntroitoDAUD_PF / GettitoCalcUDPF + 1
    'AlfaUDPV = TotMancatoIntroitoDAUD_PV / GettitoCalcUDPV + 1
    'AlfaUAPF = TotMancatoIntroitoDAUA_PF / GettitoCalcUAPF + 1
    'AlfaUAPV = TotMancatoIntroitoDAUA_PV / GettitoCalcUAPV + 1
    
    
    If TotMancatoIntroito > 0 Then
        
        TotMancatoIntroitoPrec = (GettitoCalcUDPF + GettitoCalcUDPV + GettitoCalcUAPF + GettitoCalcUAPV) - TotMancatoIntroito
        
        TotMancatoIntroito = (TotMancatoIntroitoDAUD_PF + TotMancatoIntroitoDAUD_PV + TotMancatoIntroitoDAUA_PF + TotMancatoIntroitoDAUA_PV) - TotMancatoIntroito
        
        Alfa = (TotMancatoIntroito / TotMancatoIntroitoPrec) + AlfaPrec
        
        GettitoCalcUDPFRidotto = 0
        GettitoCalcUDPVRidotto = 0
        GettitoCalcUAPFRidotto = 0
        GettitoCalcUAPVRidotto = 0
    Else
    
        Alfa = ((TotMancatoIntroitoDAUD_PF + TotMancatoIntroitoDAUD_PV + TotMancatoIntroitoDAUA_PF + TotMancatoIntroitoDAUA_PV) / _
         (GettitoCalcUDPF + GettitoCalcUDPV + GettitoCalcUAPF + GettitoCalcUAPV)) + 1
    End If
    
    Dim TariffaUDF As Double
    Dim TariffaUDV As Double
    Dim TariffaUAF As Double
    Dim TariffaUAV As Double
    
    For Each Item In lvw9.ListItems
    
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                TariffaUDV = GetValueInTag(TAG_TAR_PVUD_FLD, lvw5.ListItems(sKeyNode).ListSubItems(1).Tag)
                TariffaUDF = GetValueInTag(TAG_TAR_PFUD_FLD, lvw5.ListItems(sKeyNode).ListSubItems(idxColonnaSelezionata(keyColUDA) + 2).Tag)
                
                'Utenze domestiche PF
                TariffaUDF = TariffaUDF * Alfa
                Item.SubItems(1) = Format$(TariffaUDF, FORMAT6DEC)
                'Utenze domestiche PV
                TariffaUDV = TariffaUDV * Alfa
                Item.SubItems(2) = Format$(TariffaUDV, FORMAT6DEC)
                
                If TotMancatoIntroito > 0 Then
                    GettitoCalcUDPFRidotto = GettitoCalcUDPFRidotto + TariffaUDF * SupCorrettaConRiduzioneUD(Item.index - 1)
                    GettitoCalcUDPVRidotto = GettitoCalcUDPVRidotto + TariffaUDV * NumeroUtenzaDomesticaConRiduzione(Item.index - 1)
                End If
            
                Item.ListSubItems(1).Tag = SetValueInTag(TAG_TAR_PFUD_FLD, TariffaUDF, "")
                Item.ListSubItems(2).Tag = SetValueInTag(TAG_TAR_PVUD_FLD, TariffaUDV, "")
            
            End If

             
            sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                TariffaUAF = GetValueInTag(TAG_TAR_PFUA_FLD, lvw7.ListItems(sKeyNode).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
                TariffaUAV = GetValueInTag(TAG_TAR_PVUA_FLD, lvw7.ListItems(sKeyNode).ListSubItems(idxColonnaSelezionata(keyColUDA) + 1).Tag)
                
                'Utenze Attivita PF
                TariffaUAF = TariffaUAF * Alfa
                Item.SubItems(6) = Format$(TariffaUAF, FORMAT6DEC)
                
                'Utenze Attivita PV
                TariffaUAV = TariffaUAV * Alfa
                Item.SubItems(7) = Format$(TariffaUAV, FORMAT6DEC)
                'Utenze Attivita Tarffa
                
                Item.SubItems(5) = Format$(CDbl(TariffaUAF + TariffaUAV), FORMAT6DEC)
            
                If TotMancatoIntroito > 0 Then
                    GettitoCalcUAPFRidotto = GettitoCalcUAPFRidotto + TariffaUAF * SupConRiduzioneUAPF(Item.index - 1)
                    GettitoCalcUAPVRidotto = GettitoCalcUAPVRidotto + TariffaUAV * SupConRiduzioneUAPV(Item.index - 1)
                End If
                
                Item.ListSubItems(5).Tag = SetValueInTag(TAG_TAR_PFUA_FLD, CDbl(TariffaUAF + TariffaUAV), "")
                Item.ListSubItems(6).Tag = SetValueInTag(TAG_TAR_PFUA_FLD, TariffaUAF, "")
                Item.ListSubItems(7).Tag = SetValueInTag(TAG_TAR_PVUA_FLD, TariffaUAV, "")
            
            End If
    Next
    
    
    If TotMancatoIntroito > 0 Then
        'QUFUD = AlfaUDPF * QUFUD
        'CUD = AlfaUDPF * CUD
        'CUA = AlfaUAPV * CUA
        'QUAPUA = AlfaUAPF * QUAPUA
               
        'QUFUD = ((Alfa - 1) + AlfaPrec) * QUFUD
        'CUD = ((Alfa - 1) + AlfaPrec) * CUD
        'CUA = ((Alfa - 1) + AlfaPrec) * CUA
        'QUAPUA = ((Alfa - 1) + AlfaPrec) * QUAPUA
        
        QUFUD = ((Alfa)) * QUFUD
        CUD = ((Alfa)) * CUD
        CUA = ((Alfa)) * CUA
        QUAPUA = ((Alfa)) * QUAPUA
             
        Dim i As Integer
        For i = 0 To 14 + 7 - lvw9.ListItems.Count
            Set Item = lvw9.ListItems.Add(, "NoData" & CStr(lvw9.ListItems.Count))
            Item.ListSubItems.Add
        Next
        
        lvw9.ListItems(16).Text = "Parametro QUF"
        lvw9.ListItems(16).ForeColor = vbRed
        lvw9.ListItems(16).ListSubItems(1).Text = Round(QUFUD, 6)
        lvw9.ListItems(16).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(17).Text = "Parametro CUD"
        lvw9.ListItems(17).ListSubItems(1).Text = Round(CUD, 6)
        lvw9.ListItems(17).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(17).ForeColor = vbRed
        lvw9.ListItems(18).Text = "Parametro CUD*QUV"
        lvw9.ListItems(18).ListSubItems(1).Text = Round(CUD * QUV, 6)
        lvw9.ListItems(18).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(18).ForeColor = vbRed
        lvw9.ListItems(19).Text = "Parametro QUV"
        lvw9.ListItems(19).ListSubItems(1).Text = Round(QUV, 6)
        lvw9.ListItems(19).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(19).ForeColor = vbRed
        lvw9.ListItems(20).Text = "Parametro QUAPF"
        lvw9.ListItems(20).ListSubItems(1).Text = Round(QUAPUA, 6)
        lvw9.ListItems(20).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(20).ForeColor = vbRed
        lvw9.ListItems(21).Text = "Parametro CUA"
        lvw9.ListItems(21).ListSubItems(1).Text = Round(CUA, 6)
        lvw9.ListItems(21).ListSubItems(1).ForeColor = vbBlue
        lvw9.ListItems(21).ForeColor = vbRed
    
        If TotMancatoIntroito > 0 Then
            Txt(PAG7_GETTITOUD_ID) = InValutaEst(GettitoCalcUDPFRidotto + GettitoCalcUDPVRidotto, "")
            Txt(PAG7_GETTITOUA_ID) = InValutaEst(GettitoCalcUAPFRidotto + GettitoCalcUAPVRidotto, "")
        End If
    End If
    
    TotMancatoIntroito = (TotMancatoIntroitoDAUD_PF + TotMancatoIntroitoDAUD_PV + TotMancatoIntroitoDAUA_PF + TotMancatoIntroitoDAUA_PV)
    AlfaPrec = Alfa
    
    
    AllargaLeColonneLVW lvw9, False, 1, 1, 1, -1, -1
End Sub



Function Risultato507(valuta As String) As Boolean
    Dim Item As ListItem
    Dim rs507UA As ADODB.Recordset
    Dim rsUD As ADODB.Recordset
    Dim IQSUD As Double
    Dim IPSUD As Double
    Dim sTarsuUA As String
    Dim sTarsuUD As String
    Dim Alfa As Double
    Dim sDesAttivita As String
    Dim sDesComponenti As String
    
    Dim sKey As String
    Dim i As Integer
    Dim GettitiTotPerComponente() As Double
    
    On Error GoTo Risultato507ERR
    
    Risultato507 = False
    With lvw10
        
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "Utenza Attività"
        .ColumnHeaders.Add , , "Tari " & SimboloValuta(valuta) & "/mq anno", , lvwColumnRight
        .ColumnHeaders.Add , , "", 10
        .ColumnHeaders.Add , , "Utenza Domestica"
        .ColumnHeaders.Add , , "Tari " & SimboloValuta(valuta) & "/mq anno", , lvwColumnRight
        
        .ListItems.Clear
            
            
            
        'Coeff 507
        Set rs507UA = SelezionaCoeff507(CoeffUA507, mCodComune, mAnnoRiferimento)
        If rs507UA.EOF Then Info "Non Trovati Coefficienti 507 Utenza Attivita": Exit Function
        If rs507UA(0) = 0 Or rs507UA(1) = 0 Then Info "Non Trovati Coefficienti 507 Utenza Attivita": Exit Function
        rs507UA.MoveFirst
        Set rsUD = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento)
                
        'Calcolo
        'gettito finale ottenuto da UD e UA
        GettitoCalc507UD = 0
        GettitoCalc507UA = 0

        
        'utenza domestica
        IQSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD")
        IPSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD")
        sTarsuUD = CMG * IQSUD * IPSUD
        GettitoCalc507UD = (sTarsuUD * rsDatiTecniciComune("TotSuperficieUtenzaDomestica"))
        
        If Not TariffeUtenzeDomestiche507(GettitiTotPerComponente()) Then Exit Function
                
        For i = 1 To IIf(rs507UA.RecordCount > rsUD.RecordCount, rs507UA.RecordCount, rsUD.RecordCount)
            
            If Not rs507UA.EOF Then
                sKey = SetValueInTag(TAG_UA_FLD, "%" & rs507UA("CodAttivita"), "")
                sDesAttivita = rs507UA("Descrizione")
                sTarsuUA = CMG * rs507UA("IPS") * rs507UA("IQS")
                
                GettitoCalc507UA = GettitoCalc507UA + (sTarsuUA * rs507UA("Superficie"))
                rs507UA.MoveNext
            Else
                sKey = SetValueInTag(TAG_UA_FLD, -1, "")
                sTarsuUA = " "
                sDesAttivita = " "
            End If
        
               
            If Not rsUD.EOF Then
                If chkTarsu507perNComponenti.value = 1 Then
                    sDesComponenti = rsUD("Descrizione")
                    sTarsuUD = GettitiTotPerComponente(rsUD("IdNumeriComponenti") - 1) / SupCorrettaSenzaRiduzioneUD(rsUD("IdNumeriComponenti") - 1) 'rsUD("Superficie")
                    sKey = SetValueInTag(TAG_UD_FLD, "%" & rsUD("IdNumeriComponenti"), sKey)
                ElseIf rsUD("IdNumeriComponenti") = 1 Then
                        sDesComponenti = "Abitazione uso civile"
                        sKey = SetValueInTag(TAG_UD_FLD, "%" & rsUD("IdNumeriComponenti"), sKey)
                    Else
                        sKey = SetValueInTag(TAG_UD_FLD, -1, sKey)
                        sTarsuUD = " "
                        sDesComponenti = " "
                End If
                rsUD.MoveNext
            Else
                sKey = SetValueInTag(TAG_UD_FLD, -1, sKey)
                sTarsuUD = " "
                sDesComponenti = " "
            End If
                        
            
       
        
            Set Item = lvw10.ListItems.Add(, sKey, sDesAttivita)
            
            
            Item.ListSubItems.Add , "TariffaUA", Format$(sTarsuUA, FORMAT6DEC)
            Item.ListSubItems("TariffaUA").Tag = SetValueInTag(TAG_TAR_PFUA_FLD, sTarsuUA, "")
            Item.ListSubItems.Add , , " "
            
            
            Item.ListSubItems.Add , "UD", sDesComponenti
            Item.ListSubItems.Add , "TariffaUD", Format$(sTarsuUD, FORMAT6DEC)
            Item.ListSubItems("TariffaUD").Tag = SetValueInTag(TAG_TAR_PFUD_FLD, sTarsuUD, "")
        
'        For i = 1 To 20 - lvw10.ListItems.Count
'            Set Item = .ListItems.Add(, "NoData" & CStr(i), "")
 '           Item.ListSubItems.Add , , ""
  '          Item.ListSubItems.Add , , ""
  '          Item.ListSubItems.Add , , ""
   '         Item.ListSubItems.Add , , ""
    '    Next
        
        
        Next
        

        Dim tarsu As Double
        Alfa = (GettitoCalc507UD + GettitoCalc507UA) / CDbl(Txt(PAG1_COSTOCOPERTURA_ID))
        For i = 1 To .ListItems.Count
            tarsu = CDbl(.ListItems(i).SubItems(1)) / Alfa
            .ListItems(i).SubItems(1) = InValutaEst(tarsu, valuta, 4)
            On Error Resume Next
            tarsu = -1
            tarsu = CDbl(.ListItems(i).SubItems(4)) / Alfa
            If (tarsu <> -1) Then
                .ListItems(i).SubItems(4) = InValutaEst(tarsu, valuta, 4)
            End If
            On Error GoTo Risultato507ERR
        Next


        
        AllargaLeColonneLVW lvw10, False, 1, 1, 1, 1
        .Visible = True
    End With
    

    Set rs507UA = Nothing
    Set rsUD = Nothing
    
    If Command1(MOSTRARIDUZIONI).Visible Then Command1_Click (MOSTRARIDUZIONI)
    
    Risultato507 = True
    
    Exit Function
Risultato507ERR:
    MsgBox Error$, vbCritical, App.Title
End Function



Function Risultato507NEW(valuta As String) As Boolean
    Dim Item As ListItem
    Dim rs507UA As ADODB.Recordset
    Dim rs507UD As ADODB.Recordset
    Dim rsUD As ADODB.Recordset
    Dim rsUA As ADODB.Recordset
    Dim IQSUD As Double
    Dim IPSUD As Double
    Dim sTarsuUA As String
    Dim sTarsuUD As String
    Dim sDesAttivita As String
    Dim sDesComponenti As String
    
    Dim sKey As String
    Dim i As Integer
    Dim gettitoPrimario As Double
    
    
    
    On Error GoTo Risultato507ERR
    
    Risultato507NEW = False
    With lvw10
        
        .Visible = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "Utenza Attività"
        .ColumnHeaders.Add , , "Tari " & SimboloValuta(valuta) & "/mq anno", , lvwColumnRight
        .ColumnHeaders.Add , , "", 10
        .ColumnHeaders.Add , , "Utenza Domestica"
        .ColumnHeaders.Add , , "Tari " & SimboloValuta(valuta) & "/mq anno", , lvwColumnRight
        
        .ListItems.Clear
            
        'Coeff 507 UA
        Set rs507UA = SelezionaCoeff507(CoeffUA507, mCodComune, mAnnoRiferimento)
        If rs507UA.EOF Then Info "Non Trovati Coefficienti 507 Utenza Attivita": Exit Function
        If rs507UA(0) = 0 Or rs507UA(1) = 0 Then Info "Non Trovati Coefficienti 507 Utenza Attivita": Exit Function
        rs507UA.MoveFirst
        
        Set rs507UD = SelezionaCoeff507(CoeffUD507, mCodComune, mAnnoRiferimento)
        If rs507UD.EOF Then Info "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
        If rs507UD(0) = 0 Or rs507UD(1) = 0 Then Info "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
        rs507UD.MoveFirst
        
        Set rsUA = SelezionaUtenzeAttivita(mCodComune, mAnnoRiferimento)
        Set rsUD = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento)
                
        'Calcolo
        'gettito finale ottenuto da UD e UA
        GettitoCalc507UD = 0
        GettitoCalc507UA = 0

        Dim ProduzioneRifiuti As Double
        'utenza domestica
        IQSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD")
        IPSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD")
        Dim TotSupRidotta As Double
        
        
        ProduzioneRifiuti = 0
        
        
        For i = 1 To IIf(rs507UA.RecordCount > rsUD.RecordCount, rs507UA.RecordCount, rsUD.RecordCount)
            
            If Not rs507UA.EOF Then
                
                rsUA.Filter = adFilterNone
                rsUA.Filter = "IdCategoriaAttivita=" & rs507UA("CodAttivita")
                
                TotaliDaRiduzioniUA rs507UA("CodAttivita"), rsRIDUA, , TotSupRidotta
                ProduzioneRifiuti = ProduzioneRifiuti + (KM * rs507UA("IPS") * (rsUA("Superficie") - TotSupRidotta))
                
                rs507UA.MoveNext
            End If
        
               
            If Not rsUD.EOF Then
                If chkTarsu507perNComponenti.value = 1 Then
                    rs507UD.Filter = adFilterNone
                    rs507UD.Filter = "CodUtenza=" & rsUD("IdNumeriComponenti")
                
                    TotaliDaRiduzioniUD rsUD("IdNumeriComponenti"), rsRIDUD, , TotSupRidotta
                    ProduzioneRifiuti = ProduzioneRifiuti + (KM * rs507UD("Valore") * (rsUD("Superficie") - TotSupRidotta))
                    
                Else
                    TotaliDaRiduzioniUD rsUD("IdNumeriComponenti"), rsRIDUD, , TotSupRidotta
                    ProduzioneRifiuti = ProduzioneRifiuti + (KM * IPSUD * (rsUD("Superficie") - TotSupRidotta))
                End If
                
                rsUD.MoveNext
            End If
        Next
        
        
        
        
        CoeffCorrettivo507IPS = rsDatiTecniciComune("QRifiutiTotali") / ProduzioneRifiuti
        
        On Error Resume Next
        rs507UA.MoveFirst
        rsUD.MoveFirst
        On Error GoTo Risultato507ERR
        
        'ciclo per calcolo Gettito primario
        For i = 1 To IIf(rs507UA.RecordCount > rsUD.RecordCount, rs507UA.RecordCount, rsUD.RecordCount)
            
            If Not rs507UA.EOF Then
                sKey = SetValueInTag(TAG_UA_FLD, "%" & rs507UA("CodAttivita"), "")
                sDesAttivita = rs507UA("Descrizione")
                
                rsUA.Filter = adFilterNone
                rsUA.Filter = "IdCategoriaAttivita=" & rs507UA("CodAttivita")
                
                TotaliDaRiduzioniUA rs507UA("CodAttivita"), rsRIDUA, , TotSupRidotta
                gettitoPrimario = gettitoPrimario + (CMG * rs507UA("IQS") * (rs507UA("IPS") * CoeffCorrettivo507IPS) * (rsUA("Superficie") - TotSupRidotta))
                
                sTarsuUA = rs507UA("IPS") * CoeffCorrettivo507IPS * rs507UA("IQS") * CMG
                
                rs507UA.MoveNext
            Else
                sKey = SetValueInTag(TAG_UA_FLD, -1, "")
                sTarsuUA = " "
                sDesAttivita = " "
            End If
        
               
            If Not rsUD.EOF Then
                If chkTarsu507perNComponenti.value = 1 Then
                    sDesComponenti = rsUD("Descrizione")
                    sKey = SetValueInTag(TAG_UD_FLD, "%" & rsUD("IdNumeriComponenti"), sKey)

                    rs507UD.Filter = adFilterNone
                    rs507UD.Filter = "CodUtenza=" & rsUD("IdNumeriComponenti")
                    
                    sTarsuUD = CMG * IQSUD * (rs507UD("Valore") * CoeffCorrettivo507IPS)
                    
                    TotaliDaRiduzioniUD rsUD("IdNumeriComponenti"), rsRIDUD, , TotSupRidotta
                    gettitoPrimario = gettitoPrimario + (CMG * IQSUD * (rs507UD("Valore") * CoeffCorrettivo507IPS) * (rsUD("Superficie") - TotSupRidotta))
                    
                Else
                    If rsUD("IdNumeriComponenti") = 1 Then
                        sDesComponenti = "Abitazione uso civile"
                        sKey = SetValueInTag(TAG_UD_FLD, "%" & rsUD("IdNumeriComponenti"), sKey)
                        sTarsuUD = CMG * IQSUD * (IPSUD * CoeffCorrettivo507IPS)
                          
                    Else
                        sKey = SetValueInTag(TAG_UD_FLD, -1, sKey)
                        sTarsuUD = " "
                        sDesComponenti = " "
                    End If
                    TotaliDaRiduzioniUD rsUD("IdNumeriComponenti"), rsRIDUD, , TotSupRidotta
                    gettitoPrimario = gettitoPrimario + (CMG * IQSUD * (IPSUD * CoeffCorrettivo507IPS) * (rsUD("Superficie") - TotSupRidotta))
                End If
                
                rsUD.MoveNext
            Else
                sKey = SetValueInTag(TAG_UD_FLD, -1, sKey)
                sTarsuUD = " "
                sDesComponenti = " "
            End If
                        
            
            Set Item = lvw10.ListItems.Add(, sKey, sDesAttivita)
            
            
            Item.ListSubItems.Add , "TariffaUA", Format$(sTarsuUA, FORMAT6DEC)
            Item.ListSubItems("TariffaUA").Tag = SetValueInTag(TAG_TAR_PFUA_FLD, sTarsuUA, "")
            Item.ListSubItems.Add , , " "
            
            
            Item.ListSubItems.Add , "UD", sDesComponenti
            Item.ListSubItems.Add , "TariffaUD", Format$(sTarsuUD, FORMAT6DEC)
            Item.ListSubItems("TariffaUD").Tag = SetValueInTag(TAG_TAR_PFUD_FLD, sTarsuUD, "")
    
        Next
        
        CoeffCorrTariffa507 = gettitoPrimario / CalcoloCostoCopertura
      

        Dim tarsu As Double
        For i = 1 To .ListItems.Count
            tarsu = CDbl(.ListItems(i).SubItems(1)) / CoeffCorrTariffa507
            .ListItems(i).SubItems(1) = InValutaEst(tarsu, valuta, 4)
            On Error Resume Next
            tarsu = -1
            tarsu = CDbl(.ListItems(i).SubItems(4)) / CoeffCorrTariffa507
            If (tarsu <> -1) Then
                .ListItems(i).SubItems(4) = InValutaEst(tarsu, valuta, 4)
            End If
            On Error GoTo Risultato507ERR
        Next

        'mostra dati finali calcolo tariffa
        .ListItems(8).ListSubItems(3).ForeColor = vbRed
        .ListItems(8).ListSubItems(3).Text = "cmg finale (€/mq anno)"
        .ListItems(8).ListSubItems(4).ForeColor = vbBlue
        .ListItems(8).ListSubItems(4).Text = Format$(CMG / CoeffCorrTariffa507, FORMAT4DEC)
        
        .ListItems(9).ListSubItems(3).ForeColor = vbRed
        .ListItems(9).ListSubItems(3).Text = "produz. media finale (kg/mq anno)"
        .ListItems(9).ListSubItems(4).ForeColor = vbBlue
        .ListItems(9).ListSubItems(4).Text = Format$(KM, FORMAT4DEC)
        
        
        
        AllargaLeColonneLVW lvw10, False, 1, 1, 1, 1
        
        
        
        .Visible = True
    End With
    

    

    Set rs507UA = Nothing
    Set rsUD = Nothing
    
    If Command1(MOSTRARIDUZIONI).Visible Then Command1_Click (MOSTRARIDUZIONI)
    
    
    Risultato507NEW = True
    
    Exit Function
Risultato507ERR:
    MsgBox Error$, vbCritical, App.Title
End Function






Function TariffeUtenzeDomestiche507(GettitoTotPerComponente() As Double) As Boolean
    Dim rs507UD As ADODB.Recordset
    Dim Gettito As Double
    Dim i As Integer
    Dim ii As Integer
    Dim NComFamiglia As Double
    Dim Alfa As Double

    '
    Erase GettitoTotPerComponente
    
    Erase SupCorrettaConRiduzioneUD
    Erase SupCorrettaSenzaRiduzioneUD
    SuperficeCorrettaUD mCodComune, mAnnoRiferimento, SupCorrettaConRiduzioneUD, SupCorrettaSenzaRiduzioneUD, rsRIDUD, False
    
    TariffeUtenzeDomestiche507 = False
    'dati anagrafici comune
    rsDatiAnagraficiComune.MoveFirst
    
    'coeff utenze domestiche 507
    Set rs507UD = SelezionaCoeff507(CoeffUD507, mCodComune, mAnnoRiferimento)
    If rs507UD.EOF Then MsgBox "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
    If rs507UD(0) = 0 Then MsgBox "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
    
    

    rs507UD.MoveFirst
    Do While Not rs507UD.EOF
        'Numero Componenti
        Select Case rs507UD("CodUtenza")
            Case 1
                NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs507UD("CodUtenza")) & "Componente")
            Case Is >= 2
                If rs507UD("CodUtenza") = 6 Then
                    NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                Else
                    NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs507UD("CodUtenza")) & "Componenti")
                End If
        End Select
    
        'seconda casa
        If rsDatiTecniciComune("Id_NumCompSecondaCasa") = rs507UD("CodUtenza") Then
            NComFamiglia = NComFamiglia + rsDatiTecniciComune("NRuoliSecondaCasa")
        End If
        
        
        ii = -1
        On Error Resume Next
        ii = UBound(GettitoTotPerComponente)
        On Error GoTo 0
        ReDim Preserve GettitoTotPerComponente(ii + 1)
        GettitoTotPerComponente(ii + 1) = (GettitoCalc507UD / rsDatiAnagraficiComune("NumAbitantiResidenti")) * rs507UD("valore") * NComFamiglia
        
        Gettito = Gettito + GettitoTotPerComponente(ii + 1)
        
        rs507UD.MoveNext
    Loop
    
    Set rs507UD = Nothing
    
    Alfa = GettitoCalc507UD / Gettito

    For i = LBound(GettitoTotPerComponente) To UBound(GettitoTotPerComponente)
        GettitoTotPerComponente(i) = GettitoTotPerComponente(i) * Alfa
    Next
    TariffeUtenzeDomestiche507 = True
    Exit Function
TariffeUtenzeDomestiche507ERR:
    TariffeUtenzeDomestiche507 = False
End Function

Function TariffeUtenzeDomestiche507OLD1(GettitoTotPerComponente() As Double) As Boolean
    Dim rs507UD As ADODB.Recordset
    Dim ProduzioneRifiuti As Double
    Dim CostoChilo  As Double
    Dim CostoTotaleRifiuti  As Double
    Dim i As Integer
    Dim ii As Integer
    Dim NComFamiglia As Double
    Dim Alfa As Double
    
    On Error GoTo TariffeUtenzeDomestiche507ERR
    TariffeUtenzeDomestiche507OLD1 = False
    '
    Erase GettitoTotPerComponente
    
    'dati anagrafici comune
    rsDatiAnagraficiComune.MoveFirst
    
    'coeff utenze domestiche 507
    Set rs507UD = SelezionaCoeff507(CoeffUD507, mCodComune, mAnnoRiferimento)
    If rs507UD.EOF Then Info "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
    If rs507UD(0) = 0 Then Info "Non Trovati Coefficienti 507 Utenza Domestica": Exit Function
    
    
    Erase SupCorrettaConRiduzioneUD
    Erase SupCorrettaSenzaRiduzioneUD
    SuperficeCorrettaUD mCodComune, mAnnoRiferimento, SupCorrettaConRiduzioneUD, SupCorrettaSenzaRiduzioneUD, rsRIDUD, False
    
    ProduzioneRifiuti = 0
    CostoChilo = 0
    CostoTotaleRifiuti = 0
    
    If chkTarsu507perNComponenti.value = 0 Then
        TariffeUtenzeDomestiche507OLD1 = True
        Exit Function
    End If
    
    
    For i = 1 To 3
        rs507UD.MoveFirst
        Do While Not rs507UD.EOF
            'Numero Componenti
            Select Case rs507UD("CodUtenza")
                Case 1
                    NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs507UD("CodUtenza")) & "Componente")
                Case Is >= 2
                    If rs507UD("CodUtenza") = 6 Then
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaConPiuDi5Componenti")
                    Else
                        NComFamiglia = rsDatiAnagraficiComune("FamigliaCon" & CStr(rs507UD("CodUtenza")) & "Componenti")
                    End If
            End Select
        
            'seconda casa
            If rsDatiTecniciComune("Id_NumCompSecondaCasa") = rs507UD("CodUtenza") Then
                NComFamiglia = NComFamiglia + rsDatiTecniciComune("NRuoliSecondaCasa")
            End If
            
            
            Select Case i
                Case 1
                    ProduzioneRifiuti = ProduzioneRifiuti + rs507UD("valore") * NComFamiglia * 360
                    rs507UD.MoveNext
                    If rs507UD.EOF Then
                        CostoChilo = GettitoCalc507UD / ProduzioneRifiuti
                    End If
                Case 2
                    CostoTotaleRifiuti = CostoTotaleRifiuti + rs507UD("valore").value * NComFamiglia * CostoChilo * 360
                    rs507UD.MoveNext
                    If rs507UD.EOF Then
                        Alfa = GettitoCalc507UD / CostoTotaleRifiuti
                        CostoTotaleRifiuti = 0
                    End If
                Case 3
                    ii = -1
                    On Error Resume Next
                    ii = UBound(GettitoTotPerComponente)
                    On Error GoTo 0
                    ReDim Preserve GettitoTotPerComponente(ii + 1)
                    GettitoTotPerComponente(ii + 1) = rs507UD("valore").value * NComFamiglia * CostoChilo * 360 * Alfa
                    CostoTotaleRifiuti = CostoTotaleRifiuti + GettitoTotPerComponente(ii + 1)
                    rs507UD.MoveNext
            End Select
        Loop
    Next
    
    TariffeUtenzeDomestiche507OLD1 = True
    Set rs507UD = Nothing
    Exit Function
TariffeUtenzeDomestiche507ERR:
    Err.Raise 514, "TariffeUtenzeDomestiche507", "Attenzione verificare che vi siano i coefficienti per UD, i valori ammessi sono > 0" + vbCrLf + "Errore:" + Err.description
End Function


'AS07042003
Sub RicalcolaTariffeConLeRiduzioni507OLD()
    Dim Alfa As Double
    Dim Item As ListItem
    Dim sKeyNode As String
    
    Alfa = TotMancatoIntroito507 / CDbl(Txt(PAG1_COSTOCOPERTURA_ID)) + 1
    
    For Each Item In lvw10.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                'Utenze domestiche PF
                Item.SubItems(4) = Format$(CDbl(Item.SubItems(4)) * Alfa, FORMATVALUTA)
            End If
             
             sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                'Utenze Attivita PF
                Item.SubItems(1) = Format$(CDbl(Item.SubItems(1)) * Alfa, FORMATVALUTA)
            End If
    Next
    
    'cmg corretto
    CMG = CMG * Alfa
    
    AllargaLeColonneLVW lvw9, False, 1, 1, 1, -1, -1
End Sub


Sub RicalcolaTariffeConLeRiduzioni507()
    Dim pKeyUA As String
    Dim pKeyUD As String
    Dim i As Integer
    Dim SiUA As Boolean
    Dim SiUD As Boolean
    Dim Item As ListItem
    Dim sKeyNode As String
    Dim IQSUD As Double
    Dim IPSUD As Double
    Dim Tariffa As Double
    Dim UD As String
    Dim UA As String
    Dim IntroitoUD As Double
    Dim IntroitoUA As Double
    Dim pKey As String
    Dim rsUD As ADODB.Recordset
    Dim rsUA As ADODB.Recordset
    Dim SuperficieUD As Double
    Dim SuperficieUDRid As Double
    Dim SuperficieUARid As Double
    Dim GettitoUD As Double
    Dim GettitoUA As Double
    Dim SupericeUA() As Double
    Dim Alfa  As Double
    
    Set rsUD = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento)
    SuperficieUD = 0
    Do While Not rsUD.EOF
        SuperficieUD = SuperficieUD + rsUD("Superficie").value
        rsUD.MoveNext
    Loop
    For i = 0 To UBound(SupCorrettaSenzaRiduzioneUD)
           SuperficieUD = SuperficieUD + SupCorrettaSenzaRiduzioneUD(i)
    Next
   
   
    Erase SupericeUA
    Set rsUA = SelezionaUtenzeAttivita(mCodComune, mAnnoRiferimento)
    Do While Not rsUA.EOF
        If rsUA("Superficie") > 0 Then
            i = -1
            On Error Resume Next
            i = UBound(SupericeUA)
            On Error GoTo 0
            ReDim Preserve SupericeUA(i + 1)
            
            SupericeUA(i + 1) = rsUA("Superficie")
        End If
        rsUA.MoveNext
    Loop
    
    pKeyUA = CreaLaPkey("BonusUtenzeAttivita")
    pKeyUD = CreaLaPkey("BonusUtenzeDomestiche")
    

    
    If Not rsRIDUA.EOF Then
        For i = 0 To rsRIDUA.fields.Count - 1
            If InStr(pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name) And rsRIDUA.fields(i).Name <> "Riduzione" And rsRIDUA.fields(i).Name <> "ParteFissa" And rsRIDUA.fields(i).Name <> "CodAttivita" Then
                MettivaloreInPKEY pKeyUA, SEPARATORE & rsRIDUA.fields(i).Name, rsRIDUA.fields(i).value, rsRIDUA.fields(i).Type
            End If
            SiUA = True
        Next
    End If
    
    If Not rsRIDUD.EOF Then
        For i = 0 To rsRIDUD.fields.Count - 1
            If InStr(pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name) And rsRIDUD.fields(i).Name <> "Riduzione" And rsRIDUD.fields(i).Name <> "ParteFissa" And rsRIDUD.fields(i).Name <> "IdNumeriComponenti" And rsRIDUD.fields(i).Name <> "indSingle" Then
                MettivaloreInPKEY pKeyUD, SEPARATORE & rsRIDUD.fields(i).Name, rsRIDUD.fields(i).value, rsRIDUD.fields(i).Type
            End If
        Next
        SiUD = True
    End If
    
    GettitoUA = 0
    GettitoUD = 0
    For Each Item In lvw10.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKeyNode <> "-1" And SiUD Then
                'Utenze domestiche
                IQSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD")
                IPSUD = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD")
                Tariffa = CMG * IQSUD * IPSUD
                UD = Mid$(sKeyNode, 2)
                
                pKey = pKeyUD
                MettivaloreInPKEY pKey, SEPARATORE & "IdNumeriComponenti", UD, adNumeric
                rsRIDUD.Filter = adFilterNone
                rsRIDUD.Filter = GetPkeyForWhereCondition(pKey)
                
                SuperficieUDRid = 0
                IntroitoUD = 0
                Do While Not rsRIDUD.EOF
                    If rsRIDUD("ParteFissa") = True And _
                        ((chkTarsu507perNComponenti = 1 And rsRIDUD("indSingle").value = False) Or _
                        (chkTarsu507perNComponenti = 0)) Then
                        
                        SuperficieUDRid = SuperficieUDRid + rsRIDUD("Valore") * (rsRIDUD("Riduzione") / 100)
                        IntroitoUD = IntroitoUD + Tariffa * (rsRIDUD("Riduzione") / 100) * rsRIDUD("Valore")
                    End If
                    
                    rsRIDUD.MoveNext
                Loop
                
                GettitoUD = Tariffa * (SuperficieUD - SuperficieUDRid) '+ IntroitoUD
                SiUD = True
            End If
            
            sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" And SiUA Then
                'Attivita
                Tariffa = CDbl(GetValueInTag(TAG_TAR_PFUA_FLD, Item.ListSubItems("TariffaUA").Tag))
                UA = Mid$(sKeyNode, 2)
                
                pKey = pKeyUA
                MettivaloreInPKEY pKey, SEPARATORE & "CodAttivita", UA, adNumeric
                rsRIDUA.Filter = adFilterNone
                rsRIDUA.Filter = GetPkeyForWhereCondition(pKey)

                SuperficieUARid = 0
                IntroitoUA = 0
                Do While Not rsRIDUA.EOF
                    If rsRIDUA("ParteFissa") = True Then
                        SuperficieUARid = SuperficieUARid + rsRIDUA("Valore") * (rsRIDUA("Riduzione") / 100)
                        IntroitoUA = IntroitoUA + Tariffa * (rsRIDUA("Riduzione") / 100) * rsRIDUA("Valore")
                    End If
                    
                    rsRIDUA.MoveNext
                Loop
                
                
                If UBound(SupericeUA) < Item.index - 1 Then
                    Err.Raise 5000, "RicalcolaTariffeColLeRiduzioni507", "Attenzione coefficienti utenza Attività non bilanciati con superfici"
                End If
                
                If SuperficieUARid > 0 Then
                    GettitoUA = GettitoUA + Tariffa * (SupericeUA(Item.index - 1) - SuperficieUARid) '+ IntroitoUA
                Else
                    GettitoUA = GettitoUA + Tariffa * (SupericeUA(Item.index - 1))
                End If
            End If
    Next

    'alfa GettitoUD+GettitoUA / GettitoTot
    'divido tutte le tariffr per alfa
    Alfa = (GettitoUD + GettitoUA) / CalcoloCostoCopertura
    For Each Item In lvw10.ListItems
            If InStr(1, Item.Key, "NoData") > 0 Then Exit For
            sKeyNode = GetValueInTag(TAG_UD_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                'Utenze domestiche
                Tariffa = CDbl(GetValueInTag(TAG_TAR_PFUD_FLD, Item.ListSubItems("TariffaUD").Tag)) / Alfa
                Item.ListSubItems("TariffaUD").Tag = SetValueInTag(TAG_TAR_PFUD_FLD, Tariffa, "")
                Item.ListSubItems("TariffaUD") = Format$(Tariffa, FORMAT6DEC)
            End If
            
            sKeyNode = GetValueInTag(TAG_UA_FLD, Item.Key)
            If sKeyNode <> "-1" Then
                'Attivita
                Tariffa = CDbl(GetValueInTag(TAG_TAR_PFUA_FLD, Item.ListSubItems("TariffaUA").Tag)) / Alfa
                Item.ListSubItems("TariffaUA").Tag = SetValueInTag(TAG_TAR_PFUA_FLD, Tariffa, "")
                Item.ListSubItems("TariffaUA") = Format$(Tariffa, FORMAT6DEC)
            End If
    Next


    
    'applica coefficiente corretto
    AllargaLeColonneLVW lvw10, False, 1, 1, 1, -1, -1
End Sub

Sub MettiInValuta(lvw As ListView, valuta As String, Optional nDec As Integer = 2)
    Dim i As Long
    Dim ii As Long
    Dim Item As ListItem
    Dim Valore  As String
    
    Screen.MousePointer = vbHourglass
    lvw9.Visible = False
    For i = 1 To lvw.ListItems.Count
        Set Item = lvw.ListItems(i)
        For ii = 1 To Item.ListSubItems.Count
            'cerco nel tag il valore non approssimato
            Valore = GetValueInTag(TAG_TAR_PFUD_FLD, Item.ListSubItems(ii).Tag)
            If Not IsNumeric(Valore) Then
                Valore = GetValueInTag(TAG_TAR_PVUD_FLD, Item.ListSubItems(ii).Tag)
                If Not IsNumeric(Valore) Then
                    Valore = GetValueInTag(TAG_TAR_PVUA_FLD, Item.ListSubItems(ii).Tag)
                    If Not IsNumeric(Valore) Then
                        Valore = GetValueInTag(TAG_TAR_PFUA_FLD, Item.ListSubItems(ii).Tag)
                    End If
                End If
            End If
            If IsNumeric(Valore) Then
                If valuta = gValutaDefault Then
                    Item.SubItems(ii) = Format$(Valore, FORMAT6DEC)
                Else
                    Item.SubItems(ii) = InValutaEst(Valore, valuta, nDec)
                End If
            End If
        Next
    Next
    AllargaLeColonneLVW lvw9, False, 1, 1, 1, -1, -1
    lvw9.Visible = True
    
    Screen.MousePointer = vbDefault
    
End Sub

