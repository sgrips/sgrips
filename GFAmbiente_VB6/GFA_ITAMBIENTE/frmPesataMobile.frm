VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesataMobile 
   Caption         =   "Gestione Pesate"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainer 
      Height          =   6615
      Index           =   0
      Left            =   120
      ScaleHeight     =   6555
      ScaleWidth      =   14955
      TabIndex        =   18
      Top             =   480
      Width           =   15015
      Begin VB.Frame Frame1 
         Caption         =   "Trasportatore"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   58
         Top             =   3360
         Width           =   8895
         Begin VB.ComboBox cmbTarga 
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton cmbRicerca 
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
            Height          =   480
            Index           =   1
            Left            =   1440
            Picture         =   "frmPesataMobile.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   420
            Width           =   495
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
            Index           =   3
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Codice del trasportatore che ha portato il rifiuto"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   2160
            TabIndex        =   9
            ToolTipText     =   "Inserire il numero formulario di ingresso"
            Top             =   1080
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpickFormulario 
            Height          =   450
            Left            =   4560
            TabIndex        =   59
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   794
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   54788097
            CurrentDate     =   39352
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   33
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   480
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
            Index           =   32
            Left            =   4200
            TabIndex        =   68
            Top             =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   31
            Left            =   4200
            TabIndex        =   67
            Top             =   240
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
            Index           =   30
            Left            =   5400
            TabIndex        =   66
            Top             =   465
            Width           =   675
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   29
            Left            =   5400
            TabIndex        =   65
            Top             =   240
            Width           =   825
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
            Index           =   28
            Left            =   2160
            TabIndex        =   64
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
            Index           =   27
            Left            =   2160
            TabIndex        =   63
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice"
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
            Index           =   26
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   585
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   17
            Left            =   4560
            TabIndex        =   61
            Top             =   840
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   24
            Left            =   2160
            TabIndex        =   60
            Top             =   840
            Width           =   1350
         End
      End
      Begin VB.Frame fmeUtente 
         Caption         =   "Utente"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   8895
         Begin VB.ComboBox cmbTarga 
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
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Text            =   "Combo3"
            Top             =   1680
            Width           =   2055
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
            Index           =   0
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
            Height          =   480
            Index           =   0
            Left            =   1440
            Picture         =   "frmPesataMobile.frx":076A
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   420
            Width           =   495
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
            Index           =   2
            Left            =   4680
            TabIndex        =   102
            Top             =   465
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comune"
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
            Index           =   1
            Left            =   4680
            TabIndex        =   101
            Top             =   240
            Width           =   720
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   25
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   480
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
            Index           =   23
            Left            =   2160
            TabIndex        =   55
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Progressivo"
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
            Index           =   22
            Left            =   2160
            TabIndex        =   54
            Top             =   840
            Width           =   990
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
            Index           =   21
            Left            =   6240
            TabIndex        =   53
            Top             =   1665
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fino al"
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
            Left            =   6240
            TabIndex        =   52
            Top             =   1440
            Width           =   555
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
            Index           =   19
            Left            =   3840
            TabIndex        =   51
            Top             =   1665
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valido dal"
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
            Left            =   3840
            TabIndex        =   50
            Top             =   1440
            Width           =   810
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
            Index           =   16
            Left            =   6240
            TabIndex        =   45
            Top             =   240
            Width           =   660
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
            Index           =   15
            Left            =   6240
            TabIndex        =   44
            Top             =   465
            Width           =   855
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   14
            Left            =   6240
            TabIndex        =   43
            Top             =   840
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
            Index           =   13
            Left            =   6240
            TabIndex        =   42
            Top             =   1065
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   12
            Left            =   3840
            TabIndex        =   41
            Top             =   840
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
            Index           =   11
            Left            =   3840
            TabIndex        =   40
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Badge"
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
            Index           =   10
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   540
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
            Index           =   9
            Left            =   120
            TabIndex        =   38
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codice"
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
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   585
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
            TabIndex        =   36
            Top             =   240
            Width           =   1365
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
            TabIndex        =   35
            Top             =   465
            Width           =   675
         End
      End
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
         Height          =   450
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin MSComctlLib.ImageCombo imgCombo1 
         Height          =   450
         Left            =   4200
         TabIndex        =   5
         Top             =   2760
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   794
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList1"
      End
      Begin VB.Frame fmeTotali 
         Caption         =   "Totalizzazioni"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   9240
         TabIndex        =   28
         Top             =   1200
         Width           =   2535
         Begin MSComctlLib.ListView lvwTotMovimenti 
            Height          =   615
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1085
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
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
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Tot.Pesate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Tot.Bonus"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Max Qty Giornaliera"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Max Qty Annua"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Peso rilevato"
         Top             =   2760
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvwImmobili 
         Height          =   495
         Left            =   9240
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- FATTURABILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   35
         Left            =   5640
         TabIndex        =   89
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Immobili"
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
         Index           =   6
         Left            =   9240
         TabIndex        =   26
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ricerca Tip.Rifiuto"
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
         Index           =   8
         Left            =   2280
         TabIndex        =   33
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipologia Rifiuto"
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
         Index           =   7
         Left            =   4200
         TabIndex        =   27
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantità"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   690
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Salva"
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
      Index           =   7
      Left            =   9960
      TabIndex        =   99
      Top             =   10320
      Width           =   1095
   End
   Begin VB.PictureBox picContainer 
      Height          =   4215
      Index           =   1
      Left            =   600
      ScaleHeight     =   4155
      ScaleWidth      =   13995
      TabIndex        =   24
      Top             =   3240
      Width           =   14055
      Begin VB.ComboBox cmbRicercaMov 
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
         Index           =   5
         ItemData        =   "frmPesataMobile.frx":0ED4
         Left            =   10560
         List            =   "frmPesataMobile.frx":0EDC
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbRicercaMov 
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
         Index           =   4
         ItemData        =   "frmPesataMobile.frx":0EE7
         Left            =   8400
         List            =   "frmPesataMobile.frx":0EF8
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbRicercaMov 
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
         Index           =   3
         ItemData        =   "frmPesataMobile.frx":0F32
         Left            =   2280
         List            =   "frmPesataMobile.frx":0F40
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox cmbRicercaMov 
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
         Index           =   2
         ItemData        =   "frmPesataMobile.frx":0F6D
         Left            =   6480
         List            =   "frmPesataMobile.frx":0F75
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox chkEscludiCancellati 
         Caption         =   "Escludi Cancellati"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13320
         TabIndex        =   84
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   76
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdRicercaUtenteMov 
         Height          =   480
         Index           =   1
         Left            =   1560
         Picture         =   "frmPesataMobile.frx":0F80
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   960
         Width           =   480
      End
      Begin VB.ComboBox cmbTargaMov 
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
         Index           =   1
         Left            =   2280
         TabIndex        =   74
         Text            =   "Combo3"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cmbTargaMov 
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
         Left            =   2280
         TabIndex        =   71
         Text            =   "Combo3"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cmbRicercaMov 
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
         Index           =   1
         ItemData        =   "frmPesataMobile.frx":16EA
         Left            =   120
         List            =   "frmPesataMobile.frx":16F8
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cmbRicercaMov 
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
         ItemData        =   "frmPesataMobile.frx":171B
         Left            =   6240
         List            =   "frmPesataMobile.frx":172C
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxtMov 
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
         Index           =   1
         Left            =   6240
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdRicercaUtenteMov 
         Height          =   480
         Index           =   0
         Left            =   1560
         Picture         =   "frmPesataMobile.frx":1766
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   480
      End
      Begin VB.TextBox TxtMov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   54788097
         CurrentDate     =   39318
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   21
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   54788097
         CurrentDate     =   39318
      End
      Begin MSComctlLib.ImageCombo ImgCombo2 
         Height          =   360
         Left            =   8280
         TabIndex        =   72
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImageList1"
      End
      Begin MSComctlLib.ListView lvwMovimenti 
         Height          =   1815
         Left            =   120
         TabIndex        =   100
         Top             =   2400
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Causale"
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
         Index           =   12
         Left            =   10560
         TabIndex        =   98
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registrato"
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
         Index           =   11
         Left            =   8400
         TabIndex        =   96
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Utenza"
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
         Index           =   10
         Left            =   2280
         TabIndex        =   88
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Tariffa"
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
         Index           =   9
         Left            =   6480
         TabIndex        =   86
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Targa Traspor."
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
         Index           =   6
         Left            =   2280
         TabIndex        =   77
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipologia Rifiuto"
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
         Index           =   34
         Left            =   8280
         TabIndex        =   73
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   7
         Left            =   2280
         TabIndex        =   56
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Flag Bonus"
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
         Index           =   5
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stato Movimento"
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
         Index           =   4
         Left            =   6240
         TabIndex        =   48
         Top             =   120
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rifiuto"
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
         Index           =   3
         Left            =   6240
         TabIndex        =   47
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice Utente"
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
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Fine"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   32
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Inzio"
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
         Index           =   0
         Left            =   4200
         TabIndex        =   31
         Top             =   120
         Width           =   810
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
      Index           =   6
      Left            =   4800
      TabIndex        =   83
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancella"
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
      Left            =   3720
      TabIndex        =   79
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ristampa"
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
      Left            =   7200
      TabIndex        =   70
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Esporta"
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
      Left            =   6000
      TabIndex        =   23
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Salva/Stampa"
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
      Left            =   11280
      TabIndex        =   12
      Top             =   10320
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   3720
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesataMobile.frx":1ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPesataMobile.frx":264A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Lettura Peso"
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
      Left            =   8400
      TabIndex        =   11
      Top             =   10320
      Width           =   1335
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
      Left            =   13080
      TabIndex        =   13
      Top             =   10320
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   6015
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registrazione Movimento"
            Key             =   "aqpeso"
            Object.ToolTipText     =   "Aquisisce il  peso e le informazioni utente dalla pesa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Elenco Movimenti"
            Key             =   "elpesate"
            Object.ToolTipText     =   "Visualizza l'elenco delle pesate"
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   9240
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   80
      Top             =   10440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.PictureBox PicTotali 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8535
      TabIndex        =   90
      Top             =   10320
      Width           =   8535
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Index           =   3
         Left            =   5760
         TabIndex        =   94
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bonus Totale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3960
         TabIndex        =   93
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Index           =   1
         Left            =   2160
         TabIndex        =   92
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantita Totale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Width           =   1980
      End
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
      TabIndex        =   81
      Top             =   10200
      Width           =   5340
   End
End
Attribute VB_Name = "frmPesataMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHdrIcons As New cLVHeaderSortIcons


Private offsetX As Single
Private offsetY As Single

Private Const CMD_ESCI = 0
Private Const CMD_LETTURAQUANTITA = 1
Private Const CMD_SALVAESTAMPA = 2
Private Const CMD_RISTAMPAMOV = 3
Private Const CMD_ESPORTAMOV = 4
Private Const CMD_CANCELLAMOV = 5
Private Const CMD_RICERCAMOV = 6
Private Const CMD_SALVAMOV = 7


Private Const LBL_COMUNE = 2

Private Const LBL_RAGIONESOCIALE = 4
Private Const LBL_RAGIONESOCIALE_TRASP = 28
Private Const LBL_QUANTITA = 5
Private Const LBL_IMMOBILI = 6
Private Const LBL_BADGE = 9
Private Const LBL_CODICEFISCALE = 11
Private Const LBL_CODICEFISCALE_TRASP = 32
Private Const LBL_PARTITAIVA = 13
Private Const LBL_PARTITAIVA_TRASP = 30
Private Const LBL_INDIRIZZO = 15
Private Const LBL_BADGE_DT_INI = 19
Private Const LBL_BADGE_DT_FIN = 21
Private Const LBL_BADGE_PROGRESSIVO = 23
Private Const LBL_TRASPOR_RAGIONESOCIALE = 28
Private Const LBL_TRASPOR_CODCEFISCALE = 32
Private Const LBL_TRASPOR_PIVA = 30
Private Const LBL_FATTURABILE = 35

Private Const TXT_COD_UTENTE = 0
Private Const TXT_QUANTITA = 1
Private Const TXT_NRFORMULARIO = 4
Private Const TXT_FINDCODCER = 2
Private Const TXT_COD_TRASPORTATORE = 3

Private Const TXT_FINDCODCERMOV = 1

Private Const CB_FINDSTRING As Long = &H14C

Private Esci As Boolean
Private msgPesa As messaggioPesa
Private Canceling As Boolean
Private Deleting As Boolean

Private mTipoUtenza As String
Private mForzaturaMovimento As Boolean
Private rsCodiciCER  As ADODB.Recordset
Private rsCliente As ADODB.Recordset
Private mTotResiduoFatt As Double
Private sWhereRifiuti As String
Private midMovimentiDaStampare As String
Private mlastUtente As Long


Private Sub cmbRicercaMov_Change(Index As Integer)
    If Index = 0 Then
        chkEscludiCancellati.enabled = False
        If cmbRicercaMov(Index).ItemData(cmbRicercaMov(Index).ListIndex) = -1 Then
            chkEscludiCancellati.enabled = True
        End If
    End If

End Sub

Private Sub cmbRicercaMov_Click(Index As Integer)
    cmbRicercaMov_Change Index
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim codImmobile As String
    Dim idMovimento As Long
    Dim lpesa As frmLetturaPesata
    
    Select Case Index
        Case CMD_ESCI
        
            If MDIMain.Timer1.enabled = True Then
                 StampaMovimenti (True)
                 If MDIMain.Timer1.enabled = True Then Exit Sub
            End If
            
            Unload Me
        Case CMD_LETTURAQUANTITA
        
        
        Dim peso As Double
        Dim codBadge As String


            
            If cmd(CMD_LETTURAQUANTITA).Caption = "Anulla Pesata" Then
                cmd(CMD_LETTURAQUANTITA).Caption = "Lettura Peso"
                cmd(CMD_ESCI).enabled = True
                picContainer(0).enabled = True

                cmd(CMD_SALVAESTAMPA).enabled = True
                cmd(CMD_SALVAMOV).enabled = True
                cmd(CMD_ESCI).enabled = True

                Esci = True

                applicaRestrizioniUtente
            Else
                Esci = False
                picContainer(0).enabled = False
                cmd(CMD_ESCI).enabled = False
                cmd(CMD_SALVAESTAMPA).enabled = False
                cmd(CMD_SALVAMOV).enabled = False
                cmd(CMD_LETTURAQUANTITA).Caption = "Anulla Pesata"
                            
                            Set lpesa = New frmLetturaPesata
                        lpesa.showMe peso, codBadge
                        On Error Resume Next
                        Unload lpesa
                    Set lpesa = Nothing
            
            'messaggio = "1402/07/200815:17:1702/07/200815:17:34   144   145    200101       CARTA CARTONE    200101       CARTA CARTONE   1170007230010                         3.25 kg                             0.00 kg        0 kg        0 kg     3,25 kg21"
            With msgPesa
                .codBadge = codBadge
                .CodRifiuto = "200101"
                .Data = Now
                .PesoNetto = peso
                
            End With
            
            
            impostaDatiDaBadge msgPesa
            End If
            
        Case CMD_SALVAESTAMPA, CMD_SALVAMOV
        
            cmd(CMD_ESCI).enabled = False
            cmd(CMD_SALVAESTAMPA).enabled = False
            cmd(CMD_SALVAMOV).enabled = False
            TabStrip.enabled = False
                
            If (validaDati) Then
                If (salvaPesata(mForzaturaMovimento, idMovimento)) Then
                    
                    midMovimentiDaStampare = midMovimentiDaStampare & idMovimento & ","
                    If Index = CMD_SALVAESTAMPA Then
                        StampaMovimenti True
                    Else
                        MDIMain.attivaMultiPesata (True)
                        'blocco campi per evitare l'inserimento di dati diversi dalla prima pesata escluso il Peso e il Tipo rifiuto
                        txt(TXT_COD_TRASPORTATORE).Locked = True
                        txt(TXT_NRFORMULARIO).Locked = True
                        cmbTarga(1).Locked = True
                        cmbTarga(0).Locked = True
                        lvwImmobili.enabled = False
                    End If
                    
                    mTotResiduoFatt = 0
                    
                    'imposta campi form con i default
                    txt(TXT_QUANTITA).Text = vbNullString
                    txt(TXT_FINDCODCER).Text = vbNullString
                    
                    codImmobile = lvwImmobili.SelectedItem.key
                    
                    caricaDatiUtente txt(TXT_COD_UTENTE)
                    
                    If (lvwImmobili.ListItems(codImmobile).Checked = False) Then
                        lvwImmobili.ListItems(codImmobile).Checked = True
                        lvwImmobili_ItemCheck lvwImmobili.ListItems(codImmobile)
                    End If
                    
                    MsgBox "Operazione eseguita con successo!", vbExclamation, App.Title
                    
                    MDIMain.refreshForm
                    
                End If
            End If
            
            cmd(CMD_ESCI).enabled = True
            cmd(CMD_SALVAESTAMPA).enabled = True
            cmd(CMD_SALVAMOV).enabled = True
            TabStrip.enabled = True
        Case CMD_RICERCAMOV
            CaricaMovimenti
        Case CMD_ESPORTAMOV
            lblEsportazione.Visible = True
            ProgressBar1.Visible = True
            Esportazione
            lblEsportazione.Visible = False
            ProgressBar1.Visible = False
        Case CMD_CANCELLAMOV
        
            
            If Not lvwMovimenti.SelectedItem Is Nothing Then
                If MsgBox("Attenzione si vuole veramente cancellare il movimento ?" & vbCrLf & "Se il movimento è già contabilizzato non può essere cancellato", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    
                    gDBC.BeginTrans
                    
                    gDBC.Execute "Update Movimenti set stato = '2' where stato ='0' and id=" & Mid$(lvwMovimenti.SelectedItem.key, 2)
                    
                    gDBC.CommitTrans
                    cmd_Click CMD_RICERCAMOV
                End If
            End If
                
            
        Case CMD_RISTAMPAMOV
            If Not lvwMovimenti.SelectedItem Is Nothing Then
                If lvwMovimenti.SelectedItem.ListSubItems("NumeroFormulario") <> "" Then
                    ReportPesataConTrasportatore.Stampa Mid$(lvwMovimenti.SelectedItem.key, 2) & ",", pfrmPesata, True
                    Set ReportPesataConTrasportatore = Nothing
                Else
                    ReportPesata.Stampa Mid$(lvwMovimenti.SelectedItem.key, 2) & ",", pfrmPesata, True
                    Set ReportPesata = Nothing
                End If
            End If
    End Select
    
End Sub

Public Sub StampaMovimenti(Cancel As Boolean)
    
    If MDIMain.Timer1.enabled = True And Cancel Then
        If MsgBox("Attenzione, si vuole procedere con la stampa dei movimenti per l'utente:" & Label1(LBL_RAGIONESOCIALE).Caption, vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    End If
    
    
    If (txt(TXT_COD_TRASPORTATORE).enabled And txt(TXT_COD_TRASPORTATORE).Text <> "") Then
        ReportPesataConTrasportatore.Stampa midMovimentiDaStampare, pfrmPesata, CBool(gStampaPesataAnteprima), gStampaPesataNrCopie
    Else
        ReportPesata.Stampa midMovimentiDaStampare, pfrmPesata, CBool(gStampaPesataAnteprima), gStampaPesataNrCopie
         
         
    End If
    
    Unload ReportPesata
    Unload ReportPesataConTrasportatore
    
    midMovimentiDaStampare = vbNullString
    MDIMain.attivaMultiPesata (False)
    mlastUtente = 0
    
    'Ri-Attivo i campi
    txt(TXT_COD_TRASPORTATORE).Locked = False
    txt(TXT_NRFORMULARIO).Locked = False
    cmbTarga(1).Locked = False
    cmbTarga(0).Locked = False
    lvwImmobili.enabled = True
    
    'CLEAR fields trasportatore
    txt(TXT_COD_TRASPORTATORE).Text = vbNullString
    txt(TXT_NRFORMULARIO).Text = vbNullString
        
    caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE)
End Sub



Private Sub Esportazione()
    Dim Item As ListItem
    Dim ItemH As ColumnHeader
    Dim ItemS As ListSubItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim mChecked As Long
    
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gMovimentiFile & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
   
    If lvwMovimenti.ListItems.Count = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = lvwMovimenti.ListItems.Count
    ProgressBar1.Value = 0
    
    For Each ItemH In lvwMovimenti.ColumnHeaders
        txtStream.Write ItemH.Text
        If ItemH.Index < lvwMovimenti.ColumnHeaders.Count Then
            txtStream.Write gMovimentiSeparatore
        End If
    Next
    txtStream.Write vbCrLf
    
    For Each Item In lvwMovimenti.ListItems
        ProgressBar1.Value = ProgressBar1.Value + 1
        txtStream.Write Chr$(34) & Item.Text & Chr$(34)
        txtStream.Write gMovimentiSeparatore
        For Each ItemS In Item.ListSubItems
            
            txtStream.Write Chr$(34) & IIf(ItemS.Tag <> "", ItemS.Tag, ItemS.Text) & Chr$(34)
            If ItemS.Index < Item.ListSubItems.Count Then
                txtStream.Write gMovimentiSeparatore
            End If

        Next
        txtStream.Write vbCrLf
    Next
    
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    
    fso.MoveFile fileEXP, gPercorsoEsp & "\" & gMovimentiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
    
    
    Set fso = Nothing
    
    
    
    CaricaMovimenti
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title
End Sub

Private Function validaDati() As Boolean
    Dim campoInErrore As Boolean
    Dim MaxGGUD As Double
    Dim MaxGGUA As Double
    Dim MaxYYUA As Double
    Dim MaxYYUD As Double
    Dim MaxFatturabileUA As Double
    Dim TotMovGG As Double
    Dim TotMovYY As Double
    Dim TotResiduo As Double
    Dim quantita As Double
    Dim rs As ADODB.Recordset
    
    validaDati = False
    campoInErrore = False
    
    txt(TXT_COD_UTENTE).ToolTipText = ""
    txt(TXT_QUANTITA).ToolTipText = ""
    imgCombo1.ToolTipText = ""
    lvwImmobili.ToolTipText = ""
    Label1(LBL_BADGE).ToolTipText = ""
    Label1(LBL_BADGE_DT_INI).ToolTipText = ""
    Label1(LBL_BADGE_DT_FIN).ToolTipText = ""
    
    dtpickFormulario.ToolTipText = ""
    cmbTarga(1).ToolTipText = ""
    cmbTarga(0).ToolTipText = ""
    
    txt(TXT_COD_UTENTE).BackColor = -2147483643
    txt(TXT_QUANTITA).BackColor = -2147483643
    imgCombo1.BackColor = -2147483643
    lvwImmobili.BackColor = -2147483643
    Label1(LBL_BADGE).BackStyle = 0
    Label1(LBL_BADGE_DT_INI).BackStyle = 0
    Label1(LBL_BADGE_DT_FIN).BackStyle = 0
    
    dtpickFormulario.CalendarBackColor = -2147483643
    cmbTarga(1).BackColor = -2147483643
    cmbTarga(0).BackColor = -2147483643
        
    
    If (txt(TXT_COD_TRASPORTATORE).enabled And txt(TXT_COD_TRASPORTATORE).Text <> vbNullString) Then
    'validazioni su trasportatore
        If (txt(TXT_NRFORMULARIO).Text = "") Then
            txt(TXT_NRFORMULARIO).ToolTipText = "Attenzione numero formulario è obbligatorio"
            txt(TXT_NRFORMULARIO).BackColor = vbRed
            txt(TXT_NRFORMULARIO).SetFocus
            
            campoInErrore = True
        End If
        
        If (IsNull(dtpickFormulario.Value)) Then
            dtpickFormulario.ToolTipText = "Attenzione la data formulario è obbligatoria"
            dtpickFormulario.CalendarBackColor = vbRed
            dtpickFormulario.SetFocus
            
            campoInErrore = True
        End If

        If (cmbTarga(1).Text = "") Then
            cmbTarga(1).ToolTipText = "Attenzione la targa è obbligatoria"
            cmbTarga(1).BackColor = vbRed
            cmbTarga(1).SetFocus
            
            campoInErrore = True
        Else
            Set rs = gDBC.Execute("select codutente,codSoggetto from automezzi where codSoggetto='U' and  codutente<>" & txt(TXT_COD_TRASPORTATORE) & " and targa='" & cmbTarga(1).Text & "'")
            If Not rs.EOF Then
                cmbTarga(1).ToolTipText = "Attenzione la targa è già presente ed assegnata al utente:" & rs(0) & " tipo:" & rs(1)
                cmbTarga(1).BackColor = vbRed
                cmbTarga(1).SetFocus
                
                campoInErrore = True
            End If
            rs.Close
        End If
    End If
    
    If txt(TXT_COD_UTENTE) = "" Then
        txt(TXT_COD_UTENTE).ToolTipText = "Attenzione codice utente obbligatorio"
        txt(TXT_COD_UTENTE).BackColor = vbRed
        txt(TXT_COD_UTENTE).SetFocus

        campoInErrore = True
    End If
    
    If (cmbTarga(0).Text <> "") Then
        Set rs = gDBC.Execute("select codutente,codSoggetto from automezzi where codSoggetto='U' and  codutente<>" & txt(TXT_COD_UTENTE) & " and targa='" & cmbTarga(0).Text & "'")
        If Not rs.EOF Then
            cmbTarga(0).ToolTipText = "Attenzione la targa è già presente ed assegnata al utente:" & rs(0) & " tipo:" & rs(1)
            cmbTarga(0).BackColor = vbRed
            cmbTarga(0).SetFocus
                
            campoInErrore = True
        End If
        rs.Close
    End If
    
    If Format$(Date, "YYYY/MM/DD") < Format$(Label1(LBL_BADGE_DT_INI), "YYYY/MM/DD") Then
        Label1(LBL_BADGE_DT_INI).ToolTipText = "badge non ancora attivo"
        Label1(LBL_BADGE_DT_INI).BackStyle = 1
        Label1(LBL_BADGE_DT_INI).BackColor = vbRed
        
        If campoInErrore = False Then txt(TXT_COD_UTENTE).SetFocus
        campoInErrore = True
    End If
    
    If Format$(Date, "YYYY/MM/DD") > Format$(Label1(LBL_BADGE_DT_FIN), "YYYY/MM/DD") Then
    
        Label1(LBL_BADGE_DT_FIN).ToolTipText = "badge scaduto"
        Label1(LBL_BADGE_DT_FIN).BackStyle = 1
        Label1(LBL_BADGE_DT_FIN).BackColor = vbRed
        
        If campoInErrore = False Then txt(TXT_COD_UTENTE).SetFocus
        campoInErrore = True
    End If
    
    If imgCombo1.SelectedItem Is Nothing Then
        imgCombo1.ToolTipText = "Attenzione il tipo rifuto è obbligatorio"
        imgCombo1.BackColor = vbRed
        If campoInErrore = False Then imgCombo1.SetFocus
        campoInErrore = True
    ElseIf (Right$(imgCombo1.SelectedItem.key, 2) = "00") Or Len(imgCombo1.SelectedItem.key) < 6 Then
    
            imgCombo1.ToolTipText = "Attenzione il tipo rifuto non può essere una catedoria o una sottogategoria"
            imgCombo1.BackColor = vbRed
            If campoInErrore = False Then imgCombo1.SetFocus
            campoInErrore = True
    End If
    
    If lvwImmobili.Tag = vbNullString Then
        lvwImmobili.ToolTipText = "Attenzione l'immobile è obbligatorio"
        lvwImmobili.BackColor = vbRed
        If campoInErrore = False Then lvwImmobili.SetFocus
        campoInErrore = True
    End If
    
    If CDbl("0" & txt(TXT_QUANTITA)) <= 0 Then
        txt(TXT_QUANTITA).ToolTipText = "Attenzione la quantità deve essere > 0"
        txt(TXT_QUANTITA).BackColor = vbRed
        If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
        campoInErrore = True
    ElseIf campoInErrore = False Then
        
        'verifico se abilitato alla registrazione movimento
        If mTipoUtenza = "D" And getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.flgAmmessoPerUD) = False Then
            txt(TXT_QUANTITA).ToolTipText = "Attenzione l'Utenza Domestica non è abilitata a portare il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text
            txt(TXT_QUANTITA).Text = 0
            txt(TXT_QUANTITA).BackColor = vbRed
            MsgBox txt(TXT_QUANTITA).ToolTipText, vbExclamation + vbOKOnly, App.Title
            If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
            campoInErrore = True
            
        
        ElseIf mTipoUtenza <> "D" And getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.flgAmmessoPerUA) = False Then
            txt(TXT_QUANTITA).ToolTipText = "Attenzione l'Utenza Attività non è abilitata a portare il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text
            txt(TXT_QUANTITA).Text = 0
            txt(TXT_QUANTITA).BackColor = vbRed
            MsgBox txt(TXT_QUANTITA).ToolTipText, vbExclamation + vbOKOnly, App.Title
            If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
            campoInErrore = True
        
        Else
            MaxGGUD = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxGGUD)
            MaxYYUD = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxYYUD)
            
            MaxGGUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxGGUA)
            MaxYYUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxYYUA)
            
            MaxFatturabileUA = getPropietaCER(imgCombo1.SelectedItem.Tag, ProprietaCER.qtyMaxFatturabile)
            
            quantita = Round(getNumber(txt(TXT_QUANTITA).Text), 2)
            TotMovYY = getTotaleMovimentiAnno(imgCombo1.SelectedItem.key, Mid$(lvwImmobili.Tag, 2))
            TotMovGG = getTotaleMovimentiGiorno(imgCombo1.SelectedItem.key, Mid$(lvwImmobili.Tag, 2))
            
            If mTipoUtenza = "D" Then
                'limite quantita annuo UD
                If MaxYYUD > 0 And (MaxYYUD < TotMovYY + quantita) Then
                    txt(TXT_QUANTITA).ToolTipText = "Attenzione si è superato il limite annuo di:" & MaxYYUD
                    TotResiduo = IIf(TotMovYY >= MaxYYUD, 0, TotMovYY + quantita - MaxYYUD)
                    TotMovGG = MaxGGUD + TotResiduo
                End If
                'limite quantita giornaliero UD
                If MaxGGUD > 0 And (MaxGGUD < TotMovGG + quantita) Then
                    txt(TXT_QUANTITA).ToolTipText = txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Attenzione si è superato il limite giornaliero di:" & MaxGGUD
                    TotResiduo = IIf(TotMovGG >= MaxGGUD, 0, TotMovGG + quantita - MaxGGUD)
                End If
            Else
                'limite quantita annuo UA
                If MaxYYUA > 0 And (MaxYYUA < TotMovYY + quantita) Then
                    txt(TXT_QUANTITA).ToolTipText = "Attenzione si è superato il limite annuo di:" & MaxYYUA
                     TotResiduo = IIf(TotMovYY >= MaxYYUA, 0, TotMovYY + quantita - MaxYYUA)
                    TotMovGG = MaxGGUA + TotResiduo
                End If
                'limite quantita giornaliero UA
                If MaxGGUA > 0 And (MaxGGUA < TotMovGG + quantita) Then
                    txt(TXT_QUANTITA).ToolTipText = txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Attenzione si è superato il limite giornaliero di:" & MaxGGUA
                    TotResiduo = IIf(TotMovGG >= MaxGGUA, 0, TotMovGG + quantita - MaxGGUA)
                End If
                
                'Fatturazione
                mTotResiduoFatt = 0
                If getPropietaCER(imgCombo1.SelectedItem.Tag, FlgAbilitatoFatturazione) = True Then
                'If getPropietaCER(imgCombo1.SelectedItem.Tag, FlgAbilitatoFatturazione) = True And txt(TXT_QUANTITA).ToolTipText = "" Then
                    'limite quantita Fatturazione UA
                    If MaxFatturabileUA < TotMovYY + quantita Then
                        'txt(TXT_QUANTITA).ToolTipText = txt(TXT_QUANTITA).ToolTipText & ",Si è superato il massimo consentito di:" & MaxFatturabileUA & " non fatturabile!"
                         mTotResiduoFatt = IIf(TotMovYY >= MaxFatturabileUA, quantita, TotMovYY + quantita - MaxFatturabileUA)
                    End If
                End If
                
            End If
            mForzaturaMovimento = False
            If txt(TXT_QUANTITA).ToolTipText <> "" Then
                'If (mTotResiduoFatt <= 0) Then
                    If (MsgBox(txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Si vuole comunque forzare il movimento?", vbQuestion + vbYesNo, App.Title) = vbNo) Then
                        txt(TXT_QUANTITA).Text = 0
                        If (TotResiduo > 0) Then
                            MsgBox "Occore effettuare una altra pesata togliendo la quantità in eccesso pari a:" & TotResiduo, vbExclamation + vbOKOnly, App.Title
                        Else
                            MsgBox "Non è possibile effetuare altre registrazioni su questo immobile per il rifiuto:" & vbCrLf & imgCombo1.SelectedItem.Text, vbExclamation + vbOKOnly, App.Title
                        End If
                        
                        txt(TXT_QUANTITA).BackColor = vbRed
                        If campoInErrore = False Then txt(TXT_QUANTITA).SetFocus
                        campoInErrore = True
                    Else
                        mForzaturaMovimento = True
                    End If
                'Else
                '    MsgBox txt(TXT_QUANTITA).ToolTipText & vbCrLf & "Verrà creato un record per la fatturazione pari a quantità:" & mTotResiduoFatt, vbExclamation + vbOKOnly, App.Title
                'End If
            End If
        End If
    End If
    
    validaDati = Not campoInErrore
End Function

Private Function salvaPesata(forzatura As Boolean, idMovimento As Long) As Boolean
    On Error GoTo salvaPesataERR
    Dim codImmobile As Long
    Dim codUtente As Long
    Dim codCategoria As String
    Dim codSottoCategoria As String
    Dim codCer As String
    Dim DataMovimento As String
    Dim quantita As Double
    Dim numeroPezzi As Double
    Dim BonusDovuto As Double
    Dim nrMov As Long
    Dim idRegistro As Long
    Dim nrRegistro As Long
    Dim manuale As Boolean
    Dim registrato As Boolean
    
    gDBC.BeginTrans
    
    
    nrMov = getNumeratore(enumEntita.Movimenti)
    salvaPesata = False
    codImmobile = lvwImmobili.SelectedItem.Text
    codUtente = txt(TXT_COD_UTENTE).Text
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = Left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codCer = Right$(codCer, 2)
    
    DataMovimento = IIf(CLng(msgPesa.Data) <= 0, Now, msgPesa.Data)
    quantita = Round(getNumber(txt(TXT_QUANTITA).Text), 2)
    numeroPezzi = 0
    BonusDovuto = getNumber(getPropietaCER(imgCombo1.SelectedItem.Tag, CorrispettivoBonus))
    BonusDovuto = Round(BonusDovuto * quantita, 4)
    manuale = IIf(CLng(msgPesa.Data) <= 0, True, False)
    registrato = IIf(txt(TXT_NRFORMULARIO).Text <> "", True, False)
    
    idMovimento = creaMovimento(nrMov, codImmobile, codUtente, codCategoria, codSottoCategoria, codCer, _
        DataMovimento, quantita, "0", manuale, frmLogin.codUser, _
        BonusDovuto, getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura), _
        cmbTarga(0).Text, txt(TXT_NRFORMULARIO).Text, Format$(dtpickFormulario.Value, "YYYY-MM-DD"), _
         forzatura, cmbTarga(1).Text, CLng(IIf(txt(TXT_COD_TRASPORTATORE).Text = "", -1, _
         txt(TXT_COD_TRASPORTATORE).Text)), "00", registrato, gCodPesaIsola)


    salvaTarga txt(TXT_COD_UTENTE).Text, cmbTarga(0).Text, "U"
    'salvaTarga txt(TXT_COD_TRASPORTATORE).Text, cmbTarga(1).Text, "T"

    'crea record per fatturazione
    If mTotResiduoFatt > 0 Then
        creaRecordPerFatturazione mTotResiduoFatt, DataMovimento, codCategoria & codSottoCategoria & codCer
    End If

    
    If txt(TXT_NRFORMULARIO).Text <> "" Then
    
        nrRegistro = getNumeratore(REGISTRO, False)
        
        idRegistro = creaVoceDiRegistro(nrRegistro, "C", Format$(DataMovimento, "yyyy/MM/dd"), txt(TXT_NRFORMULARIO).Text, Format$(dtpickFormulario.Value, "YYYY-MM-DD") _
            , codCategoria, codSottoCategoria, codCer, "", getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura), _
                quantita, "", 0, "", Now, frmLogin.codUser, "")
    
    
        creaLotto idRegistro, codCategoria, codSottoCategoria, codCer, _
                  quantita, 0, "I", 0, Now, frmLogin.codUser, Now, frmLogin.codUser
    
    End If
    
    gDBC.CommitTrans
    
         
     
    salvaPesata = True
    Exit Function
salvaPesataERR:
    
    gDBC.RollbackTrans
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Function



Private Sub creaRecordPerFatturazione(quantita As Double, DataMovimento As String, codCer As String)
    Dim cmd As ADODB.Command
    Dim p As Parameter
    Dim rs As ADODB.Recordset
    
    'se utenza Non domestica verifico se il totale delle pesate annuali supera il max
    'se si viene fatturata la differenza
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO ST_AMIU_FILE_ESTERNO ( AZIENDA, PROGRESSIVO_INVIO, CODCLI, RAGSO1, " & _
            "VIA, CAP, CITTA, PIVA, CODFIS, PROV, TELEF1, TELEF2, CODPSE, QUANT, DATA_EFFETTIVA, TIPO_MODELLO, " & _
            "NOTE1, STATO_INVIO, CODICE_RIFIUTO, PRGCLI, RAGSO_SEDE, VIA_SEDE, CAP_SEDE, CITTA_SEDE, PROV_SEDE, " & _
            "DATA_SCARICO, TIPCLI, NUMCON ) values (" & _
            "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?, ?, ?)"
            
    cmd.Parameters.Append cmd.CreateParameter("AZIENDA", adVarChar, adParamInput, 10)
    cmd.Parameters.Append cmd.CreateParameter("PROGRESSIVO_INVIO", adInteger, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("CODCLI", adVarChar, adParamInput, 5)
    cmd.Parameters.Append cmd.CreateParameter("RAGSO1", adVarChar, adParamInput, 60)
    cmd.Parameters.Append cmd.CreateParameter("VIA", adVarChar, adParamInput, 30)
    cmd.Parameters.Append cmd.CreateParameter("CAP", adVarChar, adParamInput, 5)
    cmd.Parameters.Append cmd.CreateParameter("CITTA", adVarChar, adParamInput, 45)
    cmd.Parameters.Append cmd.CreateParameter("PIVA", adVarChar, adParamInput, 11)
    cmd.Parameters.Append cmd.CreateParameter("CODFIS", adVarChar, adParamInput, 16)
    cmd.Parameters.Append cmd.CreateParameter("PROV", adVarChar, adParamInput, 4)
    cmd.Parameters.Append cmd.CreateParameter("TELEF1", adVarChar, adParamInput, 18)
    cmd.Parameters.Append cmd.CreateParameter("TELEF2", adVarChar, adParamInput, 18)
    cmd.Parameters.Append cmd.CreateParameter("CODPSE", adVarChar, adParamInput, 4)
    
    
    Set p = cmd.CreateParameter("QUANTITA", adDouble, adParamInput)
    p.NumericScale = 2
    p.Precision = 6
    cmd.Parameters.Append p
    
    cmd.Parameters.Append cmd.CreateParameter("DATA_EFFETTIVA", adDate, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("TIPO_MODELLO", adVarChar, adParamInput, 4)
    cmd.Parameters.Append cmd.CreateParameter("NOTE1", adVarChar, adParamInput, 50)
    cmd.Parameters.Append cmd.CreateParameter("STATO_INVIO", adInteger, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("CODICE_RIFIUTO", adVarChar, adParamInput, 6)
    cmd.Parameters.Append cmd.CreateParameter("PRGCLI", adVarChar, adParamInput, 2)
    cmd.Parameters.Append cmd.CreateParameter("RAGSO_SEDE", adVarChar, adParamInput, 60)
    cmd.Parameters.Append cmd.CreateParameter("VIA_SEDE", adVarChar, adParamInput, 30)
    cmd.Parameters.Append cmd.CreateParameter("CAP_SEDE", adVarChar, adParamInput, 5)
    cmd.Parameters.Append cmd.CreateParameter("CITTA_SEDE", adVarChar, adParamInput, 45)
    cmd.Parameters.Append cmd.CreateParameter("PROV_SEDE", adVarChar, adParamInput, 4)
    cmd.Parameters.Append cmd.CreateParameter("DATA_SCARICO", adDate, adParamInput)
    cmd.Parameters.Append cmd.CreateParameter("TIPCLI", adVarChar, adParamInput, 14)
    cmd.Parameters.Append cmd.CreateParameter("NUMCON", adInteger, adParamInput, 18)
    
    Set cmd.ActiveConnection = gDBC
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    rs.Open ("SELECT utenti.*, Comuni_1.Comune AS ComuneResidenza, Comuni_1.Provincia AS ProvinciaResidenza, Comuni.Comune AS ComuneRecapito, Comuni.Provincia AS ProvinciaRecapito, Comuni.Cap AS CapRecapito, Comuni_1.Cap AS CapResidenza  " & _
             "FROM (utenti LEFT JOIN Comuni AS Comuni_1 ON utenti.CodComuneResidenza = Comuni_1.CODISTAT) LEFT JOIN Comuni ON utenti.CodComuneRecapito = Comuni.CODISTAT Where codUtente = " & txt(TXT_COD_UTENTE).Text)
    Set rs.ActiveConnection = Nothing
    
    cmd.Parameters("AZIENDA").Value = "ST_"
    cmd.Parameters("PROGRESSIVO_INVIO").Value = getNumeratore(FATTURE, False)
    cmd.Parameters("CODCLI").Value = txt(TXT_COD_UTENTE).Text
    cmd.Parameters("RAGSO1").Value = nullToBlank(rs!RagioneSociale)
    cmd.Parameters("VIA").Value = nullToBlank(rs!IndirizzoResidenza)
    cmd.Parameters("CAP").Value = nullToBlank(rs!CapResidenza)
    cmd.Parameters("CITTA").Value = nullToBlank(rs!ComuneResidenza)
    cmd.Parameters("PIVA").Value = nullToBlank(rs!PartitaIva)
    cmd.Parameters("CODFIS").Value = nullToBlank(rs!CodFiscale)
    cmd.Parameters("PROV").Value = nullToBlank(rs!ProvinciaResidenza)
    cmd.Parameters("TELEF1").Value = nullToBlank(rs!NumeroTelefono)
    cmd.Parameters("TELEF2").Value = Null
    cmd.Parameters("CODPSE").Value = "CNFR"
    cmd.Parameters("QUANTITA").Value = quantita
    cmd.Parameters("DATA_EFFETTIVA").Value = DataMovimento
    cmd.Parameters("TIPO_MODELLO").Value = "xxx"
    cmd.Parameters("NOTE1").Value = Null
    cmd.Parameters("STATO_INVIO").Value = 5
    cmd.Parameters("CODICE_RIFIUTO").Value = codCer
    cmd.Parameters("PRGCLI").Value = Null
    cmd.Parameters("RAGSO_SEDE").Value = nullToBlank(rs!NominativoRecapito)
    cmd.Parameters("VIA_SEDE").Value = nullToBlank(rs!IndirizzoRecapito)
    cmd.Parameters("CAP_SEDE").Value = nullToBlank(rs!CapRecapito)
    cmd.Parameters("CITTA_SEDE").Value = nullToBlank(rs!ComuneRecapito)
    cmd.Parameters("PROV_SEDE").Value = nullToBlank(rs!ProvinciaRecapito)
    cmd.Parameters("DATA_SCARICO").Value = Null
    cmd.Parameters("TIPCLI").Value = Null
    cmd.Parameters("NUMCON").Value = Null
    
    rs.Close
    cmd.Execute
End Sub


Private Sub cmbRicerca_Click(Index As Integer)
    Select Case Index
        Case 0
            frmRicercaUtente.schowMe txt(TXT_COD_UTENTE), "U"
            caricaDatiUtente txt(TXT_COD_UTENTE).Text
        Case 1
            frmRicercaUtente.schowMe txt(TXT_COD_TRASPORTATORE), "T"
            caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE).Text
        Case 1
    End Select
End Sub



 
Private Sub caricaDatiTrasportatore(CodTrasportatore As String)
    Dim rs As Recordset
    
    Label1(LBL_TRASPOR_CODCEFISCALE).Caption = ""
    Label1(LBL_TRASPOR_PIVA).Caption = ""
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
    
    
    If CodTrasportatore <> vbNullString Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.LockType = adLockBatchOptimistic
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Source = "select *,comuni.comune from Trasportatori LEFT JOIN Comuni ON Trasportatori.CodComuneSedeLegale = Comuni.CODISTAT where idTrasportatore=" & CodTrasportatore
        
        Set rs.ActiveConnection = gDBC
        rs.Open
        Set rs.ActiveConnection = Nothing
                
        If rs.EOF Then
            MsgBox "Attenzione Trasportatore:" & CodTrasportatore & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
        Else
            
            Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = rs!RagioneSociale
            DoEvents
            
            
            Label1(LBL_TRASPOR_CODCEFISCALE).Left = Label1(LBL_TRASPOR_RAGIONESOCIALE).Width + Label1(LBL_TRASPOR_RAGIONESOCIALE).Left + 2 * offsetX
            Label1(LBL_TRASPOR_CODCEFISCALE - 1).Left = Label1(LBL_TRASPOR_RAGIONESOCIALE).Width + Label1(LBL_TRASPOR_RAGIONESOCIALE).Left + 2 * offsetX
            Label1(LBL_TRASPOR_CODCEFISCALE).Caption = nullToBlank(rs!CodFiscale)
            DoEvents
            
            If (Label1(LBL_TRASPOR_CODCEFISCALE - 1).Width > Label1(LBL_TRASPOR_CODCEFISCALE).Width) Then
                Label1(LBL_TRASPOR_PIVA).Left = Label1(LBL_TRASPOR_CODCEFISCALE - 1).Width + Label1(LBL_TRASPOR_CODCEFISCALE - 1).Left + 2 * offsetX
                Label1(LBL_TRASPOR_PIVA - 1).Left = Label1(LBL_TRASPOR_CODCEFISCALE - 1).Width + Label1(LBL_TRASPOR_CODCEFISCALE - 1).Left + 2 * offsetX
            Else
                Label1(LBL_TRASPOR_PIVA).Left = Label1(LBL_TRASPOR_CODCEFISCALE).Width + Label1(LBL_TRASPOR_CODCEFISCALE).Left + 2 * offsetX
                Label1(LBL_TRASPOR_PIVA - 1).Left = Label1(LBL_TRASPOR_CODCEFISCALE).Width + Label1(LBL_TRASPOR_CODCEFISCALE).Left + 2 * offsetX
            End If
            Label1(LBL_TRASPOR_PIVA).Caption = nullToBlank(rs!PartitaIva)
            
            DoEvents
            
            Set rsCliente = rs.Clone
            
        End If
        rs.Close
    End If
    
    'carica combo targhe
    cmbTarga(1).Clear
    
    If CodTrasportatore <> "" Then
        If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
            DevEnvMDB.rscmdAutomezzi.Close
        End If
        DevEnvMDB.cmdAutomezzi CodTrasportatore, "T"
        Do While Not DevEnvMDB.rscmdAutomezzi.EOF
            cmbTarga(1).AddItem DevEnvMDB.rscmdAutomezzi!Targa
            DevEnvMDB.rscmdAutomezzi.MoveNext
        Loop
        If cmbTarga(1).ListCount > 0 Then cmbTarga(1).ListIndex = 0
    End If
    
    Set rs = Nothing
End Sub


Private Sub caricaDatiUtente(codUtente As String)
    Dim rs As Recordset
    
    Label3(1).Caption = 0
    Label3(3).Caption = 0
    Label1(LBL_COMUNE).Caption = ""
    Label1(LBL_RAGIONESOCIALE).Caption = ""
    Label1(LBL_BADGE).Caption = ""
    Label1(LBL_CODICEFISCALE).Caption = ""
    Label1(LBL_PARTITAIVA).Caption = ""
    Label1(LBL_BADGE_DT_INI).Caption = ""
    Label1(LBL_BADGE_DT_FIN).Caption = ""
    Label1(LBL_BADGE_PROGRESSIVO).Caption = ""
    
    mTipoUtenza = ""
    
    If codUtente = vbNullString Then Exit Sub
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Source = "select utenti.*,codBadge,DataBadgeInizio,DataBadgefine,ProgressivoBadge,comuni.comune, (badge.DataBadgeFine - date()) as ScadenzaGG from ((Utenti  LEFT JOIN Comuni ON Utenti.CodComuneResidenza = Comuni.CODISTAT) inner join badge on badge.codUtente = utenti.codutente) where badge.default=true and utenti.codUtente=" & codUtente
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
            
    If rs.EOF Then
        MsgBox "Attenzione CodUtente:" & codUtente & " non trovato in Anagrafica", vbExclamation + vbOKOnly, App.Title
    Else
        
        Label1(LBL_RAGIONESOCIALE).Caption = IIf(nullToBlank(rs!RagioneSociale) = vbNullString, nullToBlank(rs!Nome) & " " & nullToBlank(rs!Cognome), nullToBlank(rs!RagioneSociale))
        DoEvents
        
        Label1(LBL_COMUNE).Left = Label1(LBL_RAGIONESOCIALE).Width + Label1(LBL_RAGIONESOCIALE).Left + 2 * offsetX
        Label1(LBL_COMUNE).Caption = nullToBlank(rs!comune.Value)
        Label1(1).Left = Label1(LBL_RAGIONESOCIALE).Width + Label1(LBL_RAGIONESOCIALE).Left + 2 * offsetX
        
        DoEvents
        
        If (Label1(1).Width > Label1(LBL_COMUNE).Width) Then
            Label1(LBL_INDIRIZZO).Left = Label1(1).Width + Label1(1).Left + 2 * offsetX
            Label1(16).Left = Label1(1).Width + Label1(1).Left + 2 * offsetX
        
        Else
            Label1(LBL_INDIRIZZO).Left = Label1(LBL_COMUNE).Width + Label1(LBL_COMUNE).Left + 2 * offsetX
            Label1(16).Left = Label1(LBL_COMUNE).Width + Label1(LBL_COMUNE).Left + 2 * offsetX
        End If
        Label1(LBL_INDIRIZZO).Caption = nullToBlank(rs!IndirizzoResidenza) & ", " & nullToBlank(rs!NumCivicoResidenza)
        
        DoEvents
        
        Label1(LBL_BADGE).Caption = nullToBlank(rs!codBadge)
        Label1(LBL_CODICEFISCALE).Caption = nullToBlank(rs!CodFiscale)
        Label1(LBL_PARTITAIVA).Caption = nullToBlank(rs!PartitaIva)
        Label1(LBL_BADGE_DT_INI).Caption = nullToBlank(rs!DataBadgeInizio)
        Label1(LBL_BADGE_DT_FIN).Caption = nullToBlank(rs!DataBadgefine)
        Label1(LBL_BADGE_PROGRESSIVO).Caption = nullToBlank(rs!ProgressivoBadge)
        
        
        Set rsCliente = rs.Clone
                
    End If
    rs.Close
    
    caricaImmobili codUtente
    
    DoEvents
    
    'carica combo targhe
    cmbTarga(0).Clear
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    DevEnvMDB.cmdAutomezzi codUtente, "U"
    Do While Not DevEnvMDB.rscmdAutomezzi.EOF
        cmbTarga(0).AddItem DevEnvMDB.rscmdAutomezzi!Targa
        DevEnvMDB.rscmdAutomezzi.MoveNext
    Loop
    
    If mlastUtente <> 0 And mlastUtente <> codUtente And midMovimentiDaStampare <> vbNullString Then
        MsgBox "Attenzione L'utente è cambiato, verrà ora avviata la stampa movimenti", vbExclamation + vbOKOnly, App.Title
        StampaMovimenti False
    End If
    
    mlastUtente = codUtente
    
    If cmbTarga(0).ListCount > 0 Then cmbTarga(0).ListIndex = 0
End Sub

Private Sub caricaTotali(codImmobile As Long)
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim UnitaMisura As String
    Dim totQTA As Double
    Dim totBONUS As Double
    
    
    lvwTotMovimenti.ListItems.Clear
    
    If codImmobile = 0 Then Exit Sub
    
    Set rs = gDBC.Execute("SELECT Movimenti.Stato, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer,Movimenti.UnitaMisura, CodiciCer.Descrizione, Sum(Movimenti.BonusDovuto) AS Bonus, Sum(Movimenti.Quantita) AS Quantita, " & _
        "IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxGGUD],[CodiciCer].[qtyMaxGGUA]) AS qtyMaxGG, IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxAAUD],[CodiciCer].[qtyMaxAAUA]) AS qtyMaxAA " & _
        "FROM (Movimenti " & _
        "INNER JOIN CodiciCer ON (Movimenti.CodCategoria = CodiciCer.CodCategoria) AND (Movimenti.CodCer = CodiciCer.CodCer) AND (Movimenti.CodSottoCategoria = CodiciCer.CodSottoCategoria)) " & _
        "INNER JOIN Immobili ON Movimenti.CodImmobile = Immobili.CodImmobile and Movimenti.CodUtente = Immobili.CodUtente " & _
        "WHERE Movimenti.codImmobile = " & codImmobile & " and Movimenti.Stato='0' and (codCausale='00'" & _
        " or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
        "GROUP BY Movimenti.Stato, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer,Movimenti.UnitaMisura, CodiciCer.Descrizione, IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxGGUD],[CodiciCer].[qtyMaxGGUA]), IIf([Immobili].[TipoUtenza]='D',[CodiciCer].[qtyMaxAAUD],[CodiciCer].[qtyMaxAAUA]), CodiciCer.FlgBonus " & _
        "ORDER BY CodiciCer.FlgBonus ")
    
    Set Item = lvwTotMovimenti.ListItems.Add(, , "")
    Item.ForeColor = vbRed
    Item.ListSubItems.Add , , "Tot.Quantità"
    Item.ListSubItems(Item.ListSubItems.Count).ForeColor = vbRed
    Item.ListSubItems.Add , , "Tot.Bonus (€)"
    Item.ListSubItems(Item.ListSubItems.Count).ForeColor = vbRed
    Item.ListSubItems.Add , , "Max Qty Giornaliera"
    Item.ListSubItems(Item.ListSubItems.Count).ForeColor = vbRed
    Item.ListSubItems.Add , , "Max Qty Annua"
    Item.ListSubItems(Item.ListSubItems.Count).ForeColor = vbRed
    
    totQTA = 0
    totBONUS = 0
    
    Do While Not rs.EOF
        Set Item = lvwTotMovimenti.ListItems.Add(, "_" & rs!codCategoria & rs!codSottoCategoria & rs!codCer, rs!descrizione & " (" & rs!UnitaMisura & ")")
        
        totQTA = totQTA + rs!quantita
        totBONUS = totBONUS + rs!Bonus

        Item.ListSubItems.Add , , rs!quantita
        Item.ListSubItems.Add , , rs!Bonus
        Item.ListSubItems.Add , , nullToBlank(rs!qtyMaxGG)
        Item.ListSubItems.Add , , nullToBlank(rs!qtyMaxAA)
        rs.MoveNext
    Loop
    
    Label3(1).Caption = FormatNumber(totQTA, 2, vbUseDefault, vbUseDefault, vbTrue) & " Kg"
    Label3(3).Caption = FormatNumber(totBONUS, 2, vbUseDefault, vbUseDefault, vbTrue) & " €"
 
        
    
    rs.Close
    
    AllargaLeColonneLVW lvwTotMovimenti, False
End Sub


Private Sub cmdRicercaUtenteMov_Click(Index As Integer)
    Dim utente As TextBox
    Dim codUte As String
    
    If Index = 0 Then
        Set utente = TxtMov(0)
        codUte = "U"
    Else
        Set utente = TxtMov(2)
        codUte = "T"
    End If
    
    If (cmdRicercaUtenteMov(0).Tag = "") Then
        frmRicercaUtente.schowMe utente, codUte
    End If
    cmdRicercaUtenteMov(0).Tag = ""
   
    'carica combo targhe
    cmbTargaMov(Index).Clear
    If (utente = "") Then Exit Sub
    
    cmbTargaMov(Index).AddItem "Tutte"
    If DevEnvMDB.rscmdAutomezzi.state = adStateOpen Then
        DevEnvMDB.rscmdAutomezzi.Close
    End If
    
    DevEnvMDB.cmdAutomezzi utente, codUte
    Do While Not DevEnvMDB.rscmdAutomezzi.EOF
        cmbTargaMov(Index).AddItem DevEnvMDB.rscmdAutomezzi!Targa
        DevEnvMDB.rscmdAutomezzi.MoveNext
    Loop

    cmbTargaMov(Index).ListIndex = 0
End Sub

Private Sub imgCombo1_Click()
    Label1(LBL_FATTURABILE).Visible = False
    'se viene selezionato un codice CER imposto unità di misura
    If Not imgCombo1.SelectedItem Is Nothing Then
        If (Right$(imgCombo1.SelectedItem.key, 2) <> "00") And Len(imgCombo1.SelectedItem.key) > 6 Then
            Label1(LBL_QUANTITA) = "Quantità (" & getPropietaCER(imgCombo1.SelectedItem.Tag, UnitaDiMisura) & ")"
            
            Label1(LBL_FATTURABILE).Visible = mTipoUtenza = "N" And CBool(getPropietaCER(imgCombo1.SelectedItem.Tag, FlgAbilitatoFatturazione))
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    
    Label1(LBL_FATTURABILE).Visible = False
    lblEsportazione.Visible = False
    ProgressBar1.Visible = False
    applicaRestrizioniUtente
    
    For i = picContainer.LBound To picContainer.ubound
        picContainer(i).Visible = False
    Next i

    caricaComboRifiuti imgCombo1, Combo2(0)
    caricaComboRifiuti ImgCombo2, Combo2(1)
    
    Set rs = CreateObject("ADODB.Recordset")
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from causali"
    rs.Open
    Do While Not rs.EOF
        cmbRicercaMov(5).AddItem rs!codCausale & " - " & rs!descrizione
        rs.MoveNext
    Loop
    
    dtpickFormulario.Value = Now
    
    cmbRicercaMov(0).ListIndex = 1
    cmbRicercaMov(1).ListIndex = 0
    cmbRicercaMov(2).ListIndex = 0
    cmbRicercaMov(3).ListIndex = 0
    
    cmbTarga(0).Text = ""
    'cmbTarga(1).Text = ""
    cmbTargaMov(0).AddItem "Tutte"
    cmbTargaMov(0).ListIndex = 0
    cmbTargaMov(1).AddItem "Tutte"
    cmbTargaMov(1).ListIndex = 0
    
    Set rs = gDBC.Execute("select * from tariffe order by codTariffa")
    Do While Not rs.EOF
        cmbRicercaMov(2).AddItem (rs!codTariffa & " - " & rs!descrizione)
        rs.MoveNext
    Loop
    
    
    Label1(LBL_COMUNE).Caption = ""
    Label1(LBL_RAGIONESOCIALE).Caption = ""
    Label1(LBL_PARTITAIVA).Caption = ""
    Label1(LBL_CODICEFISCALE).Caption = ""
    Label1(LBL_INDIRIZZO).Caption = ""
    Label1(LBL_BADGE_DT_FIN).Caption = ""
    Label1(LBL_BADGE_DT_INI).Caption = ""
    Label1(LBL_BADGE_PROGRESSIVO).Caption = ""
    Label1(LBL_BADGE).Caption = ""
    
    
    
    Label1(LBL_TRASPOR_CODCEFISCALE).Caption = ""
    Label1(LBL_TRASPOR_PIVA).Caption = ""
    Label1(LBL_TRASPOR_RAGIONESOCIALE).Caption = ""
    
    lvwImmobili.ColumnHeaders.Add , "cod", "Cod.Imm."
    lvwImmobili.ColumnHeaders.Add , , "Tipo"
    lvwImmobili.ColumnHeaders(lvwImmobili.ColumnHeaders.Count).Width = 0
    lvwImmobili.ColumnHeaders.Add , , "Comune"
    lvwImmobili.ColumnHeaders.Add , "TipoUtenza", "Tipo Utenza"
    lvwImmobili.ColumnHeaders.Add , , "Indirizzo"
    lvwImmobili.ColumnHeaders.Add , , "SubAlterno"
    lvwImmobili.ColumnHeaders.Add , , "Num.Interno"
    lvwImmobili.ColumnHeaders.Add , , "Cod.Tariffa"
    lvwImmobili.ColumnHeaders.Add , , "Tariffa"
    
    
    lvwMovimenti.ColumnHeaders.Clear
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Utente"
    lvwMovimenti.ColumnHeaders.Add , , "Utente"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Immobile"
    lvwMovimenti.ColumnHeaders.Add , , "Immobile"
    lvwMovimenti.ColumnHeaders.Add , , "Tip.Ute."
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Tariffa"
    lvwMovimenti.ColumnHeaders.Add , , "Targa"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Cer"
    lvwMovimenti.ColumnHeaders.Add , , "Tipo Rifiuto"
    lvwMovimenti.ColumnHeaders.Add , , "Data Mov."
    lvwMovimenti.ColumnHeaders.Add , , "Quantità"
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "U.M."
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "Bonus €"
    lvwMovimenti.ColumnHeaders(lvwMovimenti.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovimenti.ColumnHeaders.Add , , "Causale"
    lvwMovimenti.ColumnHeaders.Add , , "Reg."
    lvwMovimenti.ColumnHeaders.Add , , "Stato"
    lvwMovimenti.ColumnHeaders.Add , , "Cod.Trasp."
    lvwMovimenti.ColumnHeaders.Add , , "Trasportatore"
    lvwMovimenti.ColumnHeaders.Add , , "Targa Trasp."
    lvwMovimenti.ColumnHeaders.Add , "NumeroFormulario", "Nr.Formulario"
    lvwMovimenti.ColumnHeaders.Add , , "Data Formulario"
    lvwMovimenti.ColumnHeaders.Add , , "Tipo"
    lvwMovimenti.ColumnHeaders.Add , , "Utente Ins."
    lvwMovimenti.ColumnHeaders.Add , , "Data Ins."
    
    
    DTPicker(0).Value = Year(Now) & "/" & Right$(0 & (Month(Now)), 2) & "/" & 1
    '
    DTPicker(1).Value = DateAdd("d", -Day(DateAdd("m", 1, Now)), DateAdd("m", 1, Now))
        
    cmbRicercaMov_Change 0
    
    cmbRicercaMov(4).Clear
    cmbRicercaMov(4).AddItem ("Tutti")
    cmbRicercaMov(4).AddItem ("Da Registrare")
    cmbRicercaMov(4).AddItem ("Registrati")
    cmbRicercaMov(4).ListIndex = 0

    For i = Label3.LBound To Label3.ubound
        Label3(i).Visible = False
    Next
    
End Sub

Private Sub CaricaMovimenti()
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim where As String
    Dim codCer As String
    Dim codSottoCategoria As String
    Dim codCategoria As String
    Dim qty As Double

    
   ' where = "(Movimenti.codCausale ='00' or Movimenti.codCausale ='" & gCodCausaleMovManuale & "')"
    
     where = ""
    Screen.MousePointer = vbHourglass
    
    If TxtMov(0).Text <> "" Then
        If where <> "" Then where = where & " and "
        where = "Immobili.CodUtente =" & TxtMov(0)
    End If
    
    If TxtMov(2).Text <> "" Then
        If where <> "" Then where = where & " and "
        where = "Movimenti.CodTrasportatore =" & TxtMov(2)
    End If
    
    
    If cmbRicercaMov(0).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
            
        where = where & "Movimenti.Stato ='" & cmbRicercaMov(0).ItemData(cmbRicercaMov(0).ListIndex) & "'"
    ElseIf (chkEscludiCancellati.enabled And chkEscludiCancellati.Value = 1) Then
        If where <> "" Then where = where & " and "
        
        where = where & " Movimenti.Stato <>'2 '"
    End If
    
    If cmbRicercaMov(5).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.codCausale ='" & Left$(cmbRicercaMov(5).Text, 2) & "'"
    End If
    
    If cmbRicercaMov(1).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "CodiciCer.FlgBonus =" & cmbRicercaMov(1).ItemData(cmbRicercaMov(1).ListIndex)
    End If
    
    If cmbRicercaMov(3).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Immobili.TipoUtenza ='" & Left$(cmbRicercaMov(3).Text, 1) & "'"
    End If
    
    If cmbRicercaMov(2).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Immobili.CodTariffa ='" & Left$(cmbRicercaMov(2).Text, 5) & "'"
    End If
    
    If cmbRicercaMov(4).ListIndex > 0 Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.flgRegistrato =" & IIf(cmbRicercaMov(4).ListIndex = 1, 0, -1)
    End If
    

    
    If Not ImgCombo2.SelectedItem Is Nothing Then
        If (Right$(ImgCombo2.SelectedItem.key, 2) <> "00") And Len(ImgCombo2.SelectedItem.key) >= 6 Then
            
            codCer = Mid$(ImgCombo2.SelectedItem.key, 3)
            codCategoria = Left$(codCer, 2)
            codSottoCategoria = Mid$(codCer, 3, 2)
            codCer = Right$(codCer, 2)
            
            
            If where <> "" Then where = where & " and "
            where = where & "Movimenti.CodCer ='" & codCer & "' and Movimenti.CodSottoCategoria='" & codSottoCategoria & "' and Movimenti.CodCategoria='" & codCategoria & "'"
        End If
    End If
    If cmbTargaMov(0).Text <> "Tutte" Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.Targa like '" & cmbTargaMov(0).Text & "%'"
    End If
    
    If cmbTargaMov(1).Text <> "Tutte" Then
        If where <> "" Then where = where & " and "
        where = where & "Movimenti.TargaTrasportatore like '" & cmbTargaMov(1).Text & "%'"
    End If
    
    
    If Not IsNull(DTPicker(0).Value) Then
        If where <> "" Then where = where & " and "
        where = where & "dataMovimento >=#" & Format$(DTPicker(0).Value, "yyyy-mm-dd") & " 00.00.00#"
    End If
    
    If Not IsNull(DTPicker(1).Value) Then
        If where <> "" Then where = where & " and "
        where = where & " dataMovimento <=#" & Format$(DTPicker(1).Value, "yyyy-mm-dd") & " 23.59.59#"
    End If
    
    
    
    Set rs = gDBC.Execute("SELECT trasportatori.RagioneSociale as RagioneSocialeT,Utenti.Cognome, Utenti.Nome,Movimenti.CodImmobile, Utenti.RagioneSociale,Movimenti.id,Movimenti.nrMovimento,Comuni.Comune, Immobili.Indirizzo, Immobili.NumCivico," & _
            "Immobili.Subalterno, Immobili.NumInterno, Immobili.Scala, Immobili.Piano, " & _
            "Immobili.CodComune,Immobili.CodTariffa,Immobili.TipoUtenza,(CodiciCer.CodCategoria + CodiciCer.CodSottoCategoria + CodiciCer.CodCer ) as CodiceCer, CodiciCer.Descrizione AS Rifiuto,Movimenti.Manuale, CodiciCer.FlgBonus, " & _
            "Movimenti.DataMovimento,Movimenti.targa,Movimenti.CodTrasportatore,Movimenti.TargaTrasportatore,Movimenti.NumeroFormulario,Movimenti.DataFormulario, Movimenti.Quantita,Movimenti.unitamisura, Movimenti.BonusDovuto, Movimenti.CodUser, Movimenti.DataIns, " & _
            "StatoMovimento.descrizione AS descrizioneStato,Movimenti.Stato AS Stato, Immobili.CodUtente,Causali.Descrizione as CausaleDes,Movimenti.flgRegistrato " & _
            "FROM (((((((Movimenti Left Outer JOIN Immobili ON Movimenti.CodImmobile = Immobili.CodImmobile and Movimenti.CodUtente = Immobili.CodUtente) " & _
            "INNER JOIN CodiciCer ON (Movimenti.CodCer = CodiciCer.CodCer) AND (Movimenti.CodSottoCategoria = CodiciCer.CodSottoCategoria) AND (Movimenti.CodCategoria = CodiciCer.CodCategoria)) " & _
            "INNER JOIN StatoMovimento ON Movimenti.Stato = StatoMovimento.codStato) " & _
            "Left outer JOIN Comuni ON Immobili.CodComune = Comuni.CODISTAT) " & _
            "Left Outer JOIN Utenti ON Utenti.CodUtente = Movimenti.CodUtente) " & _
            "Left outer JOIN trasportatori ON Movimenti.CodTrasportatore = trasportatori.idTrasportatore) " & _
            "INNER JOIN Causali ON Movimenti.codCausale = Causali.codCausale) " & _
            "WHERE " & where & " order by Immobili.codUtente,Movimenti.DataMovimento, iif (Utenti.Cognome = null ,Utenti.Cognome + Utenti.Nome, Utenti.RagioneSociale) ")
            

    lvwMovimenti.ListItems.Clear
    
    Do While Not rs.EOF
        Set Item = lvwMovimenti.ListItems.Add(, "_" & rs!id, nullToBlank(rs!codUtente))
        
        Item.ListSubItems.Add , , IIf(nullToBlank(rs!Cognome) = "", nullToBlank(rs!RagioneSociale), nullToBlank(rs!Cognome) & " " & nullToBlank(rs!Nome))
        Item.ListSubItems.Add , , nullToBlank(rs!codImmobile)
        Item.ListSubItems.Add , , IIf(IsNull(rs!Indirizzo), vbNullString, rs!Indirizzo & "," & nullToBlank(rs!NumCivico) & " Sub:" & nullToBlank(rs!Subalterno) & " Num.Int:" & nullToBlank(rs!NumInterno))
        Item.ListSubItems.Add , , nullToBlank(rs!tipoUtenza)
        Item.ListSubItems.Add , , nullToBlank(rs!codTariffa)
        Item.ListSubItems.Add , , nullToBlank(rs!Targa)
        Item.ListSubItems.Add , , rs!codiceCer
        Item.ListSubItems.Add , , rs!Rifiuto
        Item.ListSubItems.Add , , Format$(rs!DataMovimento, "dd-mmm-yyyy hh:nn:ss")
        Item.ListSubItems(Item.ListSubItems.Count).Tag = Format$(rs!DataMovimento, "dd/mm/yyyy hh:nn:ss")
        qty = rs!quantita.Value
        Item.ListSubItems.Add , , FormatNumber(qty, 2, , , vbFalse)
        Item.ListSubItems(Item.ListSubItems.Count).Tag = getNumber(qty)
        Item.ListSubItems.Add , , rs!UnitaMisura
        Item.ListSubItems.Add , , rs!BonusDovuto
        Item.ListSubItems(Item.ListSubItems.Count).Tag = getNumber(rs!BonusDovuto)
        Item.ListSubItems.Add , "Causale", rs!CausaleDes
        Item.ListSubItems.Add , "flgRegistrato", IIf(rs!flgRegistrato = -1, "Si", "No")
        Item.ListSubItems.Add , "Stato", rs!descrizioneStato
        
        
        Item.ListSubItems.Add , , nullToBlank(rs!CodTrasportatore)
        Item.ListSubItems.Add , , nullToBlank(rs!RagioneSocialeT)
        Item.ListSubItems.Add , , nullToBlank(rs!TargaTrasportatore)
        Item.ListSubItems.Add , "NumeroFormulario", nullToBlank(rs!NumeroFormulario)
        Item.ListSubItems.Add , , nullToBlank(rs!DataFormulario)
        Item.ListSubItems.Add , , IIf(rs!manuale = -1, "Manuale", "Automatico")
        Item.ListSubItems.Add , , rs!codUser
        Item.ListSubItems.Add , , Format$(rs!dataIns, "dd-mmm-yyyy hh:nn:ss")
        Item.ListSubItems(Item.ListSubItems.Count).Tag = Format$(rs!dataIns, "dd/mm/yyyy hh:nn:ss")
        
        If (rs!Stato = "2") Then coloraRiga Item, vbRed
        
        rs.MoveNext
    Loop
    
    rs.Close
    AllargaLeColonneLVW lvwMovimenti, False
    Screen.MousePointer = vbDefault
End Sub

Private Sub caricaImmobili(codUtente As String)
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    
    lvwImmobili.ListItems.Clear
    
    If codUtente = vbNullString Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    Set rs = gDBC.Execute("SELECT Comuni.comune, Tariffe.Descrizione AS Tariffa, Immobili.* " & _
            "FROM (immobili INNER JOIN Comuni ON immobili.CodComune = Comuni.CODISTAT) INNER JOIN Tariffe ON immobili.CodTariffa = Tariffe.CodTariffa " & _
            "WHERE  DATE() between DataInizioValidita and DataFineValidita and   immobili.Tipo<>'Z' AND immobili.[codUtente]=" & codUtente & " " & _
            "ORDER BY immobili.FlgUtenzaDomesticaResidente")

    Do While Not rs.EOF
        Set Item = lvwImmobili.ListItems.Add(, "_" & rs!codImmobile, rs!codImmobile)
        Item.ListSubItems.Add , rs!Tipo.Name, rs!Tipo.Value
        Item.ListSubItems.Add , rs!comune.Name, rs!comune.Value
        Item.ListSubItems.Add , rs!tipoUtenza.Name, rs!tipoUtenza.Value
        Item.ListSubItems.Add , rs!Indirizzo.Name, rs!Indirizzo.Value & "," & rs!NumCivico.Value
        Item.ListSubItems.Add , rs!Subalterno.Name, nullToBlank(rs!Subalterno.Value)
        Item.ListSubItems.Add , rs!NumInterno.Name, nullToBlank(rs!NumInterno.Value)
        Item.ListSubItems.Add , rs!codTariffa.Name, nullToBlank(rs!codTariffa.Value)
        Item.ListSubItems.Add , rs!Tariffa.Name, nullToBlank(rs!Tariffa.Value)
        
        If (rs!flgUtenzaDomesticaResidente Or lvwImmobili.ListItems.Count = 1) Then
            lvwImmobili.Tag = ""
            lvwImmobili_ItemCheck Item
            
            Item.Checked = True
            lvwImmobili.Tag = Item.key
        End If
        
        rs.MoveNext
    Loop
    
    AllargaLeColonneLVW lvwImmobili, False
    Screen.MousePointer = vbDefault
    lvwImmobili.ColumnHeaders(2).Width = 0
    
    
    caricaComboRifiuti imgCombo1, Combo2(0)
    
    
    rs.Close
    Set rs = Nothing
End Sub



Private Sub caricaComboRifiuti(imgCombo As ImageCombo, cmb As ComboBox)
    Dim primaVolta As Boolean
    Dim keyCat As String
    Dim keySottCat As String
    Dim key As String
    Dim keyPrefisso As String
    Dim sWhere As String
    
    
    'filtro su combo rifiuti
    If imgCombo.Name = "imgCombo1" Then
        sWhere = ""
        If txt(TXT_FINDCODCER).Text <> "" Then
            If IsNumeric(txt(TXT_FINDCODCER)) Then
                sWhere = " Codice like '" & txt(TXT_FINDCODCER).Text & "%'"
            Else
                sWhere = " Descrizione like '" & txt(TXT_FINDCODCER).Text & "%'"
            End If
        End If
        sWhere = IIf(sWhere <> "", sWhere & " and ", sWhere) & IIf(mTipoUtenza = "D", "  flgAmmessoPerUD= true ", " flgAmmessoPerUA= true ")
    Else
        sWhere = ""
        If TxtMov(TXT_FINDCODCERMOV).Text <> "" Then
            If IsNumeric(txt(TXT_FINDCODCERMOV)) Then
                sWhere = " Codice like '" & TxtMov(TXT_FINDCODCERMOV).Text & "%'"
            Else
                sWhere = " Descrizione like '" & TxtMov(TXT_FINDCODCERMOV).Text & "%'"
            End If
        End If
    End If
    
    
    cmb.Clear
    If (rsCodiciCER Is Nothing) Then
        Set rsCodiciCER = CreateObject("ADODB.Recordset")
        rsCodiciCER.LockType = adLockBatchOptimistic
        rsCodiciCER.CursorLocation = adUseClient
        rsCodiciCER.CursorType = adOpenStatic
        
        rsCodiciCER.Source = "SELECT CodiciCer.*, (CodiciCer.CodCategoria + CodiciCer.CodSottoCategoria + CodiciCer.CodCer) as Codice,  CategoriaCER.Descrizione AS DesCategoria, SottoCategoriaCER.Descrizione AS DesSottoCategoria " & _
                    "FROM (CategoriaCER INNER JOIN SottoCategoriaCER ON CategoriaCER.CodCategoria = SottoCategoriaCER.CodCategoria) INNER JOIN CodiciCer ON (SottoCategoriaCER.CodSottocategoria = CodiciCer.CodSottoCategoria) AND (SottoCategoriaCER.CodCategoria = CodiciCer.CodCategoria) " & _
                    "WHERE CodiciCer.FlgAbilitatoInRicerca=True  " & _
                    "ORDER BY CodiciCer.FlgBonus  , CategoriaCER.Descrizione, SottoCategoriaCER.Descrizione, CodiciCer.Descrizione "
    
        
        Set rsCodiciCER.ActiveConnection = gDBC
        rsCodiciCER.Open
        Set rsCodiciCER.ActiveConnection = Nothing
    Else
        rsCodiciCER.Filter = vbNullString
        If sWhere <> vbNullString Then rsCodiciCER.Filter = sWhere
    End If
    
    primaVolta = False
    keyPrefisso = "_B"
    imgCombo.ComboItems.Clear
    imgCombo.ComboItems.Add , keyPrefisso, "Con Bonus", , , 0
    
    
    
    Do While Not rsCodiciCER.EOF
        
        If (rsCodiciCER!FlgBonus = False And Not primaVolta) Then
            keyPrefisso = "_N"
            imgCombo.ComboItems.Add , "ALTRI", "Altri", , , 0
            primaVolta = True
        
        End If
        
        
        
        If (keyCat <> keyPrefisso & rsCodiciCER!codCategoria & "00" & "00") Then
            key = keyPrefisso & rsCodiciCER!codCategoria & "00" & "00"
            'Combo1.ComboItems.Add , key, rsCodiciCer!DesCategoria & Space(5) & "[" & rsCodiciCer!codCategoria & "]", , , 3
            keyCat = key
            keySottCat = ""
        End If
        
        If (keySottCat <> keyPrefisso & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & "00") Then
            key = keyPrefisso & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & "00"
            'Combo1.ComboItems.Add , key, rsCodiciCer!DesSottoCategoria & Space(5) & "[" & rsCodiciCer!codSottoCategoria & "]", , , 4
            keySottCat = key
        End If
        
        key = keyPrefisso & rsCodiciCER!Codice
        'codice CER
        imgCombo.ComboItems.Add , key, rsCodiciCER!descrizione & Space(5) & "[" & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & rsCodiciCER!codCer & "]", IIf(rsCodiciCER!RifiutoPericoloso, 1, 2), , 9
        'corrispettivo bonus
        imgCombo.ComboItems(imgCombo.ComboItems.Count).Tag = rsCodiciCER!CorrispettivoBonus & "|" & _
            rsCodiciCER!UnitaDiMisura & "|" & rsCodiciCER!FlgAbilitatoFatturazione & "|" & _
            toZero(rsCodiciCER!qtyMaxFatturabile) & "|" & toZero(rsCodiciCER!qtyMaxGGUD) & "|" & _
            toZero(rsCodiciCER!qtyMaxAAUD) & "|" & toZero(rsCodiciCER!qtyMaxGGUA) & "|" & _
            toZero(rsCodiciCER!qtyMaxAAUA) & "|" & rsCodiciCER!flgAmmessoPerUD & "|" & rsCodiciCER!flgAmmessoPerUA
        
        
         
        cmb.AddItem rsCodiciCER!descrizione & Space(5) & "[" & rsCodiciCER!codCategoria & rsCodiciCER!codSottoCategoria & rsCodiciCER!codCer & "]"
        cmb.ItemData(cmb.ListCount - 1) = imgCombo.ComboItems.Count
        
        
        rsCodiciCER.MoveNext
    Loop
    
End Sub



Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelY
    
    
    resizeControl
End Sub

Private Sub resizeControl()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).Width - offsetX
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).Width
        
        cmd(i).Move posLeft, posTop
    Next
    
    TabStrip.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmd(0).Height - 2 * offsetX
    
    posTop = TabStrip.ClientTop
    posLeft = TabStrip.ClientLeft
    
    For i = picContainer.LBound To picContainer.ubound
        picContainer(i).Move posLeft, posTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    Next i
    
    ProgressBar1.Top = Me.ScaleHeight - ProgressBar1.Height - offsetY
    lblEsportazione.Top = ProgressBar1.Top - lblEsportazione.Height - offsetY
    
    PicTotali.Move offsetX, Me.ScaleHeight - PicTotali.ScaleHeight
End Sub


Public Sub showMe(key As String)
    Load Me
    
    TabStrip.Tabs(key).Selected = True
    
    Set m_cHdrIcons.ListView = lvwMovimenti
    lvwMovimenti.SortOrder = lvwAscending
    lvwMovimenti.SortKey = 1
    Call m_cHdrIcons.SetHeaderIcons(lvwMovimenti.SortKey, lvwMovimenti.SortOrder)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Esci = True
    
    If Not rsCodiciCER Is Nothing Then
        rsCodiciCER.Close
        Set rsCodiciCER = Nothing
    End If
    
    If Not rsCliente Is Nothing Then
        rsCliente.Close
        Set rsCliente = Nothing
    End If
    
    Set pfrmPesata = Nothing
End Sub

Private Sub ImgCombo2_Change()
Dim l As Long
Static Changing As Boolean
Static Found As Boolean
Static PartialText As String



If Changing Then Exit Sub Else Changing = True ' questo evita le chiamate nidificate della routine

If ImgCombo2.Text = "" Then
    Combo2(1).Text = ""
    GoTo Esci ' se non vi è testo esce dalla routine
End If

Combo2(1).Text = ImgCombo2.Text
PartialText = ImgCombo2.Text      'memorizza la parte di testo digitata

If Canceling Then GoTo Esci     ' se il tasto Cancel è premuto esce dalla routine

If PartialText = "" Then GoTo Esci

l = SendMessage(Combo2(1).hWnd, CB_FINDSTRING, -1, ByVal PartialText)    ' trova l'indice per l'elemento più simile al quello digitato


If l > -1 Then   ' se l'elemento viene trovato
    ' il testo parziale viene completato con quello trovato dall'API
    Combo2(1).Text = PartialText & Mid(Combo2(1).List(l), Len(PartialText) + 1)
    'ImageCombo1.Text = Combo2.Text
    ImgCombo2.ComboItems(Combo2(1).ItemData(l)).Selected = True
    If Not Deleting Then
        Combo2(1).SelStart = Len(PartialText)
    Else
        Combo2(1).SelStart = Len(PartialText) - IIf(Found, 1, 0)
    End If
    
    
    Combo2(1).SelLength = Len(Combo2(1).Text)
    
    ImgCombo2.SelStart = Combo2(1).SelStart
    ImgCombo2.SelLength = Combo2(1).SelLength
    
    DoEvents
    
    PartialText = Mid(Combo2(1).Text, 1, Combo2(1).SelStart)
    If PartialText = "" Then ImgCombo2.Text = "": Combo2(1).Text = ""
    Deleting = False
    Found = True
Else
    Found = False
    
  
End If

Esci:
    Changing = False
    Deleting = False
    Canceling = False
End Sub

Private Sub lvwImmobili_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    If lvwImmobili.Tag <> vbNullString And lvwImmobili.ListItems.Count > 1 Then lvwImmobili.ListItems(lvwImmobili.Tag).Checked = False
        
    lvwImmobili.Tag = Item.key
    mTipoUtenza = Item.ListSubItems("TipoUtenza").Text
    
    'txt(TXT_COD_TRASPORTATORE).enabled = mTipoUtenza <> "D"
    'txt(TXT_NRFORMULARIO).enabled = mTipoUtenza <> "D"
    'dtpickFormulario.enabled = mTipoUtenza <> "D"
    'cmbTarga(1).enabled = mTipoUtenza <> "D"
    'cmbRicerca(1).enabled = mTipoUtenza <> "D"
    
    txt(TXT_COD_TRASPORTATORE).enabled = True
    txt(TXT_NRFORMULARIO).enabled = True
    dtpickFormulario.enabled = True
    cmbTarga(1).enabled = True
    cmbRicerca(1).enabled = True
    
    Item.Selected = True
    
    caricaTotali Mid$(Item.key, 2)
    
End Sub

Private Sub lvwMovimenti_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwMovimenti, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub lvwMovimenti_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmd(CMD_CANCELLAMOV).enabled = Item.ListSubItems("Stato").Text <> "Contabilizzato" And Item.ListSubItems("flgRegistrato").Text <> "Si"





End Sub

Private Sub picContainer_Resize(Index As Integer)
    Dim posY As Long
    Dim posX As Long
    On Error Resume Next
    Select Case Index
        Case 0
             
            posX = Frame1.Left
            posY = Frame1.Top + offsetY + Frame1.Height
            

            
           ' fmeUtente.Width = Me.ScaleWidth
             Label1(LBL_IMMOBILI).Move lvwImmobili.Left, fmeUtente.Top
            lvwImmobili.Move fmeUtente.Left + fmeUtente.Width + offsetX, Label1(LBL_IMMOBILI).Top + offsetY * 2, Me.ScaleWidth - (fmeUtente.Left + fmeUtente.Width + 2 * offsetX), fmeUtente.Height - Label1(LBL_IMMOBILI).Height - offsetY
            
           
                        
         
            'fmeTotali.Move Frame1.Left, Frame1.Top + Frame1.Height + offsetY, Me.ScaleWidth, picContainer(Index).ScaleHeight - (Frame1.Top + Frame1.Height)
            'lvwTotMovimenti.Move offsetX, 2 * offsetY, fmeTotali.Width - 2 * offsetX, fmeTotali.Height - 3 * offsetY
            fmeTotali.Move Frame1.Left + Frame1.Width + offsetX, Frame1.Top, Me.ScaleWidth - (Frame1.Left + Frame1.Width + 2 * offsetX), picContainer(Index).ScaleHeight - (Frame1.Top)
            lvwTotMovimenti.Move offsetX, 2 * offsetY, fmeTotali.Width - 2 * offsetX, fmeTotali.Height - 3 * offsetY
            
            imgCombo1.Width = picContainer(0).ScaleWidth - imgCombo1.Left - 2 * offsetX
        Case 1
            posY = cmbRicercaMov(2).Top + offsetY + cmbRicercaMov(2).Height
            
            cmbRicercaMov(2).Width = picContainer(1).ScaleWidth - cmbRicercaMov(2).Left - 2 * offsetX
            
            ImgCombo2.Width = picContainer(1).ScaleWidth - ImgCombo2.Left - 2 * offsetX
            
            lvwMovimenti.Move offsetX, posY, picContainer(Index).ScaleWidth - 2 * offsetX, picContainer(Index).ScaleHeight - posY
    End Select
    
End Sub

Private Sub TabStrip_Click()
    
    If (TabStrip.Tag <> vbNullString) Then
        picContainer(TabStrip.Tag).Visible = False
    End If
    
    picContainer(TabStrip.SelectedItem.Index - 1).Visible = True
    TabStrip.Tag = TabStrip.SelectedItem.Index - 1
    
    cmd(CMD_RICERCAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_RISTAMPAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_ESPORTAMOV).enabled = TabStrip.Tag = 1
    cmd(CMD_CANCELLAMOV).enabled = TabStrip.Tag = 1
    
    PicTotali.Visible = TabStrip.Tag = 0
    
    
    If cmd(CMD_RICERCAMOV).enabled Then
        cmd(CMD_LETTURAQUANTITA).Tag = cmd(CMD_LETTURAQUANTITA).enabled
        cmd(CMD_SALVAESTAMPA).Tag = cmd(CMD_SALVAESTAMPA).enabled
        cmd(CMD_SALVAMOV).Tag = cmd(CMD_SALVAMOV).enabled
        cmd(CMD_LETTURAQUANTITA).enabled = False
        cmd(CMD_SALVAESTAMPA).enabled = False
        cmd(CMD_SALVAMOV).enabled = False
        
        If txt(TXT_COD_UTENTE).Text <> "" Then
            TxtMov(0).Text = txt(TXT_COD_UTENTE).Text
            cmdRicercaUtenteMov(0).Tag = "1"
            cmdRicercaUtenteMov_Click (0)
        End If
    ElseIf cmd(CMD_LETTURAQUANTITA).Tag <> "" Then
        cmd(CMD_LETTURAQUANTITA).enabled = cmd(CMD_LETTURAQUANTITA).Tag
        cmd(CMD_SALVAESTAMPA).enabled = cmd(CMD_SALVAESTAMPA).Tag
        cmd(CMD_SALVAMOV).enabled = cmd(CMD_SALVAMOV).Tag
    
    End If
    
End Sub

Private Sub letturaPeso(Index As Integer)
    Dim peso As Double
    Dim messaggio As String
    
    peso = 0
    Do While peso = 0 And Not Esci
        Select Case pesa(Index).ReadMessage(messaggio)
            Case ErroreRicevuto
            Case inlettura
            Case LetturaTerminata
                'messaggio = "1402/07/200815:17:1702/07/200815:17:34   144   145    200101       CARTA CARTONE    200101       CARTA CARTONE   1170007230010                         3.25 kg                             0.00 kg        0 kg        0 kg     3,25 kg21"
                msgPesa = getMessaggioPesata(messaggio)
                impostaDatiDaBadge msgPesa
         End Select
         DoEvents
    Loop
    
    If Esci Then peso = 0
    
        
End Sub

Private Sub impostaDatiDaBadge(msgPesa As messaggioPesa)
    If (msgPesa.PesoNetto) > 0 Then
        txt(TXT_COD_UTENTE).Text = getCodiceUtenteDaBadge(msgPesa.codBadge)
        If txt(TXT_COD_UTENTE) = "" Then
            MsgBox "attenzione utente non trovato per codice:" & msgPesa.codBadge
        Else
            caricaDatiUtente txt(TXT_COD_UTENTE).Text
        End If
        txt(TXT_QUANTITA).Text = msgPesa.PesoNetto
        cmd(CMD_SALVAESTAMPA).enabled = True
        cmd(CMD_SALVAMOV).enabled = True
        'imposta il codice cer rilevato
        txt(TXT_FINDCODCER).Text = ""
        txt(TXT_FINDCODCER).Text = IIf(Trim$(msgPesa.CodRifiuto) = "999999", "", Trim$(msgPesa.CodRifiuto))
        'imgCombo1.ComboItems("_" & Trim$(msgPesa.CodRifiuto)).Selected = True
        If Trim$(msgPesa.CodRifiuto) = "999999" Then
            imgCombo1.Text = ""
            On Error Resume Next
            imgCombo1.SetFocus
        End If
        Esci = True
        
        cmd_Click CMD_LETTURAQUANTITA
    End If
End Sub

Private Sub imgCombo1_change()
Dim l As Long
Static Changing As Boolean
Static Found As Boolean
Static PartialText As String

Label1(LBL_FATTURABILE).Visible = False

If Changing Then Exit Sub Else Changing = True ' questo evita le chiamate nidificate della routine

If imgCombo1.Text = "" Then
    Combo2(0).Text = ""
    GoTo Esci ' se non vi è testo esce dalla routine
End If

Combo2(0).Text = imgCombo1.Text
PartialText = imgCombo1.Text      'memorizza la parte di testo digitata

If Canceling Then GoTo Esci     ' se il tasto Cancel è premuto esce dalla routine

If PartialText = "" Then GoTo Esci

l = SendMessage(Combo2(0).hWnd, CB_FINDSTRING, -1, ByVal PartialText)    ' trova l'indice per l'elemento più simile al quello digitato


If l > -1 Then   ' se l'elemento viene trovato
    ' il testo parziale viene completato con quello trovato dall'API
    Combo2(0).Text = PartialText & Mid(Combo2(0).List(l), Len(PartialText) + 1)
    'imgcombo1.Text = combo2(0).Text
    imgCombo1.ComboItems(Combo2(0).ItemData(l)).Selected = True
    If Not Deleting Then
        Combo2(0).SelStart = Len(PartialText)
    Else
        Combo2(0).SelStart = Len(PartialText) - IIf(Found, 1, 0)
    End If
    
    
    Combo2(0).SelLength = Len(Combo2(0).Text)
    
    imgCombo1.SelStart = Combo2(0).SelStart
    imgCombo1.SelLength = Combo2(0).SelLength
    
    DoEvents
    
    imgCombo1_Click
    
    PartialText = Mid(Combo2(0).Text, 1, Combo2(0).SelStart)
    If PartialText = "" Then imgCombo1.Text = "": Combo2(0).Text = ""
    Deleting = False
    Found = True
Else
    Found = False
    
  
End If

Esci:
    Changing = False
    Deleting = False
    Canceling = False

End Sub



Private Sub imgCombo1_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then Canceling = True
If KeyCode = vbKeyBack And imgCombo1.Text <> "" Then Deleting = True: KeyCode = 0
End Sub

Private Sub imgCombo2_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then Canceling = True
If KeyCode = vbKeyBack And imgCombo1.Text <> "" Then Deleting = True: KeyCode = 0
End Sub



Private Sub txt_Change(Index As Integer)
    Dim cmbItem As ComboItem
    Dim sWhere As String
    
    If Index = TXT_FINDCODCER Then
    
        caricaComboRifiuti imgCombo1, Combo2(0)
        'mi posiziono sul primo rifiuto
        For Each cmbItem In imgCombo1.ComboItems
            If (Right$(cmbItem.key, 2) <> "00") And Len(cmbItem.key) >= 6 Then
                cmbItem.Selected = True
                Exit Sub
            End If
        Next
    ElseIf Index = TXT_COD_UTENTE Then
        
    End If
    
    
    
End Sub


Public Sub applicaRestrizioniUtente()
    'txt(TXT_COD_UTENTE).Enabled = frmLogin.Livello <> LVL_USER
    'txt(TXT_FINDCODCER).Enabled = frmLogin.Livello <> LVL_USER
    'txt(TXT_QUANTITA).Enabled = frmLogin.Livello <> LVL_USER
    
    'imgCombo1.Enabled = frmLogin.Livello <> LVL_USER
    'cmbRicerca(0).Enabled = frmLogin.Livello <> LVL_USER
    'cmd(CMD_SALVA).Enabled = frmLogin.Livello <> LVL_USER
End Sub

Function getPropietaCER(sp As String, proprieta As ProprietaCER) As String
    'CorrispettivoBonus|UnitaDiMisura|FlgAbilitatoFatturazione|qtyMaxFatturabile|qtyMaxGGUD|qtyMaxAAUD|qtyMaxGGUA|qtyMaxAAUA
    Dim s() As String
    
    s = Split(sp, "|")
    
    Select Case proprieta
        Case ProprietaCER.CorrispettivoBonus
            getPropietaCER = s(0)
        Case ProprietaCER.UnitaDiMisura
            getPropietaCER = s(1)
        Case ProprietaCER.FlgAbilitatoFatturazione
            getPropietaCER = s(2)
        Case ProprietaCER.qtyMaxYYUA
            getPropietaCER = s(7)
        Case ProprietaCER.qtyMaxYYUD
            getPropietaCER = s(5)
        Case ProprietaCER.qtyMaxGGUA
            getPropietaCER = s(6)
        Case ProprietaCER.qtyMaxGGUD
            getPropietaCER = s(4)
        Case ProprietaCER.qtyMaxFatturabile
            getPropietaCER = s(3)
        Case ProprietaCER.flgAmmessoPerUD
            getPropietaCER = s(8)
        Case ProprietaCER.flgAmmessoPerUA
            getPropietaCER = s(9)
    
    End Select
End Function

Function getTotaleMovimentiAnno(codiceCer As String, codImmobile As Long) As Double
    Dim rs As ADODB.Recordset
    Dim codCategoria As String, codSottoCategoria As String, codCer As String
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = Left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codCer = Right$(codCer, 2)

    Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                "From Movimenti " & _
                "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                " and  Movimenti.DataMovimento Between dateadd('d', datediff('d',  date()  ,year(date()) &'-01-01' ) , date()) +'00.00.00' And Now() AND  Movimenti.CodImmobile=" & codImmobile & _
                " AND Movimenti.CodCategoria='" & codCategoria & "' AND Movimenti.CodSottoCategoria='" & codSottoCategoria & "'  AND Movimenti.CodCer='" & codCer & "'")


    
    getTotaleMovimentiAnno = 0
    If Not rs.EOF Then
        getTotaleMovimentiAnno = 0
        If Not IsNull(rs(0)) Then
            getTotaleMovimentiAnno = rs(0)
        End If
        rs.Close
    End If
End Function

Function getTotaleMovimentiGiorno(codiceCer As String, codImmobile As Long) As Double
    Dim rs As ADODB.Recordset
    Dim codCategoria As String, codSottoCategoria As String, codCer As String
    
    codCer = Mid$(imgCombo1.SelectedItem.key, 3)
    codCategoria = Left$(codCer, 2)
    codSottoCategoria = Mid$(codCer, 3, 2)
    codCer = Right$(codCer, 2)

    Set rs = gDBC.Execute("SELECT Sum(Movimenti.quantita) AS quantita " & _
                "From Movimenti   " & _
                "WHERE Stato<>'2' and (codCausale='00' or codCausale='" & gCodCausaleMovEcoPuntoIN & "' or codCausale='" & gCodCausaleMovManuale & "')" & _
                " and Movimenti.DataMovimento Between Date()+'00.00.00' And Now() AND  Movimenti.CodImmobile=" & codImmobile & _
                " AND Movimenti.CodCategoria='" & codCategoria & "' AND Movimenti.CodSottoCategoria='" & codSottoCategoria & "'  AND Movimenti.CodCer='" & codCer & "'")


    
    getTotaleMovimentiGiorno = 0
    If Not rs.EOF Then
        getTotaleMovimentiGiorno = 0
        If Not IsNull(rs(0)) Then
            getTotaleMovimentiGiorno = rs(0)
        End If
        rs.Close
    End If
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case TXT_COD_TRASPORTATORE, TXT_COD_UTENTE
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
        
        Case TXT_QUANTITA
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete _
                And KeyAscii <> Asc(",") Then
                KeyAscii = 0
            End If
        
    End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If (Index = TXT_COD_UTENTE) Then caricaDatiUtente txt(TXT_COD_UTENTE)
    If (Index = TXT_COD_TRASPORTATORE) Then caricaDatiTrasportatore txt(TXT_COD_TRASPORTATORE)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case TXT_QUANTITA, TXT_COD_UTENTE
            If txt(Index) <> "" And Not IsNumeric(txt(Index)) Then
                MsgBox "Il valore inserito deve essere numerico", vbExclamation + vbOKOnly, App.Title
                Cancel = True
            End If
    End Select
End Sub

Private Sub TxtMov_Change(Index As Integer)
    Dim cmbItem As ComboItem
    Dim sWhere As String
    
    If Index = TXT_FINDCODCERMOV Then
    
        caricaComboRifiuti ImgCombo2, Combo2(1)
        'mi posiziono sul primo rifiuto
        For Each cmbItem In ImgCombo2.ComboItems
            If (Right$(cmbItem.key, 2) <> "00") And Len(cmbItem.key) >= 6 Then
                cmbItem.Selected = True
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub TxtMov_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 2
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack _
                And KeyAscii <> vbKeyCancel _
                And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub TxtMov_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If TxtMov(Index) <> "" And Not IsNumeric(TxtMov(Index)) Then
                MsgBox "Il valore inserito deve essere numrico", vbExclamation + vbOKOnly, App.Title
                Cancel = True
            End If
    End Select

End Sub
