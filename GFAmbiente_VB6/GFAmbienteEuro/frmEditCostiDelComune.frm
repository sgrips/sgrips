VERSION 5.00
Begin VB.Form frmEditCostiDelComune 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costi del Comune"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Dettaglio"
      Height          =   375
      Index           =   2
      Left            =   9165
      TabIndex        =   67
      Top             =   10335
      Width           =   1215
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
      Index           =   11
      Left            =   4515
      TabIndex        =   9
      Top             =   8010
      Width           =   2300
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
      Index           =   10
      Left            =   4515
      TabIndex        =   8
      Top             =   7440
      Width           =   2300
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
      Left            =   10920
      TabIndex        =   24
      Top             =   9120
      Width           =   2300
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
      Index           =   16
      Left            =   10920
      TabIndex        =   13
      Top             =   2775
      Width           =   2300
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
      Index           =   17
      Left            =   10920
      TabIndex        =   14
      Top             =   3375
      Width           =   2300
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
      Index           =   18
      Left            =   10920
      TabIndex        =   15
      Top             =   3960
      Width           =   2300
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
      Index           =   15
      Left            =   10920
      TabIndex        =   16
      Top             =   4545
      Width           =   2300
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
      Index           =   20
      Left            =   10920
      TabIndex        =   17
      Top             =   5115
      Width           =   2300
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
      Index           =   21
      Left            =   10920
      TabIndex        =   18
      Top             =   5700
      Width           =   2300
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
      Index           =   22
      Left            =   10920
      TabIndex        =   23
      Top             =   8565
      Width           =   2300
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
      Index           =   23
      Left            =   10920
      TabIndex        =   19
      Top             =   6270
      Width           =   2300
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
      Index           =   24
      Left            =   10920
      TabIndex        =   20
      Top             =   6840
      Width           =   2300
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
      Index           =   25
      Left            =   10920
      TabIndex        =   21
      Top             =   7425
      Width           =   2300
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
      Index           =   26
      Left            =   10920
      TabIndex        =   22
      Top             =   7995
      Width           =   2300
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
      Index           =   14
      Left            =   4515
      TabIndex        =   12
      Top             =   9735
      Width           =   2300
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
      Index           =   13
      Left            =   4515
      TabIndex        =   11
      Top             =   9165
      Width           =   2300
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
      Index           =   12
      Left            =   4515
      TabIndex        =   10
      Top             =   8580
      Width           =   2300
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
      Index           =   9
      Left            =   4515
      TabIndex        =   7
      Top             =   6855
      Width           =   2300
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
      Index           =   8
      Left            =   4515
      TabIndex        =   6
      Top             =   6270
      Width           =   2300
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
      Index           =   7
      Left            =   4515
      TabIndex        =   5
      Top             =   5700
      Width           =   2300
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
      Index           =   6
      Left            =   4515
      TabIndex        =   4
      Top             =   5115
      Width           =   2300
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
      Index           =   5
      Left            =   4515
      TabIndex        =   3
      Top             =   4545
      Width           =   2300
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
      Index           =   4
      Left            =   4515
      TabIndex        =   2
      Top             =   3960
      Width           =   2300
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
      Left            =   4515
      TabIndex        =   1
      Top             =   3375
      Width           =   2300
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
      Index           =   2
      Left            =   4515
      TabIndex        =   0
      Top             =   2775
      Width           =   2300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   0
      Left            =   12015
      TabIndex        =   38
      Top             =   10335
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   1
      Left            =   10575
      TabIndex        =   37
      Top             =   10335
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   60
      TabIndex        =   31
      Top             =   -15
      Width           =   13260
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   10425
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   975
         Width           =   2670
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   28
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   495
         Width           =   2670
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   27
         Left            =   10425
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   495
         Width           =   2670
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   29
         Left            =   10770
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Costo In Euro"
         Top             =   1860
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   345
         Index           =   0
         Left            =   10440
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PV:80%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   6600
         TabIndex        =   76
         Top             =   2040
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PF:80%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4950
         TabIndex        =   75
         Top             =   2040
         Width           =   1545
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   9600
         TabIndex        =   70
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   28
         Left            =   7635
         TabIndex        =   69
         Top             =   195
         Width           =   1260
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   27
         Left            =   10410
         TabIndex        =   68
         Top             =   195
         Width           =   1605
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EUR"
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
         Index           =   29
         Left            =   12555
         TabIndex        =   40
         Top             =   1860
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tariffa Riferimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   8100
         TabIndex        =   39
         Top             =   1440
         Width           =   2340
      End
      Begin VB.Label lblCampi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "01/01/1998"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   505
         Left            =   3015
         TabIndex        =   30
         Top             =   825
         Width           =   1245
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   504
         Left            =   3000
         TabIndex        =   29
         Top             =   1755
         Width           =   435
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   503
         Left            =   1320
         TabIndex        =   28
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodIstat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   501
         Left            =   1320
         TabIndex        =   26
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Index           =   500
         Left            =   240
         TabIndex        =   25
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
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
         Index           =   100
         Left            =   240
         TabIndex        =   36
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodIstat:"
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
         Index           =   210
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAP:"
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
         Index           =   220
         Left            =   2400
         TabIndex        =   34
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
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
         Index           =   506
         Left            =   2415
         TabIndex        =   33
         Top             =   825
         Width           =   585
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regione:"
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
         Index           =   240
         Left            =   240
         TabIndex        =   32
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label lblCampi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   502
         Left            =   1320
         TabIndex        =   27
         Top             =   1305
         Width           =   900
      End
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   11
      Left            =   45
      TabIndex        =   66
      Top             =   8040
      Width           =   480
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTR_RICAVI"
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
      Left            =   60
      TabIndex        =   65
      Top             =   7470
      Width           =   1440
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CK"
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
      Left            =   6900
      TabIndex        =   64
      Top             =   9180
      Width           =   315
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARC"
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
      Index           =   16
      Left            =   6915
      TabIndex        =   63
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGG"
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
      Left            =   6915
      TabIndex        =   62
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CCD"
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
      Left            =   6900
      TabIndex        =   61
      Top             =   4020
      Width           =   510
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CC"
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
      Index           =   15
      Left            =   6900
      TabIndex        =   60
      Top             =   4605
      Width           =   330
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AMM"
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
      Left            =   6915
      TabIndex        =   59
      Top             =   5235
      Width           =   555
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACC"
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
      Left            =   6915
      TabIndex        =   58
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
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
      Left            =   6900
      TabIndex        =   57
      Top             =   8625
      Width           =   180
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RN"
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
      Left            =   6900
      TabIndex        =   56
      Top             =   6300
      Width           =   345
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KN-1"
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
      Left            =   6915
      TabIndex        =   55
      Top             =   6960
      Width           =   525
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN"
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
      Left            =   6915
      TabIndex        =   54
      Top             =   7530
      Width           =   240
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FN"
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
      Left            =   6900
      TabIndex        =   53
      Top             =   8115
      Width           =   315
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGD"
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
      Index           =   14
      Left            =   45
      TabIndex        =   52
      Top             =   9735
      Width           =   540
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
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
      Left            =   60
      TabIndex        =   51
      Top             =   9150
      Width           =   165
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CON"
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
      Left            =   60
      TabIndex        =   50
      Top             =   8610
      Width           =   510
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTR_COSTI"
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
      Left            =   45
      TabIndex        =   49
      Top             =   6885
      Width           =   1335
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRD"
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
      Left            =   45
      TabIndex        =   48
      Top             =   6270
      Width           =   525
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CGIND"
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
      Left            =   60
      TabIndex        =   47
      Top             =   5685
      Width           =   780
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AC"
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
      Left            =   60
      TabIndex        =   46
      Top             =   5115
      Width           =   330
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTS"
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
      Left            =   45
      TabIndex        =   45
      Top             =   4545
      Width           =   465
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRT"
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
      Left            =   45
      TabIndex        =   44
      Top             =   3975
      Width           =   480
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CSL"
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
      Left            =   60
      TabIndex        =   43
      Top             =   3390
      Width           =   465
   End
   Begin VB.Label lblCampi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CG"
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
      Index           =   2
      Left            =   60
      TabIndex        =   42
      Top             =   2820
      Width           =   360
   End
End
Attribute VB_Name = "frmEditCostiDelComune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lvw As ListView
Private frm As Form
Private OldTxtValue As String
Private CostoSave As Double
Private flgSimulazione As Boolean


Const TOT_SERVIZIO_ID = 0
Const GETTITO_ID = 1
Const CG_ID = 2
Const CSL_ID = 3
Const CRT_ID = 4
Const CTS_ID = 5
Const AC_ID = 6
Const CGIND_ID = 7
Const CRD_ID = 8
Const CTR_COSTI_ID = 9
Const CTR_RICAVI_ID = 10
Const CTR_ID = 11
Const CON_ID = 12
Const Y_ID = 13
Const CGD_ID = 14
Const CC_ID = 15
Const CARC_ID = 16
Const CGG_ID = 17
Const CCD_ID = 18
Const CK_ID = 19
Const AMM_ID = 20
Const ACC_ID = 21
Const R_ID = 22
Const RN_ID = 23
Const KN_ID = 24
Const IN_ID = 25
Const FN_ID = 26
Const PARTE_VARIABILE_ID = 27
Const PARTE_FISSA_ID = 28
Const EURO_ID = 29

'Command
Const CMD_DETTAGLIO = 2
Const CMD_CONFERMA = 1
Const CMD_ANNULLA = 0

Const COMUNE_ID = 500
Const ISTAT_ID = 501
Const REGIONE_ID = 502
Const PROVINCIA_ID = 503
Const CAP_ID = 504
Const DATA_ID = 505

'Tipo del Dettaglio
Private mDettaglio As Integer


'TAG
Private Const TAGCRD = SEPARATORE & "CRD"
Private Const TAGTOTCRD = SEPARATORE & "TOTCRD"
Private Const TAGCRT = SEPARATORE & "CRT"
Private Const TAGTOTCRT = SEPARATORE & "TOTCRT"
Private Const TAGCTS = SEPARATORE & "CTS"
Private Const TAGTOTCTS = SEPARATORE & "TOTCTS"
Private Const TAGCTR = SEPARATORE & "CTR"
Private Const TAGTOTCTR = SEPARATORE & "TOTCTR"
Private Const TAGCSL = SEPARATORE & "CSL"
Private Const TAGTOTCSL = SEPARATORE & "TOTCSL"

Private noLost As Boolean

Public Sub Edita(plvw As ListView, pfrm As Form, modal As Boolean)
    Dim i As Integer
    
    Set frm = pfrm
    Set lvw = plvw
    
    lblCampi(COMUNE_ID) = lvw.SelectedItem.SubItems(1)
    lblCampi(ISTAT_ID) = lvw.SelectedItem.Text
    lblCampi(REGIONE_ID) = lvw.SelectedItem.SubItems(3)
    lblCampi(PROVINCIA_ID) = lvw.SelectedItem.SubItems(4)
    lblCampi(CAP_ID) = lvw.SelectedItem.SubItems(5)
    lblCampi(DATA_ID) = lvw.SelectedItem.SubItems(2)

    For i = 0 To txt.Count - 4
        txt(i) = lvw.SelectedItem.SubItems(i + 6)
        txt(i).TabIndex = i
        txt(i).Locked = IIf(lblCampi(i).Tag <> vbNullString, True, False)
        txt(i).BackColor = IIf(lblCampi(i).Tag <> vbNullString, BACKCOLOR_CAMPI_CALC, vbWhite)
        txt(i).TabStop = Not txt(i).Locked
        txt(i).MaxLength = 15
        txt(i).Tag = CDbl(txt(i))
    Next i
    
    txt(CRD_ID).Tag = vbNullString
    txt(CRT_ID).Tag = vbNullString
    txt(CTS_ID).Tag = vbNullString
    txt(CSL_ID).Tag = vbNullString
    txt(CTR_COSTI_ID).Tag = vbNullString
    
    'Carico i Costi Totali di CRT , CTS , CRD , CSL
    i = lvw.SelectedItem.ListSubItems.Count
    
    txt(CSL_ID).Tag = SetValueInTag(TAGCSL, 0, txt(CSL_ID).Tag)
    txt(CSL_ID).Tag = SetValueInTag(TAGTOTCSL, lvw.SelectedItem.SubItems(i - 4), txt(CSL_ID).Tag)
    
    txt(CRT_ID).Tag = SetValueInTag(TAGCRT, 0, txt(CRT_ID).Tag)
    txt(CRT_ID).Tag = SetValueInTag(TAGTOTCRT, lvw.SelectedItem.SubItems(i - 3), txt(CRT_ID).Tag)
    
    txt(CRD_ID).Tag = SetValueInTag(TAGCRD, 0, txt(CRD_ID).Tag)
    txt(CRD_ID).Tag = SetValueInTag(TAGTOTCRD, lvw.SelectedItem.SubItems(i - 2), txt(CRD_ID).Tag)
    
    txt(CTS_ID).Tag = SetValueInTag(TAGCTS, 0, txt(CTS_ID).Tag)
    txt(CTS_ID).Tag = SetValueInTag(TAGTOTCTS, lvw.SelectedItem.SubItems(i - 1), txt(CTS_ID).Tag)
    
    txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, 0, txt(CTR_COSTI_ID).Tag)
    txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGTOTCTR, lvw.SelectedItem.SubItems(i), txt(CTR_COSTI_ID).Tag)
    
    
    txt(PARTE_FISSA_ID) = 0
    txt(PARTE_VARIABILE_ID) = 0
    txt(TOT_SERVIZIO_ID).MaxLength = 15
    
    RicalcolaCampi
    
    flgSimulazione = False
    
    If modal Then
        Me.Show vbModal
    Else
        Me.Show vbModeless
        flgSimulazione = True
        Command1(0).Visible = False
    End If
    
End Sub



Private Sub Command1_Click(index As Integer)
    Dim i As Integer
    Dim rs As Recordset
    
    
    Select Case index
        Case CMD_CONFERMA

            
            'Varia Dati Tecnici se varia Gettito
            Set rs = DatiTecniciComune(lblCampi(ISTAT_ID), lblCampi(DATA_ID))
            
            If Not rs.EOF Then
                If CDbl(txt(GETTITO_ID)) <> rs("TotGettitoDaUtenzaDomestica") + rs("TotGettitoDaUtenzaAttivita") Then
                    gDBC.Execute ("Update DatiTecniciComune set TotGettitoDaUtenzaDomestica=0,TotGettitoDaUtenzaAttivita=0 where AnnoRiferimento='" & lblCampi(DATA_ID) & "' and CodComune='" & lblCampi(ISTAT_ID) & "'")
                End If
            End If
            
            'Conferma
            For i = 0 To txt.Count - 4
                If InStr(txt(i), ",") > 0 Then
                    lvw.SelectedItem.SubItems(i + 6) = CDbl(txt(i))
                Else
                    lvw.SelectedItem.SubItems(i + 6) = CDbl(txt(i))
                End If
            Next i
            
            'Salvo i valori di Tot CSL,CRT,CRD e CTS sul MDB
            i = lvw.SelectedItem.ListSubItems.Count
            lvw.SelectedItem.SubItems(i - 4) = GetValueInTag(TAGTOTCSL, txt(CSL_ID).Tag)
            lvw.SelectedItem.SubItems(i - 3) = GetValueInTag(TAGTOTCRT, txt(CRT_ID).Tag)
            lvw.SelectedItem.SubItems(i - 2) = GetValueInTag(TAGTOTCRD, txt(CRD_ID).Tag)
            lvw.SelectedItem.SubItems(i - 1) = GetValueInTag(TAGTOTCTS, txt(CTS_ID).Tag)
            lvw.SelectedItem.SubItems(i) = GetValueInTag(TAGTOTCTR, txt(CTR_COSTI_ID).Tag)
                
            If Not flgSimulazione Then Unload Me
            
            ScritturaSuDB lvw, "CostoComunale"
        Case CMD_ANNULLA
            'AnnullA
            Unload Me
        Case CMD_DETTAGLIO
            'Dettaglio
            CaricaDettaglio
            Exit Sub
            
    End Select
        
    If Not flgSimulazione Then frm.Visible = True
End Sub

Private Sub Command1_GotFocus(index As Integer)
    Select Case index
        Case CMD_CONFERMA
            Info "Conferma Modifica dei dati"
        Case CMD_ANNULLA
            Info "Annulla Modifica dei dati"
        Case CMD_DETTAGLIO
            Info "Visualizza dettaglio per le specifiche di calcolo (la stessa funzione può essere ottenuta facendo doppio click sul campo)"
    End Select

End Sub

Private Sub Form_Load()
    CentraME Me
    If Not frm Is Nothing Then frm.Visible = False
    
    txt(EURO_ID) = 0
    'label
    lblCampi(TOT_SERVIZIO_ID) = "Tariffa Riferimento"
    '-----------
    lblCampi(CG_ID) = "Operativi gestione (CG)"
    lblCampi(CG_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CSL_ID) = "Pulizia e lavaggio p.zze pubbliche (CSL)"
    lblCampi(CSL_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CRT_ID) = "Raccolta e trasporto RU (CRT)"
    lblCampi(CRT_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CTS_ID) = "Smaltimento e trattamento RU (CTS)"
    lblCampi(CTS_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(AC_ID) = "Altri costi (AC)"
    lblCampi(CGIND_ID) = "Gest.Ciclo serv.RU indifferenziati (CGIND)"
    lblCampi(CGIND_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CRD_ID) = "Raccolta differenziata materiale (CRD)"
    lblCampi(CRD_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CTR_COSTI_ID) = "Trattamento e riciclo (CTR_COSTI)"
    lblCampi(CTR_COSTI_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CTR_RICAVI_ID) = "Ricavi dal riciclo (CTR_RICAVI)"
    lblCampi(CTR_ID) = "Trattamento e riciclo al netto (CTR)"
    lblCampi(CTR_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CON_ID) = "Raccolta e imballaggi CONAI (CON)"
    lblCampi(Y_ID) = "Second.Terzia. a carico di prod.e utiliz.(Y)"
    lblCampi(CGD_ID) = "Gest.Ciclo raccolta differenziata (CGD)"
    lblCampi(CGD_ID).Tag = FORECOLOR_CAMPI_CALC
    '---
    lblCampi(CC_ID) = "Costi Comuni (CC)"
    lblCampi(CC_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CARC_ID) = "Risc. Accertamento contenz. (CARC)"
    lblCampi(CGG_ID) = "Generali di gestione (CGG)"
    lblCampi(CGG_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(CCD_ID) = "Costi comuni diversi (CCD)"
    lblCampi(CK_ID) = "Costi d'uso del capitale (CK)"
    lblCampi(CK_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(AMM_ID) = "Ammortamenti (AMM)"
    lblCampi(AMM_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(ACC_ID) = "Accantonamenti (ACC)"
    lblCampi(ACC_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(R_ID) = "Remunerazione del capitale (R)"
    lblCampi(R_ID).Tag = FORECOLOR_CAMPI_CALC
    lblCampi(RN_ID) = "Tasso di remunerazione capitale (RN) %"
    lblCampi(KN_ID) = "Immobiliz.netta all'anno prec. (KN-1)"
    lblCampi(IN_ID) = "Investimenti programmati (In)"
    lblCampi(FN_ID) = "Fattore correttivo per investimenti (Fn)"
    
    'Costi Parte fissa
    lblCampi(CSL_ID).ForeColor = FORECOLOR_PARTE_FISSA
    lblCampi(CARC_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    lblCampi(CGG_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    lblCampi(CCD_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    lblCampi(AC_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    lblCampi(CK_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    lblCampi(PARTE_FISSA_ID).ForeColor = lblCampi(CSL_ID).ForeColor
    txt(PARTE_FISSA_ID).BackColor = BACKCOLOR_CAMPI_CALC
    'Costi Parte Variabile
    lblCampi(CRT_ID).ForeColor = FORECOLOR_PARTE_VARIABILE
    lblCampi(CTS_ID).ForeColor = lblCampi(CRT_ID).ForeColor
    lblCampi(CRD_ID).ForeColor = lblCampi(CRT_ID).ForeColor
    lblCampi(CTR_ID).ForeColor = lblCampi(CRT_ID).ForeColor
    lblCampi(PARTE_VARIABILE_ID).ForeColor = lblCampi(CRT_ID).ForeColor
    txt(PARTE_VARIABILE_ID).BackColor = BACKCOLOR_CAMPI_CALC

    
    
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim nextX As Long
 
    nextX = Me.ScaleWidth - Command1(0).Width - 100
    For i = Command1.LBound To Command1.UBound
        If Command1(i).Visible Then
            Command1(i).Move nextX, Me.ScaleHeight - Command1(i).Height - 100
            If i < Command1.UBound Then nextX = nextX - Command1(i + 1).Width - 100
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Info vbNullString
End Sub

Private Sub txt_DblClick(index As Integer)
    'Vado nella form di dettaglio corrispondente
    noLost = True
    CaricaDettaglio
    noLost = False
End Sub

Private Sub txt_GotFocus(index As Integer)

    If noLost Then Exit Sub
    'Attiva il bottone dettaglio per i costi Accontonamento e Ammortamento
    mDettaglio = index
    If index <> ACC_ID And index <> AMM_ID And index <> CGG_ID And _
       index <> CRD_ID And index <> CTS_ID And index <> CRT_ID And _
       index <> CTR_COSTI_ID And index <> CSL_ID Then
        Command1(CMD_DETTAGLIO).Enabled = False
    Else
        Command1(CMD_DETTAGLIO).Enabled = True
    End If
    
    
    Select Case index
        Case TOT_SERVIZIO_ID
            Info ("Tariffa a regime in €/Anno articolo 2 D.P.R 158/99 equivalenza punto 1 dell'allegato 1")
        Case GETTITO_ID
            Info "Totale in €/Anno delle entrate tariffarie per l'anno di riferimento"
        Case CG_ID, CSL_ID, CRT_ID, CTS_ID, AC_ID, CGIND_ID, CRD_ID, CTR_COSTI_ID, _
            CTR_RICAVI_ID, CTR_ID, CON_ID, Y_ID, CGD_ID, CC_ID, CARC_ID, CGG_ID, _
            CCD_ID, CK_ID, AMM_ID, ACC_ID, R_ID, KN_ID, IN_ID, FN_ID
            Info lblCampi(index) & " in €/Anno"
        Case RN_ID
            Info ("Valore percentuale del tasso di remunerazione voluto")
        Case PARTE_VARIABILE_ID
            Info ("Quota variabile in €/Anno della tariffa di riferimento comma 4 art 49 D.Lgs 22/97")
        Case PARTE_FISSA_ID
            Info ("Quota fissa in €/Anno della tariffa di riferimento comma 4 art 49 D.Lgs 22/97")
        Case EURO_ID
            Info ("Tariffa a regime espressa in EURO")
        
    End Select
    
    If txt(index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub
    
    
    'Imposto il numero MAX di caratteri in EDIT MODE
    txt(index) = CDbl(txt(index))
    txt(index).MaxLength = 12
    
    
    txt(index).SelStart = 0
    txt(index).SelLength = Len(txt(index))
    OldTxtValue = txt(index)
    
    
    
End Sub

Private Sub txt_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Dim CancelERR As Boolean
    
    If KeyCode = vbKeyReturn Then
        If index + 1 > txt.UBound Then
            index = 0
        Else
            txt_Validate index, CancelERR
            If Not CancelERR Then txt(index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_LostFocus(index As Integer)
    
    If noLost Then Exit Sub
    txt(index).MaxLength = 15
    
    'Ricalcolo
    If txt(CTR_COSTI_ID) = vbNullString Then
        txt(CTR_COSTI_ID) = 0
    End If
    
    RicalcolaCampi
End Sub

Private Sub txt_Validate(index As Integer, Cancel As Boolean)
    If noLost Then Exit Sub
    
    'Il valore deve essere numerico
    If Not IsNumeric(txt(index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txt(index) = OldTxtValue: Exit Sub
    
        
    'Controlli sui campi
    Select Case index
        Case RN_ID
            'Il tasso di renumerazione deve essere un percentuale
            If (CDbl(txt(RN_ID)) < 0 Or CDbl(txt(RN_ID)) > 100) Then
                MsgBox "Il valore deve essere un percentuale", vbExclamation + vbOKOnly
                Cancel = True
            End If
        Case CTR_RICAVI_ID, CTR_COSTI_ID
            'CTR_COSTI > CTR_RICAVI
            If (CDbl(txt(CTR_COSTI_ID)) - CDbl(txt(CTR_RICAVI_ID)) < 0) Then
                MsgBox "I ricavi dalla vendita del riciclato non" + vbCrLf + "possono essere maggiori dei costi sostenuti.", vbExclamation + vbOKOnly
                Cancel = True
            End If
        Case GETTITO_ID
            'il gettito non deve essere maggiore del costo totale del servizio
            If CDbl(txt(GETTITO_ID)) > CDbl(txt(TOT_SERVIZIO_ID)) Then
                MsgBox "Il gettito non può essere maggiore del costo totale.", vbExclamation + vbOKOnly
                Cancel = True
            End If
    End Select
End Sub

Private Sub RicalcolaCampi()
    Dim i As Integer
    Dim Valore As Double
    
    
    'Mette in valore interi o decimali
    For i = 0 To txt.Count - 1
        If i = RN_ID Then
            txt(i) = Format$(CDbl(txt(i)), "###,###,###,##0.00")
        ElseIf InStr(1, txt(i), ".") = 0 Then
            txt(i) = Format$(CDbl(txt(i)), FORMATVALUTA)
        End If
    Next
    
    
    'CTS
    If CDbl(GetValueInTag(TAGCTS, txt(CTS_ID).Tag)) <> CDbl(txt(CTS_ID)) Then
        Valore = CDbl(GetValueInTag(TAGTOTCTS, txt(CTS_ID).Tag))
        txt(CTS_ID) = Format$(CostiSmaltimentoETrattamento(Valore), FORMATVALUTA)
        txt(CTS_ID).Tag = SetValueInTag(TAGTOTCTS, Valore, txt(CTS_ID).Tag)
        txt(CTS_ID).Tag = SetValueInTag(TAGCTS, CDbl(txt(CTS_ID)), txt(CTS_ID).Tag)
    End If
    
    'CSL
    If CDbl(GetValueInTag(TAGCSL, txt(CSL_ID).Tag)) <> CDbl(txt(CSL_ID)) Then
        Valore = CDbl(GetValueInTag(TAGTOTCSL, txt(CSL_ID).Tag))
        txt(CSL_ID) = Format$(CostiSpazzamento(Valore), FORMATVALUTA)
        txt(CSL_ID).Tag = SetValueInTag(TAGTOTCSL, Valore, txt(CSL_ID).Tag)
        txt(CSL_ID).Tag = SetValueInTag(TAGCSL, CDbl(txt(CSL_ID)), txt(CSL_ID).Tag)
    End If
    
    'CRT
    If CDbl(GetValueInTag(TAGCRT, txt(CRT_ID).Tag)) <> CDbl(txt(CRT_ID)) Then
        Valore = CDbl(GetValueInTag(TAGTOTCRT, txt(CRT_ID).Tag))
        txt(CRT_ID) = CostiRaccoltaETrasporto(Valore)
        txt(CRT_ID).Tag = SetValueInTag(TAGTOTCRT, Valore, txt(CRT_ID).Tag)
        txt(CRT_ID).Tag = SetValueInTag(TAGCRT, CDbl(txt(CRT_ID)), txt(CRT_ID).Tag)
    End If
    
    'CRD
    If CDbl(GetValueInTag(TAGCRD, txt(CRD_ID).Tag)) <> CDbl(txt(CRD_ID)) Then
        Valore = CDbl(GetValueInTag(TAGTOTCRD, txt(CRD_ID).Tag))
        txt(CRD_ID) = CostiRaccoltaDifferenziataMateriale(Valore)
        txt(CRD_ID).Tag = SetValueInTag(TAGTOTCRD, Valore, txt(CRD_ID).Tag)
        txt(CRD_ID).Tag = SetValueInTag(TAGCRD, CDbl(txt(CRD_ID)), txt(CRD_ID).Tag)
    End If
    

    'CTR
    If CDbl(GetValueInTag(TAGCTR, txt(CTR_COSTI_ID).Tag)) <> CDbl(txt(CTR_COSTI_ID)) Then
        Valore = CDbl(GetValueInTag(TAGTOTCTR, txt(CTR_COSTI_ID).Tag))
        txt(CTR_COSTI_ID) = Format$(CostiTrattamentoRiciclo(Valore), FORMATVALUTA)
        txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGTOTCTR, Valore, txt(CTR_COSTI_ID).Tag)
        txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, CDbl(txt(CTR_COSTI_ID)), txt(CTR_COSTI_ID).Tag)
    End If

    'CTR
    txt(CTR_ID) = CDbl(txt(CTR_COSTI_ID)) - CDbl(txt(CTR_RICAVI_ID))
    
    
    'CG
    txt(CG_ID) = 0
    For i = CSL_ID To AC_ID
        txt(CG_ID) = CDbl(txt(CG_ID)) + CDbl(txt(i))
    Next
    txt(CG_ID) = CDbl(txt(CG_ID)) + CDbl(txt(CRD_ID)) + CDbl(txt(CTR_ID)) - CDbl(txt(CON_ID)) - CDbl(txt(Y_ID))
    
    'CGIND
    txt(CGIND_ID) = 0
    For i = CSL_ID To AC_ID
        txt(CGIND_ID) = CDbl(txt(CGIND_ID)) + CDbl(txt(i))
    Next
    
    'CGD
    txt(CGD_ID) = 0
    txt(CGD_ID) = CDbl(txt(CTR_ID)) + CDbl(txt(CRD_ID))
    
    'CC
    txt(CC_ID) = 0
    For i = CARC_ID To CCD_ID
        txt(CC_ID) = CDbl(txt(CC_ID)) + CDbl(txt(i))
    Next
    
    

    'R
    txt(R_ID) = 0
    For i = KN_ID To FN_ID
        txt(R_ID) = CDbl(txt(R_ID)) + CDbl(txt(i))
    Next
    txt(R_ID) = Format$(CDbl(txt(R_ID) / 100) * CDbl(txt(RN_ID)), FORMATVALUTA)

    'CK
    txt(CK_ID) = CDbl(txt(R_ID))
    For i = AMM_ID To ACC_ID
        txt(CK_ID) = CDbl(txt(CK_ID)) + CDbl(txt(i))
    Next

    'Parte Fissa
    txt(PARTE_FISSA_ID) = Format$(CDbl(txt(CSL_ID)) + CDbl(txt(CARC_ID)) _
        + CDbl(txt(CGG_ID)) + CDbl(txt(CCD_ID)) + CDbl(txt(AC_ID)) + CDbl(txt(CK_ID)), FORMATVALUTA)
    
    'Parte Variabile
    txt(PARTE_VARIABILE_ID) = Format$(CDbl(txt(CRT_ID)) + CDbl(txt(CTS_ID)) _
        + CDbl(txt(CRD_ID)) + CDbl(txt(CTR_ID)) - CDbl(txt(CON_ID)) - CDbl(txt(Y_ID)), FORMATVALUTA)
    
    'Gettito
    If CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID)) <> 0 Then
        txt(GETTITO_ID) = CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID))
        Label1(0).Caption = "PF:" & Round((CDbl(txt(PARTE_FISSA_ID)) / (CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID)))) * 100, 0) & "%"
        Label1(1).Caption = "PV:" & Round((CDbl(txt(PARTE_VARIABILE_ID)) / (CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID)))) * 100, 0) & "%"
    End If
    
    
    'Costi Totali
    If CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID)) > 0 Then
        txt(TOT_SERVIZIO_ID).BackColor = BACKCOLOR_CAMPI_CALC
        txt(TOT_SERVIZIO_ID).Locked = True
        txt(TOT_SERVIZIO_ID) = Format$(CDbl(txt(PARTE_FISSA_ID)) + CDbl(txt(PARTE_VARIABILE_ID)), FORMATVALUTA)
    Else
        txt(TOT_SERVIZIO_ID).BackColor = vbWhite
        txt(TOT_SERVIZIO_ID).Locked = False
    End If
    
    'Costi Totali in EURO
    'txt(EURO_ID) = FORMAT$(InEuro(cdbl(txt(TOT_SERVIZIO_ID))), "###,###,##0.00")
     
    'Formatta i campi
    For i = 0 To txt.Count - 1
        If i = RN_ID Then
            txt(i) = Format$(CDbl(txt(i)), "###,###,###,##0.00")
        ElseIf InStr(1, txt(i), ".") = 0 Then
                txt(i) = Format$(CDbl(txt(i)), FORMATVALUTA)
        End If
    Next
     
    


End Sub

Private Sub CaricaDettaglio()
    'Carica la form di dettaglio del costo
    Select Case mDettaglio
        Case ACC_ID
            txt(ACC_ID) = frmAccantonamenti.Carica(CDbl(txt(ACC_ID)), lblCampi(ISTAT_ID), lblCampi(DATA_ID))
    
            txt(CRT_ID).Tag = SetValueInTag(TAGCRT, 0, txt(CRT_ID).Tag)
            txt(CRD_ID).Tag = SetValueInTag(TAGCRD, 0, txt(CRD_ID).Tag)
            txt(CTS_ID).Tag = SetValueInTag(TAGCTS, 0, txt(CTS_ID).Tag)
            txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, 0, txt(CTR_COSTI_ID).Tag)
            
           txt(ACC_ID).SetFocus
        Case AMM_ID
            txt(AMM_ID) = frmAmmortamenti.Carica(CDbl(txt(AMM_ID)), lblCampi(ISTAT_ID), lblCampi(DATA_ID))
    
            txt(CRT_ID).Tag = SetValueInTag(TAGCRT, 0, txt(CRT_ID).Tag)
            txt(CRD_ID).Tag = SetValueInTag(TAGCRD, 0, txt(CRD_ID).Tag)
            txt(CTS_ID).Tag = SetValueInTag(TAGCTS, 0, txt(CTS_ID).Tag)
            txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, 0, txt(CTR_COSTI_ID).Tag)
            
            txt(AMM_ID).SetFocus
        Case CGG_ID
            txt(CGG_ID) = frmCostiDelPersonale.Carica(lblCampi(ISTAT_ID), lblCampi(DATA_ID), CDbl(txt(CGG_ID)))
            
            txt(CRT_ID).Tag = SetValueInTag(TAGCRT, 0, txt(CRT_ID).Tag)
            txt(CRD_ID).Tag = SetValueInTag(TAGCRD, 0, txt(CRD_ID).Tag)
            txt(CTS_ID).Tag = SetValueInTag(TAGCTS, 0, txt(CTS_ID).Tag)
            txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, 0, txt(CTR_COSTI_ID).Tag)
            
            txt(CGG_ID).SetFocus
        Case CRD_ID
            CostoSave = CDbl(GetValueInTag(TAGTOTCRD, txt(CRD_ID).Tag))
            txt(CRD_ID) = frmDettaglioCostiRifiuto.CalcoloCostoRifiuto(Me, CostoSave, CostoRaccoltaDifferenziataCRD)
            txt(CRD_ID).Tag = SetValueInTag(TAGTOTCRD, CostoSave, txt(CRD_ID).Tag)
            txt(CRD_ID).Tag = SetValueInTag(TAGCRD, CDbl(txt(CRD_ID)), txt(CRD_ID).Tag)


            txt(CRD_ID).SetFocus
        Case CRT_ID
            CostoSave = CDbl(GetValueInTag(TAGTOTCRT, txt(CRT_ID).Tag))
            txt(CRT_ID) = frmDettaglioCostiRifiuto.CalcoloCostoRifiuto(Me, CostoSave, CostoRaccoltaeTraportoCRT)
            txt(CRT_ID).Tag = SetValueInTag(TAGTOTCRT, CostoSave, txt(CRT_ID).Tag)
            txt(CRT_ID).Tag = SetValueInTag(TAGCRT, CDbl(txt(CRT_ID)), txt(CRT_ID).Tag)

            txt(CRT_ID).SetFocus
        Case CSL_ID
            CostoSave = CDbl(GetValueInTag(TAGTOTCSL, txt(CSL_ID).Tag))
            txt(CSL_ID) = frmDettaglioCostiRifiuto.CalcoloCostoRifiuto(Me, CostoSave, CostoSpazzamentoCSL)
            txt(CSL_ID).Tag = SetValueInTag(TAGTOTCSL, CostoSave, txt(CSL_ID).Tag)
            txt(CSL_ID).Tag = SetValueInTag(TAGCSL, CDbl(txt(CSL_ID)), txt(CSL_ID).Tag)

            txt(CSL_ID).SetFocus
        
        Case CTS_ID
            CostoSave = CDbl(GetValueInTag(TAGTOTCTS, txt(CTS_ID).Tag))
            txt(CTS_ID) = frmDettaglioCostiRifiuto.CalcoloCostoRifiuto(Me, CostoSave, CostoTrattamentoeSmaltimentoCTS)
            txt(CTS_ID).Tag = SetValueInTag(TAGTOTCTS, CostoSave, txt(CTS_ID).Tag)
            txt(CTS_ID).Tag = SetValueInTag(TAGCTS, CDbl(txt(CTS_ID)), txt(CTS_ID).Tag)

            txt(CTS_ID).SetFocus
        Case CTR_COSTI_ID
            CostoSave = CDbl(GetValueInTag(TAGTOTCTR, txt(CTR_COSTI_ID).Tag))
            txt(CTR_COSTI_ID) = frmDettaglioCostiRifiuto.CalcoloCostoRifiuto(Me, CostoSave, CostoTrattamentoeRicicloCTR)
            txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGTOTCTR, CostoSave, txt(CTR_COSTI_ID).Tag)
            txt(CTR_COSTI_ID).Tag = SetValueInTag(TAGCTR, CDbl(txt(CTR_COSTI_ID)), txt(CTR_COSTI_ID).Tag)

            txt(CTR_COSTI_ID).SetFocus
    
    
    End Select
    
    
End Sub


Public Function CostiRaccoltaETrasporto(ByVal CostoTotale As Double, Optional CostoPersonaleCRT As Double = 0, Optional CostoPersonaleDirig As Double = 0, Optional AmmortamentiCRT As Double = 0, Optional AmmortamentiStrutture As Double = 0) As Double
    'I costi di raccolta e trasporto sono dati dal costo totale di raccolta e trasporto
    '-  gli ammortamenti relativi
    '-% costo del personale relativo
    '+3%  degli ammortamenti di strutture comuni
    '+3%  del costo del personale dirigente
    Dim PercCostoPersonale As Double
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim i As Integer
    Dim NumPdiLivello As Long

    
    CostiRaccoltaETrasporto = CostoTotale
    
    PercCostoPersonale = GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "CPPERSON")
    If PercCostoPersonale > -1 Then
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERRACTR") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Set rs1 = gDBC.Execute("Select count(*) from LivelliPersonaleImpiegato ")
        
        If rs1.EOF Then MsgBox "Attenzione mancano i Livelli del personale nella tabella LivelliPersonaleImpiegato", vbCritical + vbOKOnly
        
        
        CostoPersonaleCRT = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleCRT = CostoPersonaleCRT + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERDIRIG") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        
        
        'Costo del personale
        CostoPersonaleDirig = 0
        If Not rs.EOF Then
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleDirig = CostoPersonaleDirig + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        Set rs1 = Nothing
        Set rs = Nothing
        
        AmmortamentiCRT = 0
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento=" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMRACTR") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiCRT = rs(0)
        End If
        
        AmmortamentiStrutture = 0
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and  CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSTRUT") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiStrutture = rs(0)
        End If
                
        
        CostoPersonaleCRT = Int((PercCostoPersonale / 100) * CostoPersonaleCRT)
        AmmortamentiStrutture = Int(0.03 * AmmortamentiStrutture)
        CostoPersonaleDirig = Int(0.03 * CostoPersonaleDirig)
        
        CostiRaccoltaETrasporto = CostiRaccoltaETrasporto - (CostoPersonaleCRT + AmmortamentiCRT) + AmmortamentiStrutture + CostoPersonaleDirig
        
        
    End If
End Function


Public Function CostiRaccoltaDifferenziataMateriale(ByVal CostoTotale As Double, Optional CostoPersonaleCRD As Double = 0, Optional CostoPersonaleDirig As Double = 0, Optional AmmortamentiCRD As Double = 0, Optional AmmortamentiStrutture As Double = 0) As Double
    'I costi di Raccolta Differenziata Materiale sono dati dal costo totale
    'di Raccolta Differenziata Materiale
    
    '-  gli ammortamenti relativi
    '-% costo del personale relativo
    '+15%  degli ammortamenti di strutture comuni
    '+3%  del costo del personale dirigente
    Dim PercCostoPersonale As Double
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim i As Integer
    Dim NumPdiLivello As Long
    
    CostiRaccoltaDifferenziataMateriale = CostoTotale
    
    PercCostoPersonale = GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "CPPERSON")
    If PercCostoPersonale > -1 Then
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERRACDI") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Set rs1 = gDBC.Execute("Select count(*) from LivelliPersonaleImpiegato ")
        
        
        If rs1.EOF Then MsgBox "Attenzione mancano i Livelli del personale nella tabella LivelliPersonaleImpiegato", vbCritical + vbOKOnly
        CostoPersonaleCRD = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleCRD = CostoPersonaleCRD + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERDIRIG") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        CostoPersonaleDirig = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleDirig = CostoPersonaleDirig + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        Set rs1 = Nothing
        Set rs = Nothing
        
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento=" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMRACDI") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiCRD = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiCRD = rs(0)
        End If
        
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and  CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSTRUT") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiStrutture = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiStrutture = rs(0)
        End If
        
        
                
        CostoPersonaleCRD = Int((PercCostoPersonale / 100) * CostoPersonaleCRD)
        AmmortamentiStrutture = Int((0.15 * AmmortamentiStrutture))
        CostoPersonaleDirig = Int(0.03 * CostoPersonaleDirig)
        
        CostiRaccoltaDifferenziataMateriale = CostoTotale - (CostoPersonaleCRD + AmmortamentiCRD) + AmmortamentiStrutture + CostoPersonaleDirig
        
    End If
End Function


Public Function CostiSpazzamento(ByVal CostoTotale As Double, Optional CostoPersonaleCSL As Double = 0, Optional CostoPersonaleDirig As Double = 0, Optional AmmortamentiCSL As Double = 0, Optional AmmortamentiStrutture As Double = 0) As Double
    'I costi di Spazzamento sono dati dal costo totale
    'di Raccolta Differenziata Materiale
    
    '-  gli ammortamenti relativi
    '-% costo del personale relativo
    '+15%  degli ammortamenti di strutture comuni
    '+3%  del costo del personale dirigente
    Dim PercCostoPersonale As Double
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim i As Integer
    Dim NumPdiLivello As Long
    
    CostiSpazzamento = CostoTotale
    
    PercCostoPersonale = GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "CPPERSON")
    If PercCostoPersonale > -1 Then
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERSPAZ") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Set rs1 = gDBC.Execute("Select count(*) from LivelliPersonaleImpiegato ")
        
        
        If rs1.EOF Then MsgBox "Attenzione mancano i Livelli del personale nella tabella LivelliPersonaleImpiegato", vbCritical + vbOKOnly
        CostoPersonaleCSL = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleCSL = CostoPersonaleCSL + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERDIRIG") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        CostoPersonaleDirig = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleDirig = CostoPersonaleDirig + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        Set rs1 = Nothing
        Set rs = Nothing
        
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento=" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSPAZ") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiCSL = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiCSL = rs(0)
        End If
        
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and  CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSTRUT") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiStrutture = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiStrutture = rs(0)
        End If
        
        
                
        CostoPersonaleCSL = Int((PercCostoPersonale / 100) * CostoPersonaleCSL)
        AmmortamentiStrutture = Int((0.15 * AmmortamentiStrutture))
        CostoPersonaleDirig = Int(0.03 * CostoPersonaleDirig)
        
        CostiSpazzamento = CostoTotale - (CostoPersonaleCSL + AmmortamentiCSL) + AmmortamentiStrutture + CostoPersonaleDirig
        
    End If
End Function


Public Function CostiSmaltimentoETrattamento(ByVal CostoTotale As Double, Optional CostoPersonaleCTS As Double = 0, Optional CostoPersonaleDirig As Double = 0, Optional AmmortamentiCTS As Double = 0, Optional AmmortamentiStrutture As Double = 0, Optional Accantonamenti As Double = 0) As Double
    'I costi di Smaltimento e Trattamento sono dati dal costo totale di Smaltimento e Trattamento
    '-gli ammortamenti relativi
    '-% costo del personale relativo
    '+3%  degli ammortamenti di strutture comuni
    '+3%  del costo del personale dirigente
    '-Accantonamenti
    Dim PercCostoPersonale As Double
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim i As Integer
    Dim NumPdiLivello As Long
    
    CostiSmaltimentoETrattamento = CostoTotale
    
    PercCostoPersonale = GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "CPPERSON")
    If PercCostoPersonale > -1 Then
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERSMATR") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Set rs1 = gDBC.Execute("Select count(*) from LivelliPersonaleImpiegato ")
        
        If rs1.EOF Then MsgBox "Attenzione mancano i Livelli del personale nella tabella LivelliPersonaleImpiegato", vbCritical + vbOKOnly
        
        CostoPersonaleCTS = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleCTS = CostoPersonaleCTS + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        'Personale Dirigente
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERDIRIG") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        CostoPersonaleDirig = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleDirig = CostoPersonaleDirig + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                    
                End If
            Next
        End If
        
        Set rs1 = Nothing
        Set rs = Nothing
        
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento=" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSMATR") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiCTS = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiCTS = rs(0)
        End If
        
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and  CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSTRUT") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiStrutture = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiStrutture = rs(0)
        End If
        
                
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(A.Totale) from Accantonamenti as A inner join TipiAccantonamento as B" & _
            " on A.CodAccantonamento= B.CodAccantonamento where B.TipoCosto='CTS' and A.CodComune='" & lblCampi(ISTAT_ID) & "' and A.AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Accantonamenti = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then Accantonamenti = rs(0)
        End If
                
        CostoPersonaleCTS = Int((PercCostoPersonale / 100) * CostoPersonaleCTS)
        AmmortamentiStrutture = Int(0.15 * AmmortamentiStrutture)
        CostoPersonaleDirig = Int(0.03 * CostoPersonaleDirig)
        CostiSmaltimentoETrattamento = CostiSmaltimentoETrattamento - (CostoPersonaleCTS + AmmortamentiCTS) + AmmortamentiStrutture + CostoPersonaleDirig - Accantonamenti
    End If
End Function

Public Function CostiTrattamentoRiciclo(ByVal CostoTotale As Double, Optional CostoPersonaleCTR As Double = 0, Optional CostoPersonaleDirig As Double = 0, Optional AmmortamentiCTR As Double = 0, Optional AmmortamentiStrutture As Double = 0, Optional Accantonamenti As Double = 0) As Double
    'I costi di Smaltimento e Trattamento sono dati dal costo totale di Smaltimento e Trattamento
    '-gli ammortamenti relativi
    '-% costo del personale relativo
    '+3%  degli ammortamenti di strutture comuni
    '+3%  del costo del personale dirigente
    '-Accantonamenti
    Dim PercCostoPersonale As Double
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim i As Integer
    Dim NumPdiLivello As Long
    
    CostiTrattamentoRiciclo = CostoTotale
    
    PercCostoPersonale = GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "CPPERSON")
    If PercCostoPersonale > -1 Then
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERTRARIC") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Set rs1 = gDBC.Execute("Select count(*) from LivelliPersonaleImpiegato ")
        
        If rs1.EOF Then MsgBox "Attenzione mancano i Livelli del personale nella tabella LivelliPersonaleImpiegato", vbCritical + vbOKOnly
        
        CostoPersonaleCTR = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleCTR = CostoPersonaleCTR + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        
        'Personale Dirigente
        Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune where CodComune='" & lblCampi(ISTAT_ID) & "' and CodTipoPersonale =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "PERDIRIG") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        CostoPersonaleDirig = 0
        If Not rs.EOF Then
            'Costo del personale
            For i = 1 To rs1.fields(0)
                NumPdiLivello = Choose(i, rs("Liv2"), rs("Liv3"), rs("Liv4"), rs("Liv5"), rs("Liv6"), rs("Liv7"), rs("Liv7q"), rs("Liv8"), rs("Liv8q"), rs("Amministrativo"), rs("Manager"))
                If NumPdiLivello > 0 Then
                    CostoPersonaleDirig = CostoPersonaleDirig + (NumPdiLivello * GetRetribuzione(lblCampi(ISTAT_ID), lblCampi(DATA_ID), rs("CodTipoPersonale"), COD_LIVELLO & CStr(i)))
                End If
            Next
        End If
        Set rs1 = Nothing
        Set rs = Nothing
        
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento=" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMTRARIC") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiCTR = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiCTR = rs(0)
        End If
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(QuotaAnnua) from Ammortamenti,TipiAmmortamento where Ammortamenti.CodTipoAmmortamento= TipiAmmortamento.CodTipoAmmortamento and  CodComune='" & lblCampi(ISTAT_ID) & "' and FamigliaAmmortamento =" & GetParamValue(lblCampi(DATA_ID), lblCampi(ISTAT_ID), "AMMSTRUT") & " and AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        AmmortamentiStrutture = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then AmmortamentiStrutture = rs(0)
        End If
        
        Set rs = Nothing
        Set rs = gDBC.Execute("Select Sum(A.Totale) from Accantonamenti as A inner join TipiAccantonamento as B" & _
            " on A.CodAccantonamento= B.CodAccantonamento where B.TipoCosto='CTR' and A.CodComune='" & lblCampi(ISTAT_ID) & "' and A.AnnoRiferimento='" & lblCampi(DATA_ID) & "'")
        Accantonamenti = 0
        If Not rs.EOF Then
            If Not (rs(0) = 0 Or IsNull(rs(0))) Then Accantonamenti = rs(0)
        End If
        
        CostoPersonaleCTR = Int((PercCostoPersonale / 100) * CostoPersonaleCTR)
        AmmortamentiStrutture = Int(0.15 * AmmortamentiStrutture)
        CostoPersonaleDirig = Int(0.03 * CostoPersonaleDirig)
        CostiTrattamentoRiciclo = CostiTrattamentoRiciclo - (CostoPersonaleCTR + AmmortamentiCTR) + AmmortamentiStrutture + CostoPersonaleDirig - Accantonamenti
    End If
End Function

