VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmPubblica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pubblica i dati sul WEB"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Filtro"
      Height          =   390
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkLocale 
      Caption         =   "Local Host"
      Height          =   330
      Left            =   5460
      TabIndex        =   15
      Top             =   4320
      Width           =   1110
   End
   Begin VB.TextBox lbl 
      Height          =   1275
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2940
      Width           =   6540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tariffa"
      Height          =   2775
      Left            =   3375
      TabIndex        =   6
      Top             =   75
      Width           =   3150
      Begin VB.CheckBox Check1 
         Caption         =   "Statistiche"
         Height          =   360
         Index           =   8
         Left            =   195
         TabIndex        =   14
         Top             =   1110
         Width           =   2550
      End
      Begin VB.CheckBox Check1 
         Caption         =   "StatisticheComunali"
         Height          =   300
         Index           =   7
         Left            =   195
         TabIndex        =   13
         Top             =   2400
         Width           =   2550
      End
      Begin VB.CheckBox Check1 
         Caption         =   "UtenzeDomesticheComunali"
         Height          =   540
         Index           =   6
         Left            =   195
         TabIndex        =   10
         Top             =   1875
         Width           =   2550
      End
      Begin VB.CheckBox Check1 
         Caption         =   "NumeroComponentiFamiglia"
         Height          =   540
         Index           =   5
         Left            =   195
         TabIndex        =   9
         Top             =   630
         Width           =   2550
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CategorieAttivitaComunali"
         Height          =   540
         Index           =   4
         Left            =   195
         TabIndex        =   8
         Top             =   1455
         Width           =   2550
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CategorieAttivita"
         Height          =   540
         Index           =   3
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   2550
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Anagrafiche"
      Height          =   2760
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   3150
      Begin VB.CheckBox Check1 
         Caption         =   "Regioni"
         Height          =   540
         Index           =   2
         Left            =   195
         TabIndex        =   5
         Top             =   1020
         Width           =   990
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Comuni"
         Height          =   540
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   240
         Width           =   990
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Provincie"
         Height          =   540
         Index           =   1
         Left            =   195
         TabIndex        =   3
         Top             =   630
         Width           =   990
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Pubblica"
      Height          =   390
      Index           =   1
      Left            =   4335
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Chiudi"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   5460
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin MSComCtl2.Animation anim 
      Height          =   570
      Left            =   60
      TabIndex        =   12
      Top             =   4500
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1005
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   281
      FullHeight      =   38
   End
End
Attribute VB_Name = "frmPubblica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mWhereFiltroSuComune As String
Private mDataDa As Date
Private mDataA As Date
Private mCodComune As String

Private Sub Command1_Click(Index As Integer)
    Dim i As Long
    Dim l As Long
    Dim sErr As String
    Dim sFiltro As String
    
    Select Case Index
        Case 2
            frmFiltroComune.SetFiltro mDataDa, mDataA, mCodComune
        Case 1
            anim.Visible = True
            anim.Open App.Path & "/filecopy.avi"
            Screen.MousePointer = vbHourglass
            DoEvents
            For i = 0 To Check1.Count - 1
                If Check1(i).Value = 1 Then
                    Info (Check1(i).Caption & " in corso")
                    If i = 4 Or i = 6 Or i = 7 Then
                        sFiltro = gWhereFiltroSuComune
                    Else
                        sFiltro = vbNullString
                    End If
                    l = PublicMDBonTheWeb("Select * from " & Check1(i).Caption & " " & sFiltro, Check1(i).Caption, chkLocale.Value, sErr)
                    Info (Check1(i).Caption & " - " & sErr)
                End If
            Next
            anim.Visible = False
            Screen.MousePointer = vbDefault
            anim.Close
        Case 0
            gWhereFiltroSuComune = mWhereFiltroSuComune
            Unload Me
    End Select

End Sub

Private Sub Form_Load()
    anim.Visible = False
    mWhereFiltroSuComune = gWhereFiltroSuComune
    mDataDa = DateAdd("yyyy", -1, Now)
    mDataA = Now
    mCodComune = gComuneAttivo
End Sub


Sub Info(s As String)
    Static line As Long
    Dim l As Long
    
    lbl = lbl + s + vbCrLf
    lbl.Refresh
    
    l = SendMessage(lbl.hwnd, EM_LINESCROLL, 0, line - 1)
    line = line + 1
End Sub



