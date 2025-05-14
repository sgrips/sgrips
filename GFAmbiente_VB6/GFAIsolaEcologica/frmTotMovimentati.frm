VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTotMovimentati 
   Caption         =   "Giacenze di Magazzino"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   12780
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   0
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   4575
      TabIndex        =   3
      Top             =   4440
      Width           =   4575
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   56492033
         CurrentDate     =   39945
      End
      Begin MSComctlLib.ListView lvwMovimenti 
         Height          =   2415
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   56492033
         CurrentDate     =   39945
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Dal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Stampa"
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
      Left            =   8715
      TabIndex        =   2
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ricarica"
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
      Left            =   10035
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
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
      Left            =   11400
      TabIndex        =   0
      Top             =   8640
      Width           =   1095
   End
End
Attribute VB_Name = "frmTotMovimentati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private pSqlDettaglio As String
Private pDal As Date
Private pAl As Date

Private Sub cmd_Click(Index As Integer)
    Select Case Index
         Case 0
            Unload Me
         Case 1
            caricaTotaliNelPeriodo pDal, pAl
        Case 2
            ReportMovTotali.Stampa False, Me, pDal, pAl
    End Select
    
End Sub

Private Sub DTPicker1_Change(Index As Integer)
    If Index = 0 Then
        pDal = DTPicker1(0).Value
    Else
        pAl = DTPicker1(1).Value
    End If
 
End Sub

Private Sub Form_Load()
    Dim colH As ColumnHeader
    
    With lvwMovimenti(0)
        Set colH = .ColumnHeaders.Add(, , "Cod Cer")
        Set colH = .ColumnHeaders.Add(, , "Tip Rifiuto")
        Set colH = .ColumnHeaders.Add(, , "Qty Totale")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Tot Bonus")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Tot Punti")
        colH.Alignment = lvwColumnRight
    
    End With
    
    pDal = CDate(Year(Now) & "-01-01 00:00:00")
    pAl = Now
    
    DTPicker1(0).Value = pDal
    DTPicker1(1).Value = pAl
    
    caricaTotaliNelPeriodo pDal, pAl
    
    
End Sub

Private Sub caricaTotaliNelPeriodo(Dal As Date, al As Date)
    Dim itm As ListItem
    Dim cer As String
    
    If DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti.state = adStateOpen Then
        DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti.Close
    End If
    DevEnvMDB.cmdTotaliPerMovimentiBonusPunti Dal, al
    
    
    lvwMovimenti(0).ListItems.Clear

    Do While Not DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti.EOF
        
        cer = DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!codCategoria & _
              DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!codSottoCategoria & _
              DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!codCer & _
              DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!codSubCer & _
              DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!UnitaMisura
        With lvwMovimenti(0)
            Set itm = .ListItems.Add(, "|" & cer, cer)
            itm.ListSubItems.Add , DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!Descrizione.Name, DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!Descrizione
            itm.ListSubItems.Add , DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!quantita.Name, _
                    FormatNumber(DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!quantita.Value, 2)
            itm.ListSubItems.Add , DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!BonusDovuto.Name, _
                    FormatNumber(DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!BonusDovuto.Value, 2)
            itm.ListSubItems.Add , DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!PuntiDovuti.Name, _
                    FormatNumber(DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti!PuntiDovuti.Value, 2)

        
        End With
        
        DevEnvMDB.rscmdTotaliPerMovimentiBonusPunti.MoveNext
    Loop
    
    AllargaLeColonneLVW lvwMovimenti(0), False

End Sub




Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    
    On Error Resume Next
    
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    
    For i = pic.LBound To pic.ubound
        pic(i).Move 0, 0, Me.ScaleWidth, cmd(i).top - 50
    Next
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set pfrmTotMovimentiBonusPunti = Nothing

End Sub


Private Sub pic_Resize(Index As Integer)
    Dim topY As Integer
    Dim posHeight As Integer
    
    topY = 0
    If Index = 0 Then
        posHeight = pic(Index).ScaleHeight - Label2(0).Height - DTPicker1(0).Height - topY - 2 * offsetY
        
        DTPicker1(0).top = pic(Index).ScaleHeight - DTPicker1(0).Height - offsetY
        DTPicker1(1).top = pic(Index).ScaleHeight - DTPicker1(0).Height - offsetY
        Label2(0).top = DTPicker1(1).top - Label2(0).Height
        Label2(1).top = DTPicker1(1).top - Label2(0).Height
        lvwMovimenti(0).Move 0, topY, pic(Index).ScaleWidth, posHeight
    End If
    
    
    
End Sub



Public Sub RefreshME()
    caricaTotaliNelPeriodo pDal, pAl
End Sub
