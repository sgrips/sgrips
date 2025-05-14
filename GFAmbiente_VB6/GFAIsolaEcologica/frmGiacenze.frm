VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGiacenze 
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
      TabIndex        =   18
      Top             =   4440
      Width           =   4575
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
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
         Format          =   40566785
         CurrentDate     =   39945
      End
      Begin MSComctlLib.ListView lvwLotti 
         Height          =   2415
         Index           =   2
         Left            =   0
         TabIndex        =   19
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
         TabIndex        =   23
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
         Format          =   40566785
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      Index           =   3
      Left            =   7440
      TabIndex        =   17
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
      Index           =   2
      Left            =   8760
      TabIndex        =   16
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Tracciabilità"
      Enabled         =   0   'False
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
      Left            =   10080
      TabIndex        =   15
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   2
      Left            =   6000
      ScaleHeight     =   4185
      ScaleWidth      =   5025
      TabIndex        =   4
      Top             =   4440
      Width           =   5055
      Begin VB.Frame Frame1 
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5295
         Begin VB.OptionButton opt 
            Caption         =   "Tutti"
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
            Left            =   120
            TabIndex        =   9
            Top             =   350
            Width           =   735
         End
         Begin VB.OptionButton opt 
            Caption         =   "Con Giacenza"
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
            Left            =   960
            TabIndex        =   8
            Top             =   350
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Caption         =   "Scaricati"
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
            Index           =   2
            Left            =   2760
            TabIndex        =   7
            Top             =   350
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwLotti 
         Height          =   1575
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2778
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
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   4200
      ScaleHeight     =   3495
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   600
      Width           =   4575
      Begin VB.PictureBox picColor 
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   4635
         TabIndex        =   10
         Top             =   2640
         Width           =   4695
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giacenza In Scadenza"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   14
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lblColor 
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   1
            Left            =   2280
            TabIndex        =   13
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giacenza Scaduta"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   12
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label lblColor 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView lvwLotti 
         Height          =   2415
         Index           =   0
         Left            =   0
         TabIndex        =   3
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
   End
   Begin MSComctlLib.TabStrip TabGiacenze 
      Height          =   1335
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2355
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nel Periodo"
            Key             =   "Periodo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Totali"
            Key             =   "Totali"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dettaglio"
            Key             =   "Dettagli"
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
Attribute VB_Name = "frmGiacenze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single
Private pSqlDettaglio As String
Private Const COLOR_ALLARM = vbBlue
Private Const COLOR_EXPIRED = vbRed
Private pDal As Date
Private pAl As Date

Private Sub cmd_Click(Index As Integer)
    Select Case Index
         Case 0
            Unload Me
        Case 1
            If lvwLotti(1).SelectedItem Is Nothing Then Exit Sub
            
            TracciabilitaDaLotto Mid$(lvwLotti(1).SelectedItem.key, 2)
        Case 2
            If TabGiacenze.SelectedItem.Index = 1 Then
                caricaTotaliNelPeriodo pDal, pAl
            ElseIf TabGiacenze.SelectedItem.Index = 2 Then
                caricaTotali
            Else
                caricaDettaglio
            End If
        Case 3
            ReportGiacenze.Stampa False, Me, pDal, pAl
    End Select
    
End Sub

Private Sub TracciabilitaDaLotto(idLotto As Long)
    Dim qry As Command
    Dim codLottoUscita As String
    Dim rs As ADODB.Recordset
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "select LottoUscita From LegameLottoIngressoLottoUscita where idLottoEntrata= " & idLotto
    Set qry.ActiveConnection = gDBC
    Set rs = qry.Execute
    
    If rs.EOF Then Exit Sub
    codLottoUscita = rs(0)
    Set qry.ActiveConnection = Nothing
    rs.Close
    Set rs = Nothing
    Set qry = Nothing
    
    frmLottiLottoUscita.ShowMe codLottoUscita
    
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
    
    With lvwLotti(2)
        Set colH = .ColumnHeaders.Add(, , "Cod Cer")
        Set colH = .ColumnHeaders.Add(, , "Tip Rifiuto")
        Set colH = .ColumnHeaders.Add(, , "Qty Totale")
        colH.Alignment = lvwColumnRight
    
    End With
    
    
    With lvwLotti(0)
        Set colH = .ColumnHeaders.Add(, , "Cod Cer")
        Set colH = .ColumnHeaders.Add(, , "Tip Rifiuto")
        Set colH = .ColumnHeaders.Add(, , "Qty Totale")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Da Caricare")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Massima")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty Max Normativa")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Dta Primo Mov.")
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Ritiro")
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Rit Normativa")

    
    End With
    
    
    With lvwLotti(1)
        Set colH = .ColumnHeaders.Add(, , "Lotto")
        Set colH = .ColumnHeaders.Add(, , "Cod.Cer")
        Set colH = .ColumnHeaders.Add(, , "Tip.Rifiuto")
        Set colH = .ColumnHeaders.Add(, , "Qta.Ingresso")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty.Uscita")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Qty.Residuo")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Nr.Giorni Giac.")
        colH.Alignment = lvwColumnRight
        Set colH = .ColumnHeaders.Add(, , "Data Mod.")
        Set colH = .ColumnHeaders.Add(, , "Utente Mod.")
        Set colH = .ColumnHeaders.Add(, , "Data Ins.")
        Set colH = .ColumnHeaders.Add(, , "Utente Ins.")
    End With
    
    
    pSqlDettaglio = DevEnvMDB.Commands("cmdGiacenze").CommandText
    
    caricaDettaglio
    
    caricaTotali
    
    pDal = CDate(Year(Now) & "-01-01 00:00:00")
    pAl = Now
    
    DTPicker1(0).Value = pDal
    DTPicker1(1).Value = pAl
    
    caricaTotaliNelPeriodo pDal, pAl
    
    TabGiacenze.Tabs(1).Selected = True
   
    
End Sub

Private Sub caricaTotaliNelPeriodo(Dal As Date, al As Date)
    Dim itm As ListItem
    Dim cer As String
    
    If DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo.state = adStateOpen Then
        DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo.Close
    End If
    DevEnvMDB.cmdGiacenzeTotaliNelPeriodo Dal, al
    
    
    lvwLotti(2).ListItems.Clear

    Do While Not DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo.EOF
        
        cer = DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!codCategoria & _
              DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!codSottoCategoria & _
              DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!codCer
        With lvwLotti(2)
            Set itm = .ListItems.Add(, "|" & cer, cer)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!Descrizione.Name, DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!Descrizione
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!qtyResidua.Name, _
                    FormatNumber(DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!qtyResidua + DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo!QtyDaRegistrare, 2)

        
        End With
        
        DevEnvMDB.rscmdGiacenzeTotaliNelPeriodo.MoveNext
    Loop
    
    AllargaLeColonneLVW lvwLotti(2), False

End Sub


Private Sub caricaTotali()
    Dim itm As ListItem
    Dim cer As String
    
    If DevEnvMDB.rscmdGiacenzeTotali.state = adStateOpen Then
        DevEnvMDB.rscmdGiacenzeTotali.Close
    End If
    DevEnvMDB.cmdGiacenzeTotali gGGAllarmeGiacGestionale, gGGAllarmeGiacNormativa
    
    
    lvwLotti(0).ListItems.Clear

    Do While Not DevEnvMDB.rscmdGiacenzeTotali.EOF
        
        cer = DevEnvMDB.rscmdGiacenzeTotali!codCategoria & _
              DevEnvMDB.rscmdGiacenzeTotali!codSottoCategoria & _
              DevEnvMDB.rscmdGiacenzeTotali!codCer
        With lvwLotti(0)
            Set itm = .ListItems.Add(, "|" & cer, cer)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!Descrizione.Name, DevEnvMDB.rscmdGiacenzeTotali!Descrizione
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!qtyResidua.Name, FormatNumber(DevEnvMDB.rscmdGiacenzeTotali!qtyResidua + DevEnvMDB.rscmdGiacenzeTotali!QtyDaRegistrare, 2)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!QtyDaRegistrare.Name, FormatNumber(DevEnvMDB.rscmdGiacenzeTotali!QtyDaRegistrare, 2)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!qtyMassimaGiacenza.Name, toZero(DevEnvMDB.rscmdGiacenzeTotali!qtyMassimaGiacenza)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!qtyMassimaGiacenzaNormativa.Name, toZero(DevEnvMDB.rscmdGiacenzeTotali!qtyMassimaGiacenzaNormativa)
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!MinDiDataIns.Name, Format$(DevEnvMDB.rscmdGiacenzeTotali!MinDiDataIns, "dd/mm/yyyy")
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiro.Name, Format$(DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiro, "dd/mm/yyyy")
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiroNormativa.Name, Format$(DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiroNormativa, "dd/mm/yyyy")

        
            If (DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiro < Now) Then
                coloraRiga itm, COLOR_EXPIRED
                ElseIf (DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiroNormativa < Now) Then
                    coloraRiga itm, COLOR_EXPIRED
                ElseIf (DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiroAllarme < Now) Then
                    coloraRiga itm, COLOR_ALLARM
                ElseIf (DevEnvMDB.rscmdGiacenzeTotali!DataPrevistaRitiroNormativaAllarme < Now) Then
                    coloraRiga itm, COLOR_ALLARM
            End If
        End With
        
        DevEnvMDB.rscmdGiacenzeTotali.MoveNext
    Loop
    
    lblColor(0).BackColor = COLOR_EXPIRED
    lblColor(1).BackColor = COLOR_ALLARM
    AllargaLeColonneLVW lvwLotti(0), False
End Sub


Private Sub caricaDettaglio()
    Dim itm As ListItem
    Dim sql As String
    Dim iorderBy As Integer
    Dim orderBy As String
    Dim sData  As String
    
    If DevEnvMDB.rscmdGiacenze.state = adStateOpen Then
        DevEnvMDB.rscmdGiacenze.Close
    End If
    
    iorderBy = InStr(1, UCase(pSqlDettaglio), "ORDER BY ")
    orderBy = Mid$(pSqlDettaglio, iorderBy)
    
    sql = left$(pSqlDettaglio, iorderBy - 1) & " where "
    If (opt(1).Value = True) Then
        sql = sql & " [QtaIngresso] >0 and "
    ElseIf (opt(2).Value = True) Then
        sql = sql & " [QtaIngresso] = 0 and "
    End If
    
    sql = sql & " Lotti.Tipo='I' " & orderBy
    
    DevEnvMDB.Commands("cmdGiacenze").CommandText = sql
    
    DevEnvMDB.cmdGiacenze
    
    lvwLotti(1).ListItems.Clear
    
    Do While Not DevEnvMDB.rscmdGiacenze.EOF
        
        With lvwLotti(1)
            Set itm = .ListItems.Add(, "|" & DevEnvMDB.rscmdGiacenze!id, DevEnvMDB.rscmdGiacenze!codLotto)
            itm.ListSubItems.Add , "Cer", DevEnvMDB.rscmdGiacenze!codCategoria & _
                                            DevEnvMDB.rscmdGiacenze!codSottoCategoria & _
                                            DevEnvMDB.rscmdGiacenze!codCer
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!Descrizione.Name, DevEnvMDB.rscmdGiacenze!Descrizione
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!QtaIngresso.Name, Format$(DevEnvMDB.rscmdGiacenze!QtaIngresso, "#0.00")
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!QtaUscita.Name, DevEnvMDB.rscmdGiacenze!QtaUscita
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!qtaResidua.Name, Format$(DevEnvMDB.rscmdGiacenze!qtaResidua, "#0.00")
            
            If IsNull(DevEnvMDB.rscmdGiacenze!NrGiorniInGiacenza) Then
                sData = " "
            Else
                sData = DevEnvMDB.rscmdGiacenze!NrGiorniInGiacenza
            End If
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!NrGiorniInGiacenza.Name, sData
            
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!DataMod.Name, Format$(DevEnvMDB.rscmdGiacenze!DataMod, "dd/mm/yyyy")
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!UserMod.Name, DevEnvMDB.rscmdGiacenze!UserMod
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!dataIns.Name, Format$(DevEnvMDB.rscmdGiacenze!dataIns, "dd/mm/yyyy")
            itm.ListSubItems.Add , DevEnvMDB.rscmdGiacenze!UserIns.Name, DevEnvMDB.rscmdGiacenze!UserIns
        
        End With
        DevEnvMDB.rscmdGiacenze.MoveNext
    Loop

    AllargaLeColonneLVW lvwLotti(1), False

    If DevEnvMDB.rscmdGiacenze.state = adStateOpen Then
        DevEnvMDB.rscmdGiacenze.Close
        DevEnvMDB.Commands("cmdGiacenze").CommandText = pSqlDettaglio
    End If


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
    
    
    TabGiacenze.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmd(0).Height - 2 * offsetX
    
    For i = pic.LBound To pic.ubound
        pic(i).Move TabGiacenze.ClientLeft, TabGiacenze.ClientTop, TabGiacenze.ClientWidth, TabGiacenze.ClientHeight
    Next
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set pfrmGiacenze = Nothing

End Sub

Private Sub opt_Click(Index As Integer)
    caricaDettaglio
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
        lvwLotti(2).Move 0, topY, pic(Index).ScaleWidth, posHeight
    ElseIf Index = 2 Then
        topY = Frame1.top + Frame1.Height + 1 * offsetY
        posHeight = pic(Index).ScaleHeight - topY - 2 * offsetY
        lvwLotti(1).Move 0, topY, pic(Index).ScaleWidth, posHeight
    Else
        picColor.Move 0, pic(Index).ScaleHeight - picColor.ScaleHeight - offsetY
        posHeight = pic(Index).ScaleHeight - picColor.ScaleHeight - 2 * offsetY
        lvwLotti(0).Move 0, topY, pic(Index).ScaleWidth, posHeight
    End If
    
    
    
End Sub

Private Sub TabGiacenze_Click()
    Dim i As Integer
    
    If (TabGiacenze.SelectedItem Is Nothing) Then Exit Sub
    
    cmd(1).enabled = TabGiacenze.SelectedItem.Index = 3
    cmd(3).enabled = TabGiacenze.SelectedItem.Index = 1
    
    For i = pic.LBound To pic.ubound
        pic(i).Visible = False
    Next
    
    pic(TabGiacenze.SelectedItem.Index - 1).Visible = True
End Sub


Public Sub RefreshME()
    caricaTotali
    caricaDettaglio
    caricaTotaliNelPeriodo pDal, pAl
End Sub
