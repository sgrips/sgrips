VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "TARI - Tassa Rifiuti "
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11505
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   4560
   End
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   7275
      Left            =   10890
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5490
      Top             =   2685
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7275
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18865
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   4860
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":08CA
            Key             =   "Icon_Ana"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0D1E
            Key             =   "Icon_Man"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":15FA
            Key             =   "Icon_Factory"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1ED6
            Key             =   "Icon_Null"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":21F2
            Key             =   "Icon_Domestica"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2ACE
            Key             =   "Icon_Attivita"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":33AA
            Key             =   "Icon_Personale"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":37FE
            Key             =   "Icon_Domestica_Attivita"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":40DA
            Key             =   "Icon_Costi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":452E
            Key             =   "Icon_Ammortamento"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":498A
            Key             =   "Icon_DatiTecnici"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5266
            Key             =   "Icon_PF"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5B42
            Key             =   "Icon_PV"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&Gestione"
      Begin VB.Menu mnuParametriAmbiente 
         Caption         =   "&Parametri Ambiente"
      End
      Begin VB.Menu mnudatabase 
         Caption         =   "&Data Base"
         Begin VB.Menu mnuComprimi 
            Caption         =   "&Comprimi DB"
         End
         Begin VB.Menu mnuPathDB 
            Caption         =   "Path &DB"
         End
         Begin VB.Menu mnuProtezione 
            Caption         =   "&Protezione"
         End
      End
      Begin VB.Menu mnuPubblica 
         Caption         =   "Pu&bblica"
      End
      Begin VB.Menu fff 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsportazione 
         Caption         =   "&Esportazione"
      End
      Begin VB.Menu mnuImportazione 
         Caption         =   "I&mportazione"
      End
      Begin VB.Menu rem03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImpStampante 
         Caption         =   "&Imposta Stampante"
      End
      Begin VB.Menu mnurem 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu mnuSchede 
      Caption         =   "&Schede"
      Begin VB.Menu mnuAnagrafica 
         Caption         =   "&Anagrafica"
      End
      Begin VB.Menu mnuEconomica 
         Caption         =   "&Economica"
      End
      Begin VB.Menu mnuDatiTecnici 
         Caption         =   "&Dati Tecnici"
      End
      Begin VB.Menu mnuDatiRaccolta 
         Caption         =   "Dati &Raccolta"
      End
      Begin VB.Menu rem01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtenzeDomestiche 
         Caption         =   "&Utenze Domestiche"
      End
      Begin VB.Menu mnuCategorieAttivita 
         Caption         =   "&Utenze Attività"
      End
      Begin VB.Menu mnuCAPersonali 
         Caption         =   "&Categorie Attività Personalizzate"
      End
   End
   Begin VB.Menu mnuCostiIndivisibili 
      Caption         =   "&Costi Indivisibili"
      Visible         =   0   'False
      Begin VB.Menu mnuCostiIndiUtenzeDomestiche 
         Caption         =   "Utenze &Domestiche"
      End
      Begin VB.Menu mnuCostiIndiUtenzeAttivita 
         Caption         =   "Utenze &Attività"
      End
   End
   Begin VB.Menu mnuCacolo 
      Caption         =   "&Calcolo"
      Begin VB.Menu mnuCoefficientiItem 
         Caption         =   "&Coefficienti "
      End
      Begin VB.Menu mnuRiduzioni 
         Caption         =   "&Riduzioni"
      End
      Begin VB.Menu rem33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalcoloTariffa 
         Caption         =   "&Calcolo Tari"
      End
   End
   Begin VB.Menu mnuSimulazioni 
      Caption         =   "S&imulazioni"
      Begin VB.Menu mnuSimulazione 
         Caption         =   "&Gestione Simulazioni"
      End
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "&Report"
      Begin VB.Menu mnuStampe 
         Caption         =   "&Stampe"
      End
      Begin VB.Menu mnuStatisticheDes 
         Caption         =   "S&tatistiche"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatistiche 
         Caption         =   "&Descrizione Statistiche"
      End
   End
   Begin VB.Menu mnuGrafici 
      Caption         =   "&Grafici"
      Begin VB.Menu mnuTariffa507 
         Caption         =   "&Utenza Domestica c.652"
         Begin VB.Menu mnuGraUte507 
            Caption         =   "&Per Componente"
            Index           =   9
         End
      End
      Begin VB.Menu mnuUtenzaDomesticaGraf 
         Caption         =   "&Utenza Domestica c.651"
         Begin VB.Menu mnuGraUte 
            Caption         =   "&1 Componente"
            Index           =   1
         End
         Begin VB.Menu mnuGraUte 
            Caption         =   "&2 Componenti"
            Index           =   2
         End
         Begin VB.Menu mnuGraUte 
            Caption         =   "&3 Componenti"
            Index           =   3
         End
         Begin VB.Menu mnuGraUte 
            Caption         =   "&4 Componenti"
            Index           =   4
         End
         Begin VB.Menu mnuGraUte 
            Caption         =   "&5 Componenti"
            Index           =   5
         End
         Begin VB.Menu mnuGraUte 
            Caption         =   "&Maggiore di 5 Componenti"
            Index           =   6
         End
      End
      Begin VB.Menu mnuGraUtenzaAttivita 
         Caption         =   "&Utenza Attività c.651"
         Begin VB.Menu mnuGraAtt 
            Caption         =   "&Attività Fascia 1"
            Index           =   7
         End
         Begin VB.Menu mnuGraAtt 
            Caption         =   "&Attività Fascia 2"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnuFinestra 
      Caption         =   "&Finestra"
      WindowList      =   -1  'True
      Begin VB.Menu mnuDisponiFinestra 
         Caption         =   "Affianca &Orrizontalmente"
         Index           =   1
      End
      Begin VB.Menu mnuDisponiFinestra 
         Caption         =   "Affianca Verticalmente"
         Index           =   2
      End
      Begin VB.Menu mnuDisponiFinestra 
         Caption         =   "&Sovrapponi"
         Index           =   3
      End
      Begin VB.Menu mnuDisponiFinestra 
         Caption         =   "Disponi &Icone"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuInformazioni 
         Caption         =   "&Informazioni su "
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCreaSimulazione 
         Caption         =   "Crea Simulazione"
      End
      Begin VB.Menu mnuCancellaSimulazione 
         Caption         =   "Cancella Simulazione"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    lWidth = pic.Width
    lHeight = pic.Height
    
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ChiudiConnessione
End Sub

Private Sub mnuAnagrafica_Click()
    frmDatiAnagrafici.Show
End Sub

Private Sub mnuCalcoloTariffa_Click()
    FrmCalcoloTariffa.Avvio ', , tipoCalcolo:=mOptMetodoTariffa
End Sub

Private Sub mnuCancellaSimulazione_Click()
    frmSimulazioni.mnuCancella_Click
End Sub

Private Sub mnuCAPersonali_Click()
    FrmCategorieAttivitaPersonali.Carica
End Sub

Private Sub mnuCategorieAttivita_Click()
    FrmCategorieAttivita.Carica
End Sub


Private Sub mnuCoefficientiItem_Click()
    FrmCoefficiente.Carica
End Sub

Private Sub mnuComprimi_Click()
    On Error GoTo mnuComprimi_ClickERR
    
    Screen.MousePointer = vbHourglass
    DoEvents
    ChiudiConnessione
    If Dir(App.Path & "\Comp.mdb") <> "" Then Kill App.Path & "\Comp.mdb"
    
    gJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gPathDB & ";Jet OLEDB:Database Password=" + gPasswordDB _
    , "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Comp.mdb"
    
    Kill gPathDB
    Name App.Path & "\Comp.mdb" As gPathDB
    
    ApriConnessione
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "Compressione DataBase effettuata", vbOKOnly, App.Title
    Exit Sub
mnuComprimi_ClickERR:
    MsgBox Error$, vbCritical, App.Title
    On Error GoTo 0
    On Error Resume Next
    ApriConnessione
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuCostiIndiUtenzeAttivita_Click()
    frmCategorieAttivitaCostiInd.Carica
End Sub

Private Sub mnuCostiIndiUtenzeDomestiche_Click()
    frmUtenzeDomesticheCostiIndivi.Carica
End Sub

Private Sub mnuCreaSimulazione_Click()
    frmSimulazioni.mnuCrea_Click
End Sub

Private Sub mnuDatiRaccolta_Click()
    frmDatiRaccolta.Show
End Sub

Private Sub mnuDatiTecnici_Click()
    frmDatiTecnici.Show
End Sub

Private Sub mnuDisponiFinestra_Click(index As Integer)
    Me.Arrange Choose(index, vbCascade, vbTileHorizontal, vbTileVertical, vbArrangeIcons)
End Sub

Private Sub mnuEconomica_Click()
    FrmCostiDelComune.Carica
End Sub

Private Sub mnuEsci_Click()
    Unload Me
End Sub

Private Sub mnuEsportazione_Click()
    frmEsportazoine.Carica
End Sub


Private Sub mnuGraUte_Click(index As Integer)
    If pGrafici(index) Is Nothing Then
        Set pGrafici(index) = New frmGrafici
        pGrafici(index).ShowChart index
    End If
End Sub



Private Sub mnuGraAtt_Click(index As Integer)
    mnuGraUte_Click index
End Sub

Private Sub mnuGraUte507_Click(index As Integer)
    If pGrafici(index) Is Nothing Then
        Set pGrafici(index) = New frmGrafici
        pGrafici(index).ShowChart index
    End If
End Sub

Private Sub mnuImportazione_Click()
    frmImportazioneXLS.Show vbModeless
End Sub

Private Sub mnuImpStampante_Click()
    Me.CommonDialog1.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDNoPageNums
    Me.CommonDialog1.ShowPrinter
End Sub

Private Sub mnuInformazioni_Click()
    frmAbout.Show vbModal
End Sub



Private Sub mnuParametriAmbiente_Click()
    frmParametriSistema.Carica
End Sub

Public Sub mnuPathDB_Click()
    Dim sfile As String
    
    DoCommonDialog DT_Browse, sfile, "DDD", "Indicare Percorso DataBase", Me.hwnd, "StudioAbiente.MDB", , App.Path, "*.mdb"
    
    If sfile = "" Then sfile = App.Path
    SaveSetting App.Title, "DATABASE", "PATH", sfile
    
    ChiudiConnessione
    Do While Not ApriConnessione
        If MsgBox("DataBase non trovato in [" & sfile & "]." & vbCrLf & "Indicare il percorso dove si trova installato", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Unload Me
            Exit Sub
        End If
        DoCommonDialog DT_Browse, sfile, "DDD", "Indicare Percorso DataBase", Me.hwnd, "StudioAbiente.MDB", , App.Path, "*.mdb"
        If sfile = "" Then sfile = App.Path
        SaveSetting App.Title, "DATABASE", "PATH", sfile
    Loop
   
End Sub

Private Sub mnuProtezione_Click()
    frmProtezioneDB.Show vbModal
End Sub

Private Sub mnuPubblica_Click()
    frmPubblica.Show vbModal
End Sub

Private Sub mnuRiduzioni_Click()
    frmBonus.Carica
End Sub

Private Sub mnuSimulazione_Click()
    frmSimulazioni.showME
End Sub

Private Sub mnuStampe_Click()
    
    frmStampe.Carica
End Sub

Private Sub mnuStatistiche_Click()
    frmStatistiche.Carica
End Sub

Private Sub mnuStatisticheDes_Click()
    frmStatisticheComunali.Carica
End Sub

Private Sub mnuUtenzeDomestiche_Click()
    FrmUtenzeDomestiche.Carica
End Sub

Private Sub Pic_Paint()
    If Nopaint Then Exit Sub
    DoEvents
    pic.Cls
    DrawBackgroundPattern pic.hwnd, Horizontal
    rotate pic, "  " & App.ProductName & " " & App.LegalCopyright & " ver. " & App.Major & "." & App.Minor & "." & App.Revision, 23
    DoEvents
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)


    On Error Resume Next
    gidxInfo = gidxInfo - 1
    If gidxInfo < 0 Then gidxInfo = UBound(gInfoTxt)
    Select Case Panel.index
        Case 1
            StatusBar1.Panels(2).Text = gInfoTxt(gidxInfo)
    End Select
End Sub


Public Sub Set_Parent(hWndChild As Long)
    SetParent hWndChild, Me.hwnd
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    gUnloadMdi = True
    If frmAbout.MustRegister Then
        frmAbout.Show vbModal
    Else
        gUnloadMdi = False
        Unload frmAbout
    End If
    If gUnloadMdi Then
        Unload Me
    End If
End Sub

Public Sub attivaTimer()
    Timer1.Enabled = True
End Sub

