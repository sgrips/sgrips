VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Easy Recycling - "
   ClientHeight    =   9615
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12510
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrInvioAccessi 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   3960
      Top             =   6600
   End
   Begin VB.Timer TimerFocus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   6600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   6600
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9240
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "User"
            Object.ToolTipText     =   "Utente Collegato"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8573
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3096
            Text            =   "Accessi Da Inviare:0"
            TextSave        =   "Accessi Da Inviare:0"
            Key             =   "accessiInv"
            Object.ToolTipText     =   "Accessi da inviare attraverso la funzione di esportazione"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3149
            Text            =   "Accessi Non Inviati:0"
            TextSave        =   "Accessi Non Inviati:0"
            Key             =   "accessiNonInv"
            Object.ToolTipText     =   "Accessi  Inviati ma rifiutati per problemi anagrafici"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   794
            MinWidth        =   527
            Picture         =   "MDIMain.frx":08CA
            Key             =   "statopesa1"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   794
            MinWidth        =   527
            Picture         =   "MDIMain.frx":0C12
            Key             =   "statopesa2"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   794
            MinWidth        =   529
            Picture         =   "MDIMain.frx":0F5A
            Key             =   "statopesa3"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   794
            MinWidth        =   527
            Picture         =   "MDIMain.frx":12A2
            Key             =   "statopesa4"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   794
            MinWidth        =   527
            Picture         =   "MDIMain.frx":15EA
            Key             =   "statosbarra"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   12510
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   12450
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   582
         ButtonWidth     =   2249
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Uscita"
               Key             =   "Uscita"
               Description     =   "Uscita dal programma"
               Object.ToolTipText     =   "Uscita dal programma"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Login"
               Key             =   "Login"
               Description     =   "Login"
               Object.ToolTipText     =   "Cambia il Login utente"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Movimenti"
               Key             =   "Pesate"
               Description     =   "Gestione Pesate"
               Object.ToolTipText     =   "Gestione Movimenti Isola Ecologica"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "aqpesata"
                     Text            =   "Gestione Ingresso"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "uscita"
                     Text            =   "Gestione Uscita"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "elpesate"
                     Text            =   "Visualizza Movimenti"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "espbonus"
                     Text            =   "Esporta Bonus"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anagrafica"
               Key             =   "Anagrafica"
               Description     =   "Gestione anagrafiche"
               Object.ToolTipText     =   "Gestione anagrafiche"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Destinatari"
                     Text            =   "Destinatari"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Trasportatori"
                     Text            =   "Trasportatori"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Interm"
                     Text            =   "Intermediari"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   8040
            ScaleHeight     =   375
            ScaleWidth      =   4335
            TabIndex        =   3
            Top             =   120
            Width           =   4335
            Begin VB.Label lblInfoStampaMultipla 
               Alignment       =   1  'Right Justify
               Caption         =   "Modalità Stampa Multi Pesata Attiva"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   0
               TabIndex        =   4
               ToolTipText     =   "Click per inviare i movimenti fino ad ora registrati alla stampante"
               Top             =   0
               Visible         =   0   'False
               Width           =   4305
            End
         End
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   1560
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RTSEnable       =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1932
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8256
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":C4E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuGestione 
      Caption         =   "&Gestione"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnusep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "&Impostazioni"
         Begin VB.Menu mnuSocieta 
            Caption         =   "&Societa"
         End
         Begin VB.Menu mnuUtentiLogin 
            Caption         =   "&Utenti Login"
         End
         Begin VB.Menu mnurem 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImpostGenerali 
            Caption         =   "&Generali"
         End
         Begin VB.Menu mnuImpPesa 
            Caption         =   "&Pesa"
         End
         Begin VB.Menu mnuEsporazioneImp 
            Caption         =   "&Export/Import"
         End
      End
      Begin VB.Menu mnudatabase 
         Caption         =   "&Data Base"
         Begin VB.Menu mnuBackupDB 
            Caption         =   "&Backup/Restore"
         End
         Begin VB.Menu mnuStoricizza 
            Caption         =   "&Storicizza"
         End
         Begin VB.Menu ddd 
            Caption         =   "-"
         End
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
      Begin VB.Menu mnu01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUscita 
         Caption         =   "&Uscita"
      End
   End
   Begin VB.Menu mnuImpExp 
      Caption         =   "&Import/Export"
      Begin VB.Menu mnuEsportazione 
         Caption         =   "&Esportazione"
         Begin VB.Menu mnuMovimentiEsporta 
            Caption         =   "&Movimenti"
         End
         Begin VB.Menu mnuFattureEsp 
            Caption         =   "&Fatture"
         End
         Begin VB.Menu mnuAccessi 
            Caption         =   "&Accessi"
         End
      End
      Begin VB.Menu mnuImportazione 
         Caption         =   "&Importazione"
         Begin VB.Menu MnuImmobiliImp 
            Caption         =   "&Immobili - Utenti"
         End
         Begin VB.Menu mnuEcopunto 
            Caption         =   "&Ecopunto"
         End
      End
   End
   Begin VB.Menu mnuTabelle 
      Caption         =   "&Tabelle"
      Begin VB.Menu mnuDestinatari 
         Caption         =   "&Destinatari"
      End
      Begin VB.Menu mnuTrasportatori 
         Caption         =   "&Trasportatori"
      End
      Begin VB.Menu mnuIntermediari 
         Caption         =   "&Intermediari"
      End
      Begin VB.Menu gg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPremi 
         Caption         =   "&Premi"
      End
      Begin VB.Menu rem03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtente 
         Caption         =   "&Utenti Isola Ecologica"
      End
      Begin VB.Menu mnuCodiciCER 
         Caption         =   "&Codici CER"
      End
   End
   Begin VB.Menu mnuGenScuole 
      Caption         =   "&Scuole"
      Begin VB.Menu mnuScuole 
         Caption         =   "&Anagrafica"
      End
      Begin VB.Menu mnuMovimentiScuole 
         Caption         =   "&Movimenti"
      End
   End
   Begin VB.Menu mnudati 
      Caption         =   "&Movimenti"
      Begin VB.Menu mnuPesate 
         Caption         =   "&In Ingresso"
      End
      Begin VB.Menu mnuMovUscita 
         Caption         =   "In &Uscita"
      End
      Begin VB.Menu mnuGiacenze 
         Caption         =   "&Giacenze"
      End
      Begin VB.Menu mnuTotPerCer 
         Caption         =   "&Totali Movimentati"
      End
      Begin VB.Menu hhhh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrafici 
         Caption         =   "&Grafici"
         Begin VB.Menu mnGrafico 
            Caption         =   "&Bonus Anno"
            Index           =   1
         End
         Begin VB.Menu mnGrafico 
            Caption         =   "&Quantità Annua"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuDocumenti 
      Caption         =   "&Documenti"
      Begin VB.Menu mnuRegScaCar 
         Caption         =   "&Registri di Carico/Scarico"
      End
      Begin VB.Menu mnuFormDiTras 
         Caption         =   "&Formulari Di Trasporto"
      End
      Begin VB.Menu mnuMudEsp 
         Caption         =   "&MUD"
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
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pInvioAccessiIncorso As Boolean



Private Sub lblInfoStampaMultipla_Click()
    If lblInfoStampaMultipla.Visible = True And Not pfrmPesata Is Nothing Then
        pfrmPesata.StampaMovimenti True
    End If
End Sub

Private Sub MDIForm_Load()
    Dim rs As ADODB.Recordset
    
    Load frmIsolaEcologicaAttiva
        
    Set Toolbar1.ImageList = ImageList1
    Set Toolbar1.DisabledImageList = ImageList1
    Set Toolbar1.HotImageList = ImageList1
    
    Toolbar1.Buttons("Pesate").Image = 1
    Toolbar1.Buttons("Uscita").Image = 2
    Toolbar1.Buttons("Login").Image = 3
    Toolbar1.Buttons("Anagrafica").Image = 4
        
'    If gStartMode = 1 Then
'        Toolbar1.Buttons("Pesate").Caption = ""
'        Toolbar1.Buttons("Uscita").Caption = ""
'        Toolbar1.Buttons("Login").Caption = ""
'        Toolbar1.Buttons("Anagrafica").Caption = ""
'    End If
        
    Toolbar1.Buttons("Pesate").Visible = Not gAnagrafiche
    Toolbar1.Buttons("Anagrafica").Visible = Not gAnagrafiche
        
    mnuProtezione.Visible = frmLogin.codUser = "GFA"
                
    On Error Resume Next
    Set rs = gDBC.Execute("Select * from Societa where idSocieta = " & gIdSocietaProduttore)
    
    If rs.EOF Then
        Me.Caption = "Easy RD - Impostare la Società "
    Else
        Me.Caption = "Easy RD - " & GetCurrentComune("Comune").value
    End If
    
    StatusBar.Panels(1).Text = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision

    Set StatusBar.Panels("statopesa1").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\NotConnect.gif")
    Set StatusBar.Panels("statopesa2").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\NotConnect.gif")
    Set StatusBar.Panels("statopesa3").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\NotConnect.gif")
    Set StatusBar.Panels("statopesa4").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\NotConnect.gif")
    Set StatusBar.Panels("statosbarra").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\NotConnect.gif")



End Sub

Public Sub refreshForm()
    On Error Resume Next
    Dim f As Form
    
    For Each f In Forms
        f.RefreshME
    Next
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mnuUscita_Click
    End If
End Sub

Private Sub MDIForm_Resize()
    Picture1.left = Me.ScaleWidth - Picture1.ScaleWidth - 100
    
    If gStartMode = 0 Then
        If Not pfrmPesata Is Nothing Then
            pfrmPesata.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If
        If Not pfrmUscita Is Nothing Then
            pfrmUscita.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If
        If Not pfrmSocieta Is Nothing Then
            pfrmSocieta.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If
        If Not pfrmScuole Is Nothing Then
            pfrmScuole.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If
        If Not pfrmPremi Is Nothing Then
            pfrmPremi.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If
        If Not pfrmMovimentiScuole Is Nothing Then
            pfrmMovimentiScuole.Move CoolBar1.left, CoolBar1.Height, Me.ScaleWidth, Me.ScaleHeight
        End If

    End If
    Dim m As Form
    
    If Me.WindowState = 1 Or Me.WindowState = 2 Then
        For Each m In Forms
            If m.WindowState <> Me.WindowState Then
                m.WindowState = IIf(Me.WindowState = 1, Me.WindowState, 0)
            End If
        Next
        
        TimerFocus.Interval = 200
        TimerFocus.enabled = True
        
    End If
    

    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload frmIsolaEcologicaAttiva
    chiudiApplicazione

End Sub


Private Sub mnGrafico_Click(Index As Integer)
    If pfrmGrafici(Index) Is Nothing Then
        Set pfrmGrafici(Index) = New frmGrafici
        pfrmGrafici(Index).ShowChart Index
    Else
        pfrmGrafici(Index).SetFocus
    End If
End Sub

Private Sub mnuAccessi_Click()
    tmrInvioAccessi.enabled = False
    frmEsportazioneAccessi.Show vbModal
    tmrInvioAccessi.enabled = True
End Sub

Private Sub mnuBackupDB_Click()
    frmBackupDB.Show vbModal
End Sub


Private Sub mnuCodiciCER_Click()
    
    If pfrmCodiciCER Is Nothing Then
        Set pfrmCodiciCER = New frmCodiciCER
        pfrmCodiciCER.Show
    Else
        pfrmCodiciCER.SetFocus
    End If
    
End Sub


Private Sub mnuComprimi_Click()
    On Error GoTo mnuComprimi_ClickERR
    
    
    
    
    If Forms.Count > 3 Then
        MsgBox "Prima di utilizzare questa funzione occorre uscire da tutte le videate aperte", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    ComprimiDB
    
    
    Exit Sub
mnuComprimi_ClickERR:
    MsgBox "mnuComprimi_Click:" & Error$, vbCritical, App.Title
    On Error GoTo 0
    On Error Resume Next
    ApriConnessione
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuDestinatari_Click()
    If pfrmDestinatari Is Nothing Then
        Set pfrmDestinatari = New frmDestinatari
        pfrmDestinatari.Show
    Else
        pfrmDestinatari.SetFocus
    End If
    
    
End Sub

Private Sub mnuDisponiFinestra_Click(Index As Integer)
    Me.Arrange Choose(Index, vbCascade, vbTileHorizontal, vbTileVertical, vbArrangeIcons)
End Sub

Private Sub mnuEcopunto_Click()
On Error Resume Next
    frmEcopunto.Show vbModal
End Sub

Private Sub mnuEsporazioneImp_Click()
    frmImpEsp.Show vbModal
End Sub

Private Sub mnuFattureEsp_Click()
    frmEsportazioneFatt.Show vbModal

End Sub

Private Sub mnuFormDiTras_Click()

    If pfrmFormulari Is Nothing Then
        Set pfrmFormulari = New frmFormulari
        pfrmFormulari.Show
    Else
        pfrmFormulari.SetFocus
    End If

End Sub

Private Sub mnuGiacenze_Click()
    If pfrmGiacenze Is Nothing Then
        Set pfrmGiacenze = New frmGiacenze
        pfrmGiacenze.Show
    Else
        pfrmGiacenze.SetFocus
    End If
End Sub

Private Sub mnuMovimentiEsporta_Click()
    frmEsportazioneBonus.Show
End Sub

Private Sub mnuTotPerCer_Click()
    If pfrmTotMovimentiBonusPunti Is Nothing Then
        Set pfrmTotMovimentiBonusPunti = New frmTotMovimentati
        pfrmTotMovimentiBonusPunti.Show
    Else
        pfrmTotMovimentiBonusPunti.SetFocus
    End If

End Sub


Private Sub MnuImmobiliImp_Click()

    
    If pfrmImpImmobiliUtenti Is Nothing Then
        Set pfrmImpImmobiliUtenti = New frmImpImmobiliUtenti
        pfrmImpImmobiliUtenti.Show
    Else
        pfrmImpImmobiliUtenti.SetFocus
    End If
End Sub

Private Sub mnuImpostGenerali_Click()
    frmImpostazioni.Show vbModal
End Sub

Private Sub mnuImpPesa_Click()
    frmImpostazioniPesa.Show vbModal
End Sub


Private Sub mnuInformazioni_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuIntermediari_Click()
    If pfrmIntermediari Is Nothing Then
        Set pfrmIntermediari = New frmIntermediari
        pfrmIntermediari.Show
    Else
        pfrmIntermediari.SetFocus
    End If
    
End Sub



Private Sub mnuLogin_Click()
    frmLogin.ShowME frmLogin.codUser
End Sub



Private Sub mnuMovimentiScuole_Click()
    Dim lStyle As Long
    If pfrmMovimentiScuole Is Nothing Then
        Set pfrmMovimentiScuole = New frmMovimentiScuole
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmMovimentiScuole.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmMovimentiScuole.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmMovimentiScuole.hWnd, Me.hWnd
                pfrmMovimentiScuole.left = 0
                pfrmMovimentiScuole.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmMovimentiScuole.ShowME
    Else
        pfrmMovimentiScuole.SetFocus
    End If
End Sub

Private Sub mnuMovUscita_Click()
    Dim lStyle As Long
    If pfrmUscita Is Nothing Then
        Set pfrmUscita = New frmUscita
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmUscita.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmUscita.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmUscita.hWnd, Me.hWnd
                pfrmUscita.left = 0
                pfrmUscita.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmUscita.ShowME
    Else
        pfrmUscita.SetFocus
    End If
    
    
    
End Sub

Private Sub mnuPathDB_Click()
    apriDataBase
End Sub

Public Sub apriDataBase()
    Dim sFile As String
    
    
    DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso DataBase", Me.hWnd, "IsolaEcologica.MDB", , App.Path, "*.mdb"
    
    
    If sFile = "" Then sFile = App.Path
    SaveSetting App.Title, "DATABASE", "PATH", sFile
    
    ChiudiConnessione
    Do While Not ApriConnessione
        If MsgBox("DataBase non trovato in [" & sFile & "]." & vbCrLf & "Indicare il percorso dove si trova installato", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Unload Me
            Exit Sub
        End If
        DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso DataBase", Me.hWnd, "StudioAbiente.MDB", , App.Path, "*.mdb"
        If sFile = "" Then sFile = App.Path
        SaveSetting App.Title, "DATABASE", "PATH", sFile
    Loop


End Sub


Private Sub mnuMovimentiEsp_Click()
    On Error GoTo mnuMovimentiEsp_ClickERR
    Dim lStyle As Long
    If pfrmPesata Is Nothing Then
        Set pfrmPesata = New frmPesata
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmPesata.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmPesata.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmPesata.hWnd, Me.hWnd
                pfrmPesata.left = 0
                pfrmPesata.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmPesata.ShowME "elpesate", 1
    Else
        pfrmPesata.SetFocus
    End If
Exit Sub
mnuMovimentiEsp_ClickERR:
    ScriviLog TPL_ERRORE, App.Title, Err.Description, "mnuMovimentiEsp_Click"
End Sub

Private Sub mnuPesate_Click()
    On Error GoTo mnuPesate_ClickERR
    Dim lStyle As Long
    If pfrmPesata Is Nothing Then
        Set pfrmPesata = New frmPesata
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmPesata.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmPesata.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmPesata.hWnd, Me.hWnd
                pfrmPesata.left = 0
                pfrmPesata.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmPesata.ShowME "aqpeso"
    Else
        pfrmPesata.SetFocus
    End If
    Exit Sub
mnuPesate_ClickERR:
    ScriviLog TPL_ERRORE, App.Title, Err.Description, "mnuPesate_Click"
End Sub



Private Sub mnuProtezione_Click()
    frmProtezioneDB.Show vbModal
End Sub

Private Sub mnuRegScaCar_Click()
    If pfrmRegistri Is Nothing Then
        Set pfrmRegistri = New frmRegistri
        pfrmRegistri.Show
    Else
        pfrmRegistri.SetFocus
    End If
End Sub

Private Sub mnuScuole_Click()
    Dim lStyle As Long
    If pfrmScuole Is Nothing Then
        Set pfrmScuole = New FrmScuoleTab
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmScuole.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmScuole.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmScuole.hWnd, Me.hWnd
                pfrmScuole.left = 0
                pfrmScuole.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmScuole.ShowME
    Else
        pfrmScuole.SetFocus
    End If
End Sub


Private Sub mnuPremi_Click()
       Dim lStyle As Long
    If pfrmPremi Is Nothing Then
        Set pfrmPremi = New frmPremi
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmPremi.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmPremi.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmPremi.hWnd, Me.hWnd
                pfrmPremi.left = 0
                pfrmPremi.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmPremi.ShowME
    Else
        pfrmPremi.SetFocus
    End If
End Sub

Private Sub mnuSocieta_Click()
    Dim lStyle As Long
    If pfrmSocieta Is Nothing Then
        Set pfrmSocieta = New FrmSocietaTab
        
        If gStartMode = 1 Or gStartMode = 0 Then
            lStyle = GetWindowLong(pfrmSocieta.hWnd, GWL_STYLE)
            lStyle = lStyle And (Not MY_WS_DOCK)
            SetWindowLong pfrmSocieta.hWnd, GWL_STYLE, lStyle
            If gStartMode = 0 Then
                SetParent pfrmSocieta.hWnd, Me.hWnd
                pfrmSocieta.left = 0
                pfrmSocieta.top = 0
                MDIForm_Resize
            End If
        End If
        
        pfrmSocieta.ShowME
    Else
        pfrmSocieta.SetFocus
    End If
    
End Sub

Public Sub CaricaFormPesa()
    mnuPesate_Click
End Sub


Public Sub CaricaSocieta()
    mnuSocieta_Click
End Sub

Private Sub mnuStoricizza_Click()
    frmStoricizzazione.Show vbModal
End Sub


Private Sub mnuTrasportatori_Click()
    If pfrmTrasportatori Is Nothing Then
        Set pfrmTrasportatori = New frmTrasportatori
        pfrmTrasportatori.Show
    Else
        pfrmTrasportatori.SetFocus
    End If
    
End Sub

Private Sub mnuUscita_Click()
    Dim backup As Boolean
    
    frmAttesa.Show
    frmAttesa.SetFocus
        
    DoEvents
        
    Do While pInvioAccessiIncorso
     DoEvents
    Loop
    
    Unload frmAttesa
    
    On Error Resume Next
    tmrInvioAccessi.enabled = False
    Timer1.enabled = False
    TimerFocus.enabled = False
    
    backup = True
    If gRichiedereAUscitaProgramma Then
        If MsgBox("Si vuole effettuare il backup prima di uscire ?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            backup = False
        End If
    Else
        backup = False
    End If
    
    If backup Then
        frmBackupDB.BackupRestore True, gRichiedereAUscitaProgramma
        Unload frmBackupDB
    End If

    Unload MDIMain
End Sub

Private Sub mnuUtente_Click()
    If pfrmUtenti Is Nothing Then
        Set pfrmUtenti = New frmUtenti
        pfrmUtenti.Show
    Else
        pfrmUtenti.SetFocus
    End If
    
End Sub

Private Sub mnuUtentiLogin_Click()
    If pfrmUser Is Nothing Then
       Set pfrmUser = New frmUser
       pfrmUser.Show
    Else
        pfrmUser.SetFocus
    End If
End Sub

Private Sub MSComm1_OnComm(Index As Integer)
Dim ComEv As Long
    Exit Sub
    ComEv = MSComm1(Index).CommEvent
    Select Case ComEv
    
    ' Handle each event or error by placing
    ' code below each case statement
        Case comEventCDTO
            ScriviLog TPL_ERRORE, App.Title, "comEventCDTO", "COM"
        Case comEventCTSTO  ' CTS Timeout.
            ScriviLog TPL_ERRORE, App.Title, "comEventCTSTO", "COM"

        Case comEventFrame  ' Framing Error
        Case comEventOverrun    ' Data Lost.
            ScriviLog TPL_ERRORE, App.Title, "comEventOverrun", "COM"
        Case comEventRxParity   ' Parity Error.
            ScriviLog TPL_ERRORE, App.Title, "comEventRxParity", "COM"
        Case comEventTxFull ' Transmit buffer full.
            ScriviLog TPL_ERRORE, App.Title, "comEventTxFull", "COM"
        
        Case comEventDCB    ' Unexpected error retrieving DCB]
            ScriviLog TPL_ERRORE, App.Title, "comEventDCB", "COM"
    
        Case comEvCD    ' Change in the CD line.
            PesaCollegata Index + 1, ""
        
        Case comEvCTS   ' Change in the CTS line.
            'PesaCollegata Index + 1
        Case comEvDSR   ' Change in the DSR line.
        
            PesaCollegata Index + 1, ""
        Case comEvRing    ' Change in the Ring Indicator.
        Case comEvReceive ' Received RThreshold # of chars.
        Case comEvSend  ' There are SThreshold number of characters in the transmit
                        ' buffer.
        Case comEvEOF   ' An EOF charater was found in
                        ' the input stream
    End Select
End Sub

Private Sub Timer1_Timer()
    If lblInfoStampaMultipla.Visible And DateDiff("s", Timer1.Tag, Now) > 2 Then
        'attendo due secondi prima di nasconderlo
        lblInfoStampaMultipla.Visible = False
        lblInfoStampaMultipla.Refresh
    ElseIf Not lblInfoStampaMultipla.Visible And DateDiff("s", Timer1.Tag, Now) > 0.5 Then
        lblInfoStampaMultipla.Visible = True
        lblInfoStampaMultipla.Refresh
        Timer1.Tag = Now
    End If
    DoEvents
End Sub

Public Sub attivaMultiPesata(b As Boolean)
    Timer1.enabled = b
    lblInfoStampaMultipla.Visible = False
    Timer1.Tag = Now
    Timer1.Interval = 500
End Sub


Private Sub TimerFocus_Timer()
    TimerFocus.enabled = False
    Dim f As Form
    For Each f In Forms
        On Error Resume Next
        If f.MDIChild = False Then f.SetFocus
        
    Next
End Sub

Private Sub tmrInvioAccessi_Timer()
    Dim rec As ADODB.Recordset
    Dim frmEspAcc As frmEsportazioneAccessi
    
    On Error Resume Next
   
    pInvioAccessiIncorso = True

    If gTipoInvioAccessi = Protocollo_Invio_HTTP Then
        Set frmEspAcc = New frmEsportazioneAccessi
        frmEspAcc.InvioAutomatico
        Unload frmEspAcc
    End If
    
    Set rec = gDBC.Execute("TRANSFORM Count(*) AS Espr1 " & _
                "select q.header from ( " & _
                "SELECT 1 AS row,'ok' as statoC, 'statoInvio' as header FROM RilevazioneAccessi where stato = '0' " & _
                "Union all " & _
                "SELECT 1 AS row,'ko' as statoD, 'statoInvio' as header FROM RilevazioneAccessi where stato = '3' " & _
                ") as q " & _
                "GROUP BY 1,header PIVOT statoC")
    
    Dim daInviare  As Integer
    Dim nonInviati  As Integer
    
    daInviare = rec("ok").value
    nonInviati = rec("ko").value
    
    MDIMain.AggiornaNumAccessiDaInviare daInviare, nonInviati
    
    pInvioAccessiIncorso = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "Uscita"
             mnuUscita_Click
        Case "Login"
             mnuLogin_Click
        Case "Pesate"
            mnuPesate_Click
        Case "Anagrafica"
             frmDestinatari.Show

             
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo Toolbar1_ButtonMenuClickERR
    Select Case ButtonMenu.key
        Case "aqpesata"
            If pfrmPesata Is Nothing Then
                Set pfrmPesata = New frmPesata
                pfrmPesata.ShowME "aqpeso"
            Else
                pfrmPesata.SetFocus
            End If

        Case "elpesate"
            If pfrmPesata Is Nothing Then
                Set pfrmPesata = New frmPesata
                pfrmPesata.ShowME "elpesate"
            Else
                pfrmPesata.SetFocus
            End If
            
        Case "uscita"
            mnuMovUscita_Click
        Case "espbonus"
             mnuMovimentiEsporta_Click
        Case "Destinatari"
             mnuDestinatari_Click
        Case "Trasportatori"
             mnuTrasportatori_Click
        Case "Interm"
             mnuIntermediari_Click
    End Select
Exit Sub
Toolbar1_ButtonMenuClickERR:
    ScriviLog TPL_ERRORE, App.Title, Err.Description, "Toolbar1_ButtonMenuClick"

End Sub

Public Function PesaCollegata(idx As Integer, Des As String) As Boolean
    Dim s  As String
    PesaCollegata = pesa(idx).IsConnect(s)
    Set Me.StatusBar.Panels("statopesa" & idx).Picture = LoadPicture(App.Path & "\Icons\IconsPgm\" & IIf(PesaCollegata, "Connect.gif", "NotConnect.gif"))
    If Des <> "" Then Me.StatusBar.Panels("statopesa" & idx).ToolTipText = Des
    
End Function


Public Sub AggiornaNumAccessiDaInviare(accessiOK As Integer, accessiKO As Integer)
    StatusBar.Panels("accessiInv").Text = "Accessi Da Inviare:" & accessiOK
    StatusBar.Panels("accessiNonInv").Text = "Accessi Non Inviati:" & accessiKO
End Sub

Public Sub SbarraCollegata(attiva As Boolean, Des As String)
    Set Me.StatusBar.Panels("statosbarra").Picture = LoadPicture(App.Path & "\Icons\IconsPgm\" & IIf(attiva, "Connect.gif", "NotConnect.gif"))
    If Des <> "" Then Me.StatusBar.Panels("statosbarra").ToolTipText = Des
    
End Sub


Public Sub applicaRestrizioniUtente()
    MDIMain.mnudatabase.enabled = frmLogin.Livello <> LVL_USER
    MDIMain.mnuImpExp.enabled = frmLogin.Livello <> LVL_USER
    MDIMain.mnuUtentiLogin.enabled = frmLogin.Livello <> LVL_USER
End Sub

Public Sub attivaMenuSegreti()
    mnuProtezione.Visible = True
    DoEvents
End Sub

