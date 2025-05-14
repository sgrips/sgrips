VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCostiDelComune 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costi del Comune"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   9945
      TabIndex        =   5
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   1
      Left            =   8625
      TabIndex        =   4
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancella"
      Height          =   375
      Index           =   2
      Left            =   7305
      TabIndex        =   3
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modifica"
      Height          =   375
      Index           =   3
      Left            =   5985
      TabIndex        =   2
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Inserisci"
      Height          =   375
      Index           =   4
      Left            =   4665
      TabIndex        =   1
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdFiltro 
      Caption         =   "&Filtro"
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   7200
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvwCosti 
      Height          =   6975
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmCostiDelComune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pmodal  As Boolean
Dim pParentwin As Long

Private Sub cmdFiltro_Click()
    frmFiltroComune.Show vbModal
    If gcmdAnnulla = False Then CaricaInListView Me.LvwCosti, ICON_COSTI, Lvw_CostiComunali
    
End Sub




Private Sub Command1_Click(index As Integer)
    Dim sCodComune As String
    Dim rs As Recordset
    Dim ParAbitanti As String
    Dim OkInserimento As Boolean
    
    Select Case index
        Case 0
            If ScritturaSuDB(LvwCosti, "CostoComunale") Then
                'conferma
                Unload Me
            End If
        Case 1
            'annulla
            Unload Me
        Case 3
            'modifica
            lvwcosti_DblClick
        Case 2
            'cancella
            If LvwCosti.SelectedItem Is Nothing Then Exit Sub
            LvwCosti.SelectedItem.Ghosted = Not LvwCosti.SelectedItem.Ghosted
        Case 4
            'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
            If gComuneAttivo = ALL_COMUNI Then
                frmComuneDaLavorare.Show vbModal
                sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
            Else
                sCodComune = gComuneAttivo
            End If
            
            'Per inserire i Costi del Comune  occorre prima aver inserito i dati anagrafici
            Set rs = DatiAnagraficiComunali(sCodComune, Year(Now), ParAbitanti)
            OkInserimento = Not rs.EOF
            
            If Not OkInserimento Then
                MsgBox "Occorre inserire i dati Anagrafici del comune" & vbCrLf & "prima dell'inserimento dei Dati Economici", vbExclamation + vbOKOnly, App.Title
            Else
            
            
                'non si può inserire lo stesso comune nello stesso anno
                Set rs = gDBC.Execute("Select * from CostoComunale where AnnoRiferimento='" & Year(Now) & "' and CodComune='" & sCodComune & "'")
                If rs.EOF Then
                    gDBC.Execute "Insert into CostoComunale (CodComune,AnnoRiferimento) values ('" & sCodComune & "','" & Year(Now) & "')"
                    gWhereFiltroSuComune = vbNullString
                    CaricaInListView Me.LvwCosti, ICON_COSTI, Lvw_CostiComunali
                    LvwCosti.SelectedItem = LvwCosti.ListItems(CercaKeyInListItems(LvwCosti.ListItems, SEPARATORE & "CodComune='" & sCodComune & "'" & SEPARATORE & "AnnoRiferimento='" & Year(Now) & "'" & SEPARATORE))
                    LvwCosti.SelectedItem.EnsureVisible
                    frmEditCostiDelComune.Edita LvwCosti, Me, True
                Else
                    MsgBox "Dati già presenti nel periodo in corso", vbInformation + vbOKOnly, App.Title
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    MDIMain.mnuComprimi.Enabled = False
    LvwCosti.SmallIcons = MDIMain.imglist
    
    cmdFiltro.Visible = False
    
    'Metto sotto gestione di sincronia il data base per potere eventualmente annullare le modifiche
    ''gDBC.BeginTrans
    
    CentraME Me

    
    'se getione ALLCOMUNI
    If gComuneAttivo = ALL_COMUNI Then
        cmdFiltro.Visible = True
    End If
    
    

End Sub

'carica il costo di un detterminato comune in modalità edit
Public Sub Carica(Optional sCodComune As String = vbNullString, Optional AnnoRiferimento As String = vbNullString, Optional tipoCalcolo As tpoTariffa = TariffaNonImpostata, Optional modal As Boolean = True, Optional parentwin As Long)
    Dim datiSimulazione As Boolean
    
    
    
    If sCodComune <> vbNullString Then
        If gWhereSuComune = vbNullString Then
          gWhereFiltroSuComune = "Where "
        Else
          gWhereFiltroSuComune = " And "
        End If
        gWhereFiltroSuComune = gWhereFiltroSuComune & " AnnoRiferimento ='" & AnnoRiferimento & "' and  CodComune ='" & sCodComune & "' and flgSimulazione=true "
        datiSimulazione = True
        Me.Visible = False
    End If
    
    pmodal = modal
    pParentwin = parentwin
    
    CaricaInListView Me.LvwCosti, ICON_COSTI, Lvw_CostiComunali, datiSimulazione
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gWhereFiltroSuComune = vbNullString
    MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub lvwcosti_DblClick()
    
    'edit del record LVW corrente
    If LvwCosti.SelectedItem Is Nothing Then Exit Sub
    
    
    frmEditCostiDelComune.Edita LvwCosti, Me, pmodal

End Sub


Public Function GetEditForm() As Form

    Dim m_frmEditCostiDelComune As frmEditCostiDelComune
    
    Set m_frmEditCostiDelComune = New frmEditCostiDelComune
    
    SetParent m_frmEditCostiDelComune.hwnd, pParentwin

    m_frmEditCostiDelComune.Edita LvwCosti, Me, pmodal
    
    Set GetEditForm = m_frmEditCostiDelComune
End Function


Private Sub LvwCosti_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Command1_Click (2)
    End If
End Sub
