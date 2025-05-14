VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatiAnagrafici 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anagrafica Comunale"
   ClientHeight    =   7830
   ClientLeft      =   7200
   ClientTop       =   4155
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Stampa"
      Height          =   375
      Index           =   5
      Left            =   3375
      TabIndex        =   7
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdFiltro 
      Caption         =   "&Filtro"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Inserisci"
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   5
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modifica"
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   4
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancella"
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvwComuni 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
Attribute VB_Name = "frmDatiAnagrafici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFiltro_Click()
    
    frmFiltroComune.Show vbModal
    If gcmdAnnulla = False Then CaricaInListView Me.LvwComuni, ICON_ANAGRAFICA, Lvw_AnagraficaComuni
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim sCodComune As String
    Dim rs As Recordset
    Select Case Index
        Case 0
            If ScritturaSuDB(LvwComuni, "AnagraficaComunale") Then
                'conferma
                Unload Me
            End If
        Case 1
            'annulla
            Unload Me
        Case 3
            'modifica
            LvwComuni_DblClick
        Case 2
            'cancella
            If LvwComuni.SelectedItem Is Nothing Then Exit Sub
            LvwComuni.SelectedItem.Ghosted = Not LvwComuni.SelectedItem.Ghosted
        Case 4
            'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
            If gComuneAttivo = ALL_COMUNI Then
                frmComuneDaLavorare.Show vbModal
                sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
            Else
                sCodComune = gComuneAttivo
            End If
            
            'non si può inserire lo stesso comune nello stesso anno
            Set rs = gDBC.Execute("Select * from AnagraficaComunale where AnnoRiferimento='" & Year(Now) & "' and CodComune='" & sCodComune & "'")
            If rs.EOF Then
                gDBC.Execute "Insert into AnagraficaComunale (CodComune,AnnoRiferimento) values ('" & sCodComune & "','" & Year(Now) & "')"
                gWhereFiltroSuComune = vbNullString
                CaricaInListView Me.LvwComuni, ICON_ANAGRAFICA, Lvw_AnagraficaComuni
                LvwComuni.SelectedItem = LvwComuni.ListItems(CercaKeyInListItems(LvwComuni.ListItems, SEPARATORE & "CodComune='" & sCodComune & "'" & SEPARATORE & "AnnoRiferimento='" & Year(Now) & "'" & SEPARATORE))
                LvwComuni.SelectedItem.EnsureVisible
                frmEditAnagrafica.Edita LvwComuni
            Else
                MsgBox "Dati già presenti nel periodo in corso", vbInformation + vbOKOnly, App.Title
            End If
            
        Case 5
            frmStampe.Carica LvwComuni.SelectedItem, LvwComuni.SelectedItem.ListSubItems(1), LvwComuni.SelectedItem.ListSubItems(2)
    End Select
End Sub

Private Sub Form_Load()
    MDIMain.mnuComprimi.Enabled = False
    LvwComuni.SmallIcons = MDIMain.imglist
    
    cmdFiltro.Visible = False
    
    'Metto sotto gestione di sincronia il data base per potere eventualmente annullare le modifiche
    'gDBC.BeginTrans
    
    CentraME Me

    
    'se getione ALLCOMUNI
    If gComuneAttivo = ALL_COMUNI Then
        cmdFiltro.Visible = True
    End If
    CaricaInListView Me.LvwComuni, ICON_ANAGRAFICA, Lvw_AnagraficaComuni
End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub LvwComuni_DblClick()
    
    'edit del record LVW corrente
    If LvwComuni.SelectedItem Is Nothing Then Exit Sub
    frmEditAnagrafica.Edita LvwComuni

End Sub



Private Sub LvwComuni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Command1_Click (2)
    End If
End Sub

