VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatiRaccolta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anagrafica Dati Raccolta Porta Porta"
   ClientHeight    =   10410
   ClientLeft      =   7200
   ClientTop       =   4335
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFiltro 
      Caption         =   "&Filtro"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Inserisci"
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   5
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modifica"
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   4
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancella"
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   3
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   1
      Left            =   11040
      TabIndex        =   2
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   12360
      TabIndex        =   1
      Top             =   9960
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvwRaccolta 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   16960
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
Attribute VB_Name = "frmDatiRaccolta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFiltro_Click()
    
    frmFiltroComune.Show vbModal
    If gcmdAnnulla = False Then CaricaInListView Me.LvwRaccolta, ICON_DATIRACCOLTA, Lvw_AnagraficaRaccolta
    
End Sub

Private Sub Command1_Click(index As Integer)
    Dim sCodComune As String
    Dim rs As Recordset
    
    Select Case index
        Case 0
            If ScritturaSuDB(LvwRaccolta, "AnagraficaRaccolta") Then
                'conferma
                Unload Me
            End If
        Case 1
            'annulla
            Unload Me
        Case 3
            'modifica
            LvwRaccolta_DblClick
        Case 2
            'cancella
            If LvwRaccolta.SelectedItem Is Nothing Then Exit Sub
            LvwRaccolta.SelectedItem.Ghosted = Not LvwRaccolta.SelectedItem.Ghosted
        Case 4
            'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
            If gComuneAttivo = ALL_COMUNI Then
                frmComuneDaLavorare.Show vbModal
                sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
            Else
                sCodComune = gComuneAttivo
            End If
            
            'non si può inserire lo stesso comune nello stesso anno
            Set rs = gDBC.Execute("Select * from AnagraficaRaccolta where AnnoRiferimento='" & Year(Now) & "' and CodComune='" & sCodComune & "'")
            If rs.EOF Then
                gDBC.Execute "Insert into AnagraficaRaccolta (CodComune,AnnoRiferimento) values ('" & sCodComune & "','" & Year(Now) & "')"
                gWhereFiltroSuComune = vbNullString
                CaricaInListView LvwRaccolta, ICON_DATIRACCOLTA, Lvw_AnagraficaRaccolta
                LvwRaccolta.SelectedItem = LvwRaccolta.ListItems(CercaKeyInListItems(LvwRaccolta.ListItems, SEPARATORE & "CodComune='" & sCodComune & "'" & SEPARATORE & "AnnoRiferimento='" & Year(Now) & "'" & SEPARATORE))
                LvwRaccolta.SelectedItem.EnsureVisible
                frmEditRaccolta.Edita LvwRaccolta
            Else
                MsgBox "Dati già presenti nel periodo in corso", vbInformation + vbOKOnly, App.Title
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    MDIMain.mnuComprimi.Enabled = False
    LvwRaccolta.SmallIcons = MDIMain.imglist
    
    cmdFiltro.Visible = False
    

    
    CentraME Me

    
    'se getione ALLCOMUNI
    If gComuneAttivo = ALL_COMUNI Then
        cmdFiltro.Visible = True
    End If
    CaricaInListView LvwRaccolta, ICON_DATIRACCOLTA, Lvw_AnagraficaRaccolta
End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub LvwRaccolta_DblClick()
    
    'edit del record LVW corrente
    If LvwRaccolta.SelectedItem Is Nothing Then Exit Sub
    frmEditRaccolta.Edita LvwRaccolta
    Me.Visible = True
End Sub



Private Sub LvwRaccolta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Command1_Click (2)
    End If
End Sub

