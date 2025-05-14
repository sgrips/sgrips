VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatiTecnici 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dati Tecnici del Comune"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
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
   Begin MSComctlLib.ListView Lvw 
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
Attribute VB_Name = "frmDatiTecnici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFiltro_Click()
    
    
    frmFiltroComune.Show vbModal
    If gcmdAnnulla = False Then CaricaInListView Me.lvw, ICON_DATITECNICI, Lvw_DatiTecnici
    
End Sub

Private Sub scriviDatiAggiuntivi()
    Dim QRifiutiUtenzaDomestica As Double
    Dim ProduzioneMediaRifiuti As Double
    Dim QRifiutiPerAbitante As Double
    Dim CostoPerGestioneKGSmaltiti As Double
    Dim ProduzionePerUtenzaDomestica As Double
    Dim ProduzionePerUtenzaAttivita As Double
    Dim i As Integer
    Dim sWhere As String
    Dim rs As Recordset
    
    
    With lvw.ListItems
        For i = 1 To .Count
            If .Item(i).Ghosted = False Then
                sWhere = " Where " & GetPkeyForWhereCondition(.Item(i).Key)
                
                Set rs = DatiCostiComune(.Item(i).Text, .Item(i).ListSubItems(2).Text)
                
                
                
                QRifiutiUtenzaDomestica = 0
                ProduzioneMediaRifiuti = 0
                QRifiutiPerAbitante = 0
                CostoPerGestioneKGSmaltiti = 0
                ProduzionePerUtenzaDomestica = 0
                ProduzionePerUtenzaAttivita = 0
                
                QRifiutiUtenzaDomestica = .Item(i).ListSubItems("QRifiutiTotali") - .Item(i).ListSubItems("QRifiutiCategorieAttivita")
                
                If .Item(i).ListSubItems("QRifiutiTotali") <> 0 And (Not rs.EOF And Not rs.BOF) Then
                    CostoPerGestioneKGSmaltiti = rs("Tot_servizio") / .Item(i).ListSubItems("QRifiutiTotali")
                End If
                    
                If (CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaDomestica")) + CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaAttivita"))) <> 0 Then ProduzioneMediaRifiuti = .Item(i).ListSubItems("QRifiutiTotali") / (CLng(.Item(i).ListSubItems("TotSuperficieUtenzaDomestica")) + CLng(.Item(i).ListSubItems("TotSuperficieUtenzaAttivita")))
                If (CDbl(.Item(i).ListSubItems("NRuoliUtenzaDomestica") + .Item(i).ListSubItems("NRuoliUtenzaAttivita")) + CDbl(.Item(i).ListSubItems("NRuoliSecondaCasa"))) <> 0 Then QRifiutiPerAbitante = .Item(i).ListSubItems("QRifiutiTotali") / (CLng(.Item(i).ListSubItems("NRuoliUtenzaDomestica")) + CLng(.Item(i).ListSubItems("NRuoliUtenzaAttivita")) + CLng(.Item(i).ListSubItems("NRuoliSecondaCasa")))
                If CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaDomestica")) <> 0 Then ProduzionePerUtenzaDomestica = QRifiutiUtenzaDomestica / CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaDomestica"))
                If CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaAttivita")) <> 0 Then ProduzionePerUtenzaAttivita = CDbl(.Item(i).ListSubItems("QRifiutiCategorieAttivita")) / CDbl(.Item(i).ListSubItems("TotSuperficieUtenzaAttivita"))
                
                
                gDBC.Execute "update DatiTecniciComune set QRifiutiUtenzaDomestica=" & Replace(QRifiutiUtenzaDomestica, ",", ".") & _
                         ",ProduzioneMediaRifiuti=" & Replace(ProduzioneMediaRifiuti, ",", ".") & ",QRifiutiPerAbitante=" & Replace(QRifiutiPerAbitante, ",", ".") & _
                         ",CostoPerGestioneKGSmaltiti=" & Replace(CostoPerGestioneKGSmaltiti, ",", ".") & _
                         ",ProduzionePerUtenzaDomestica=" & Replace(ProduzionePerUtenzaDomestica, ",", ".") & ",ProduzionePerUtenzaAttivita=" & Replace(ProduzionePerUtenzaAttivita, ",", ".") & _
                         sWhere & " and flgSimulazione=false "

            End If
            
            
        Next
    End With
End Sub


Private Sub Command1_Click(index As Integer)
    Dim sCodComune As String
    Dim rs As ADODB.Recordset
    Dim ParAbitanti As String
    Dim OkInserimento As Boolean
    

    Select Case index
        Case 0
            If ScritturaSuDB(lvw, "DatiTecniciComune") Then
                'gDBC.CommitTrans
                'conferma
                                
                scriviDatiAggiuntivi
                
                Unload Me
            Else
                'gDBC.RollbackTrans
                'ri inizializzo la tranzazione
                'gDBC.BeginTrans
            End If
        Case 1
            'annulla
            'gDBC.RollbackTrans
            Unload Me
        Case 3
            'modifica
            Lvw_DblClick
        Case 2
            'cancella
            If lvw.SelectedItem Is Nothing Then Exit Sub
            lvw.SelectedItem.Ghosted = Not lvw.SelectedItem.Ghosted
        Case 4
            'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
            If gComuneAttivo = ALL_COMUNI Then
                frmComuneDaLavorare.Show vbModal
                sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
            Else
                sCodComune = gComuneAttivo
            End If
            
            
            'Per inserire i dati tecnici del comune occorre prima aver inserito i dati anagrafici e
            'economici
            Set rs = DatiAnagraficiComunali(sCodComune, Year(Now), ParAbitanti)
            OkInserimento = Not rs.EOF
            Set rs = DatiCostiComune(sCodComune, Year(Now))
            OkInserimento = Not rs.EOF And OkInserimento
            
            If Not OkInserimento Then
                MsgBox "Occorre inserire i dati Anagrafici ed Economici del comune" & vbCrLf & "prima dell'inserimento dei Dati Tecnici", vbExclamation + vbOKOnly, App.Title
            Else
                
                'non si può inserire lo stesso comune nello stesso anno
                Set rs = gDBC.Execute("Select * from DatiTecniciComune where AnnoRiferimento='" & Year(Now) & "' and CodComune='" & sCodComune & "'")
                If rs.EOF Then
                    gDBC.Execute "Insert into DatiTecniciComune (CodComune,AnnoRiferimento) values ('" & sCodComune & "','" & Year(Now) & "')"
                    gWhereFiltroSuComune = vbNullString
                    CaricaInListView Me.lvw, ICON_DATITECNICI, Lvw_DatiTecnici
                    lvw.SelectedItem = lvw.ListItems(CercaKeyInListItems(lvw.ListItems, SEPARATORE & "CodComune='" & sCodComune & "'" & SEPARATORE & "AnnoRiferimento='" & Year(Now) & "'" & SEPARATORE))
                    lvw.SelectedItem.EnsureVisible
                    frmEditDatiTecnici.Edita lvw
                Else
                    MsgBox "Dati già presenti nel periodo in corso", vbInformation + vbOKOnly, App.Title
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    MDIMain.mnuComprimi.Enabled = False
    lvw.SmallIcons = MDIMain.imglist
    
    cmdFiltro.Visible = False
    
    'Metto sotto gestione di sincronia il data base per potere eventualmente annullare le modifiche
    'gDBC.BeginTrans
    
    CentraME Me

    
    'se getione ALLCOMUNI
    If gComuneAttivo = ALL_COMUNI Then
        cmdFiltro.Visible = True
    End If
    CaricaInListView Me.lvw, ICON_DATITECNICI, Lvw_DatiTecnici
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub Lvw_DblClick()
    
    'edit del record LVW corrente
    If lvw.SelectedItem Is Nothing Then Exit Sub
    frmEditDatiTecnici.Edita lvw

End Sub



Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Command1_Click (2)
    End If
End Sub
