VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEsportazioneBonus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Esportazione"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ricarica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   11640
      TabIndex        =   10
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Già Esportati"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwMovEsp 
      Height          =   6015
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Esporta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   12840
      TabIndex        =   7
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   14040
      TabIndex        =   6
      Top             =   7080
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   16515073
      CurrentDate     =   39357
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   16515073
      CurrentDate     =   39357
   End
   Begin VB.Label lblEsportazione 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Avanzamento Esportazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   11340
   End
   Begin VB.Label lblEsportazione 
      AutoSize        =   -1  'True
      Caption         =   "Al"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   315
      Width           =   150
   End
   Begin VB.Label lblEsportazione 
      AutoSize        =   -1  'True
      Caption         =   "Dal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   285
   End
End
Attribute VB_Name = "frmEsportazioneBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChecked As Long
Private mWhere As String

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
            If MsgBox("Attenzine procedendo con l'esportazione i movimenti " & vbCrLf _
                & " passeranno dallo stato da contabilizzare a contabilizzato. Continuare con l'operazione ?", vbExclamation + vbYesNo, App.Title) = VbMsgBoxResult.vbYes Then
            
                Command1(0).enabled = False
                Command1(1).enabled = False
                Command1(2).enabled = False
                
                Esportazione
                
                Command1(0).enabled = True
                Command1(1).enabled = True
                Command1(2).enabled = True
            End If
        
        Case 2
            CaricaMovimenti
    End Select
End Sub

Public Sub ShowME(tipoForm As Long)
    Load Me
    
    If tipoForm = 0 Then
    
    Else
    
    End If
    
    Me.Show
End Sub

Private Sub Form_Load()
    lvwMovEsp.ColumnHeaders.Clear
    lvwMovEsp.ColumnHeaders.Add , , "Cod.Utente"
    lvwMovEsp.ColumnHeaders.Add , , "Utente"
    lvwMovEsp.ColumnHeaders.Add , , "C.Fiscale/P.IVA"
    lvwMovEsp.ColumnHeaders.Add , , "Badge"
    lvwMovEsp.ColumnHeaders.Add , , "Cod.Immobile"
    lvwMovEsp.ColumnHeaders.Add , , "Immobile"
    lvwMovEsp.ColumnHeaders.Add , , "Targa"
    lvwMovEsp.ColumnHeaders.Add , , "Cod.Cer"
    lvwMovEsp.ColumnHeaders.Add , , "Tipo Rifiuto"
    lvwMovEsp.ColumnHeaders.Add , , "Data Mov."
    lvwMovEsp.ColumnHeaders.Add , , "Quantità"
    lvwMovEsp.ColumnHeaders(lvwMovEsp.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovEsp.ColumnHeaders.Add , , "U.M."
    lvwMovEsp.ColumnHeaders(lvwMovEsp.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovEsp.ColumnHeaders.Add , , "Bonus €"
    lvwMovEsp.ColumnHeaders(lvwMovEsp.ColumnHeaders.Count).Alignment = lvwColumnRight
    lvwMovEsp.ColumnHeaders.Add , , "Bonus Punti"
    lvwMovEsp.ColumnHeaders(lvwMovEsp.ColumnHeaders.Count).Alignment = lvwColumnRight

    
    DTPicker1(0).value = "01-01-" & Year(Date)
    DTPicker1(1).value = Date
    
    
    DTPicker1(0).value = Null
    DTPicker1(1).value = Null
    
    CaricaMovimenti
End Sub


Private Sub Esportazione()
    Dim item As ListItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim rs As ADODB.Recordset
    Dim qry As ADODB.Command
    Dim codComuneSocietaProd As String
    
    codComuneSocietaProd = GetCurrentComune("CodIstat").value
   
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gMovimentiFile & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
    
    If mChecked = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = mChecked
    ProgressBar1.value = 0
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "select * from Badge  where codUtente=? and Default=true"
    Set qry.ActiveConnection = gDBC
    qry.Prepared = True
    qry.Parameters.Append qry.CreateParameter("A", adInteger, adParamInput)
    gDBC.BeginTrans
    For Each item In lvwMovEsp.ListItems
        If (item.Checked) Then
            
            ProgressBar1.value = ProgressBar1.value + 1
            qry.Parameters(0).value = item.Text
            Set rs = qry.Execute
            If Not rs.EOF Then
 
            
                txtStream.WriteLine item.ListSubItems("idMovimento") & gMovimentiSeparatore & _
                    codComuneSocietaProd & gMovimentiSeparatore & _
                    Format$(item.ListSubItems("codPesa"), "000") & gMovimentiSeparatore & _
                    rs!codBadge & gMovimentiSeparatore & _
                    Format$(item.ListSubItems("DataMovimento"), "dd/MM/yyyy") & gMovimentiSeparatore & _
                    Format$(item.ListSubItems("DataMovimento"), "hh\:nn\:ss") & gMovimentiSeparatore & _
                    item.ListSubItems("quantita") & gMovimentiSeparatore & _
                    item.ListSubItems("CodiceCer") & gMovimentiSeparatore & _
                    "0" & gMovimentiSeparatore & _
                    item.Text & gMovimentiSeparatore & _
                    item.ListSubItems("TipoAnagrafe") & gMovimentiSeparatore & _
                    item.ListSubItems("CodiceImmobile") & gMovimentiSeparatore & _
                    item.ListSubItems("CodFiscaleIva")
                    
    
                'modifico lo stato dei movimenti
                gDBC.Execute "update movimenti set stato='1' where id=" & Mid$(item.key, 2)
            End If
            rs.Close
            Set rs = Nothing
        End If
    Next
    
    Set qry.ActiveConnection = Nothing
    Set qry = Nothing
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    Dim nomefile As String
    nomefile = gPercorsoEsp & "\" & gMovimentiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
    fso.MoveFile fileEXP, nomefile
    
    If gFTP_Active Then
        'invio dati via FTP
        Dim FTP As frmFTP
        Set frmFTP = New frmFTP
  
        frmFTP.FileDestination = "out/" & fso.GetFileName(nomefile)
        frmFTP.FileSource = nomefile
        frmFTP.Rename = True
        frmFTP.DoWork 1
        Set frmFTP = Nothing
    End If
    
    Set fso = Nothing
    
    gDBC.CommitTrans
    
    CaricaMovimenti
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title
        Exit Sub
EsportazioneERR:
    On Error Resume Next
    gDBC.RollbackTrans
End Sub

Private Sub CaricaMovimenti()
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    
    
    Screen.MousePointer = vbHourglass
    'mWhere = " (BonusDovuto > 0 or PuntiDovuti > 0) "
    mWhere = vbNullString
    If (Not IsNull(DTPicker1(0).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "DataMovimento >= #" & Format$(DTPicker1(0).value, "YYYY-MM-DD 00:00:00") & "#"
    End If

    If (Not IsNull(DTPicker1(1).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "DataMovimento <= #" & Format$(DTPicker1(1).value, "YYYY-MM-DD 23:59:59") & "#"
    End If

    If (Check1.value = 1) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "Stato = '1'"
    Else
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "Stato = '0'"
    End If
    
    If mWhere <> "" Then mWhere = mWhere & " AND "
    mWhere = mWhere & "Movimenti.codComune = '" & GetCurrentComune("CodIstat").value & "'"


    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    rs.Source = "SELECT Utenti.Cognome,badge.codBadge, Utenti.Nome,Utenti.Cognome,Utenti.CodFiscale,Utenti.PartitaIva,Movimenti.CodImmobile, Utenti.RagioneSociale,Movimenti.id,Comuni.Comune, Immobili.Indirizzo, Immobili.NumCivico," & _
            "Immobili.Subalterno, Immobili.NumInterno, Immobili.Scala, Immobili.Piano, " & _
            "Immobili.CodComune,(CodiciCerComunali.CodCategoria + CodiciCerComunali.CodSottoCategoria + CodiciCerComunali.CodCer + CodiciCerComunali.CodSubCer ) as CodiceCer, CodiciCerComunali.Descrizione AS Rifiuto,Movimenti.Manuale, CodiciCerComunali.FlgBonus, " & _
            "Movimenti.DataMovimento,Movimenti.Targa,Movimenti.NumeroFormulario,Movimenti.DataFormulario, Movimenti.Quantita,Movimenti.unitamisura, Movimenti.BonusDovuto,Movimenti.PuntiDovuti, Movimenti.CodUser, Movimenti.DataIns, " & _
            "StatoMovimento.descrizione AS descrizioneStato, Immobili.CodUtente,Movimenti.id,Movimenti.codPesa,Utenti.TipoAnagrafe " & _
            "FROM (((((Movimenti INNER JOIN Immobili ON Movimenti.CodImmobile = Immobili.CodImmobile and Movimenti.CodUtente = Immobili.CodUtente and  Immobili.CodComune =Movimenti.CodComune  ) " & _
            "INNER JOIN CodiciCerComunali ON  Movimenti.CodComune = CodiciCerComunali.CodComune and Movimenti.CodSubCer = CodiciCerComunali.CodSubCer and (Movimenti.CodCer = CodiciCerComunali.CodCer) AND (Movimenti.CodSottoCategoria = CodiciCerComunali.CodSottoCategoria) AND (Movimenti.CodCategoria = CodiciCerComunali.CodCategoria)) " & _
            "INNER JOIN StatoMovimento ON Movimenti.Stato = StatoMovimento.codStato) " & _
            "INNER JOIN Comuni ON Movimenti.CodComune = Comuni.CODISTAT) " & _
            "INNER JOIN Utenti ON Utenti.CodUtente=Movimenti.CodUtente  and Utenti.CodComune=Movimenti.CodComune ) " & _
            "INNER JOIN badge ON  badge.CodUtente = Movimenti.CodUtente and  badge.CodComune=Movimenti.CodComune " & _
            "WHERE " & mWhere & " order by Immobili.codUtente,Movimenti.DataMovimento, iif (Utenti.Cognome = null ,Utenti.Cognome + Utenti.Nome, Utenti.RagioneSociale) "
    
    
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    mChecked = 0
    lvwMovEsp.ListItems.Clear
    Do While Not rs.EOF
        Set item = lvwMovEsp.ListItems.Add(, "_" & rs!id, rs!codUtente)
        item.ListSubItems.Add , , IIf(nullToBlank(rs!Cognome) = "", rs!ragioneSociale, rs!Cognome & " " & rs!Nome)
        item.ListSubItems.Add , "CodFiscaleIva", IIf(nullToBlank(rs!CodFiscale) = vbNullString, nullToBlank(rs!PartitaIva), nullToBlank(rs!CodFiscale))
        item.ListSubItems.Add , , rs!codBadge
        item.ListSubItems.Add , "CodiceImmobile", rs!codImmobile
        item.ListSubItems.Add , , rs!Indirizzo & "," & nullToBlank(rs!NumCivico) & " Sub:" & nullToBlank(rs!Subalterno) & " Num.Int:" & nullToBlank(rs!NumInterno)
        item.ListSubItems.Add , , nullToBlank(rs!Targa)
        item.ListSubItems.Add , "CodiceCer", rs!codiceCer
        item.ListSubItems.Add , "DesCer", rs!Rifiuto
        item.ListSubItems.Add , "DataMovimento", Format$(rs!DataMovimento, "dd-mmm-yyyy hh:nn:ss")
        item.ListSubItems.Add , "quantita", rs!quantita
        item.ListSubItems.Add , , rs!UnitaMisura
        item.ListSubItems.Add , , rs!BonusDovuto
        item.ListSubItems.Add , , rs!PuntiDovuti
        item.ListSubItems.Add , , rs!descrizioneStato
        item.ListSubItems.Add , , nullToBlank(rs!NumeroFormulario)
        item.ListSubItems.Add , , nullToBlank(rs!DataFormulario)
        item.ListSubItems.Add , "idMovimento", nullToBlank(rs!id)
        item.ListSubItems.Add , "codPesa", nullToBlank(rs!codPesa)
        item.ListSubItems.Add , "TipoAnagrafe", nullToBlank(rs!TipoAnagrafe)
        item.Checked = True
        rs.MoveNext
        
        mChecked = mChecked + 1
    Loop
    
    
    
    AllargaLeColonneLVW lvwMovEsp, False
    
    Screen.MousePointer = vbDefault
    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub lvwMovEsp_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked = True Then
        mChecked = mChecked + 1
    Else
        mChecked = mChecked - 1
    End If
End Sub
