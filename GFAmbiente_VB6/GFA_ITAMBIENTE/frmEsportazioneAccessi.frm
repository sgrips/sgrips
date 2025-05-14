VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEsportazioneAccessi 
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
      Format          =   126025729
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
      Format          =   126025729
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
Attribute VB_Name = "frmEsportazioneAccessi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChecked As Long
Private mWhere As String
Private pOpzioneAvvio As Integer


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
                
            Command1(0).enabled = False
            Command1(1).enabled = False
            Command1(2).enabled = False
            
            Esportazione
            
            Command1(0).enabled = True
            Command1(1).enabled = True
            Command1(2).enabled = True
        
        Case 2
            CaricaMovimenti
    End Select
End Sub

Private Sub Form_Load()
    lvwMovEsp.ColumnHeaders.Clear
    lvwMovEsp.ColumnHeaders.Add , , "Badge"
    lvwMovEsp.ColumnHeaders.Add , , "Nominativo"
    lvwMovEsp.ColumnHeaders.Add , , "Data Ingresso"
    lvwMovEsp.ColumnHeaders.Add , , "Cod.Fiscale/IVA"
    lvwMovEsp.ColumnHeaders.Add , , "Ospite"
    lvwMovEsp.ColumnHeaders.Add , , "id"
    lvwMovEsp.ColumnHeaders.Add , , "stato"
    
    DTPicker1(0).value = "01-01-" & Year(Date)
    DTPicker1(1).value = Date
    
    
    DTPicker1(0).value = Null
    DTPicker1(1).value = Null
    
    If pOpzioneAvvio = 0 Then pOpzioneAvvio = 1
    
    CaricaMovimenti
End Sub

Public Sub InvioAutomatico()
    pOpzioneAvvio = 2
    Load Me
    Esportazione
End Sub


Private Sub Esportazione()
    Dim statoInvio As EnumStatoInvio
    Dim sKey() As String
    Dim item As ListItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim FlagFileCreato As Boolean
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gAccessiFile & IIf(gAccessiFileType = 0, ".TXT", ".CSV")
    
    If mChecked = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = mChecked
    ProgressBar1.value = 0
    gDBC.BeginTrans
    statoInvio = Invio_Ok
    FlagFileCreato = False
    
    'nella modalità HTTP il record ma marcato come inviato solo se siè riusciti ad inviarlo via webservice
    If (gTipoInvioAccessi = Protocollo_Invio_HTTP) Then
        Dim HTTP As frmHTTP
        
        Set frmHTTP = New frmHTTP
        frmHTTP.URLToken = gURLTokenSicurezza
        frmHTTP.URLWebService = gURLWebService
    End If
    
    For Each item In lvwMovEsp.ListItems
        sKey = Split(Mid$(item.key, 2), "|")

        If (item.Checked) Then
            ProgressBar1.value = ProgressBar1.value + 1
            
            If (gTipoInvioAccessi = Protocollo_Invio_HTTP) Then
                frmHTTP.WS_Barcode_cfisc = item.Text
                frmHTTP.WS_Cod_pda = "PDA_01"
                frmHTTP.WS_CodFisc = item.ListSubItems("CodFiscaleIva")
                frmHTTP.WS_Data_movim = Format$(item.ListSubItems("DataOraIngresso"), "dd/MM/yyyy")
                frmHTTP.WS_Flag_ospite = IIf(item.ListSubItems("Ospite") = "Si", "1", "0")
                frmHTTP.WS_Id_movim = item.ListSubItems("Id")
                frmHTTP.WS_Istat_comune = GetCurrentComune!codistat
                frmHTTP.WS_Ora_movim = Format$(item.ListSubItems("DataOraIngresso"), "hh\:nn\:ss")
                
                
                statoInvio = frmHTTP.DoWork(pOpzioneAvvio)
            End If
            
            If statoInvio = Invio_Ok Then
                txtStream.WriteLine GetCurrentComune!codistat & gAccessiSeparatore & item.Text & gAccessiSeparatore & _
                    Format$(item.ListSubItems("DataOraIngresso"), "dd/MM/yyyy") & gAccessiSeparatore & _
                    Format$(item.ListSubItems("DataOraIngresso"), "hh\:nn\:ss") & gAccessiSeparatore & _
                    item.ListSubItems("CodFiscaleIva") & gAccessiSeparatore & IIf(item.ListSubItems("Ospite") = "Si", "1", "0") & gAccessiSeparatore & _
                    item.ListSubItems("Id") & gAccessiSeparatore & "PDA_01"
    
                'modifico lo stato dei movimenti
                gDBC.Execute "update RilevazioneAccessi set stato='1' where id=" & sKey(0)
                FlagFileCreato = True
            ElseIf statoInvio = Invio_Errore_Rifiutato Then
                'modifico lo stato dei movimenti rifiutato
                gDBC.Execute "update RilevazioneAccessi set stato='3' where id=" & sKey(0)
            End If
        End If
    Next
    
    If (gTipoInvioAccessi = Protocollo_Invio_HTTP) Then
        On Error Resume Next
        Unload frmHTTP
        Set frmHTTP = Nothing
    End If
    
    If (FlagFileCreato) Then
        
        If Not fso.FolderExists(gPercorsoEsp) Then
            fso.CreateFolder gPercorsoEsp
        End If
        
        txtStream.Close
        Set txtStream = Nothing
        Dim nomefile As String
        nomefile = gPercorsoEsp & "\" & gAccessiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gAccessiFileType = 0, ".TXT", ".CSV")
        fso.MoveFile fileEXP, nomefile
        
        Select Case gTipoInvioAccessi
         Case ProtocolloInvioEnum.Protocollo_Invio_FTP
            'invio dati via FTP
            Dim FTP As frmFTP
            Set frmFTP = New frmFTP
         
           
            frmFTP.FileDestination = "out/" & fso.GetFileName(nomefile)
            frmFTP.FileSource = nomefile
            frmFTP.Rename = True
            frmFTP.DoWork 1
            
            On Error Resume Next
            Set frmHTTP = Nothing
        End Select
    End If
    
    gDBC.CommitTrans
    Set fso = Nothing
    
    If Me.Visible Then CaricaMovimenti
    
    Exit Sub
EsportazioneERR:
    On Error Resume Next
    gDBC.RollbackTrans
End Sub

Private Sub CaricaMovimenti()
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim codFiscIva As String
    
    Screen.MousePointer = vbHourglass
    mWhere = vbNullString
    If (Not IsNull(DTPicker1(0).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "DataOraIngresso >= #" & Format$(DTPicker1(0).value, "YYYY-MM-DD 00:00:00") & "#"
    End If

    If (Not IsNull(DTPicker1(1).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "DataOraIngresso <= #" & Format$(DTPicker1(1).value, "YYYY-MM-DD 23:59:59") & "#"
    End If

    If (Check1.value = 1) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "Stato = '1'"
    Else
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "Stato = '0' or Stato = '3'"
    End If
    
    If mWhere <> "" Then mWhere = mWhere & " AND "
    mWhere = mWhere & " rilAcc.CodComune ='" & GetCurrentComune!codistat & "'"


    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    rs.Source = "SELECT " & IIf(pOpzioneAvvio = 2, "TOP 10", "") & " RILACC.id,RILACC.ospite,RILACC.stato,RILACC.codBadge as codbadge1,Badge.CodBadge," & _
                "Badge.codComune,Utenti.Nome,Utenti.Cognome,Utenti.RagioneSociale, RILACC.DataOraIngresso, Utenti.CodFiscale, Utenti.PartitaIva " & _
                "FROM  (RilevazioneAccessi as rilAcc LEFT JOIN Badge ON (RILACC.CodBadge = Badge.CodBadge) AND (RILACC.CodComune = Badge.CodComune)) " & _
                "LEFT JOIN Utenti ON (Badge.CodComune = Utenti.CodComune) AND (Badge.CodUtente = Utenti.CodUtente) " & _
                "where " & mWhere
 
    
    
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    
    Dim ragioneSociale As String
    
    mChecked = 0
    lvwMovEsp.ListItems.Clear
    Do While Not rs.EOF
    
        ragioneSociale = IIf(nullToBlank(rs!Nome) = vbNullString, nullToBlank(rs!ragioneSociale), nullToBlank(rs!Nome) & " " & nullToBlank(rs!Cognome))
        If Trim(ragioneSociale) = vbNullString Then ragioneSociale = "Ospite"
        Set item = lvwMovEsp.ListItems.Add(, "_" & rs!id, IIf(nullToBlank(rs!codBadge) = vbNullString, rs!codbadge1, nullToBlank(rs!codBadge)))
        item.ListSubItems.Add , "Nominativo", ragioneSociale
        item.ListSubItems.Add , "DataOraIngresso", Format$(rs!DataOraIngresso, "dd-MMM-yyyy hh:mm:ss")
        
        codFiscIva = IIf(nullToBlank(rs!CodFiscale) = vbNullString, nullToBlank(rs!PartitaIva), nullToBlank(rs!CodFiscale))
        
        If (rs!ospite And codFiscIva = vbNullString) Then codFiscIva = rs!codbadge1
        
        item.ListSubItems.Add , "CodFiscaleIva", codFiscIva
        item.ListSubItems.Add , "Ospite", IIf(rs!ospite, "Si", "No")
        item.ListSubItems.Add , "Id", rs!id
        
       
        Select Case rs!Stato
            Case "0"
                item.ListSubItems.Add , "Stato", "Uscito"
            Case "1"
                item.ListSubItems.Add , "Stato", "Esportato"
            Case "2"
                item.ListSubItems.Add , "Stato", "Entrato"
            Case "3"
                item.ListSubItems.Add , "Stato", "Rifiutato"
        End Select
        
  
        item.Checked = True
        rs.MoveNext
        
        mChecked = mChecked + 1
    Loop
    
    AllargaLeColonneLVW lvwMovEsp, False
    
    Screen.MousePointer = vbDefault
    If rs.state = adStateOpen Then
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
