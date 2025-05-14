VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRicercaUtente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ricerca Utente"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12150
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   240
      MaxLength       =   16
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Top             =   480
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerca"
      Default         =   -1  'True
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
      Left            =   8280
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleziona"
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
      Left            =   9600
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Index           =   0
      Left            =   10920
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      MaxLength       =   16
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codice Utente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codice Badge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ragione Sociale / Cognome Nome- utilizzare (*) come carattere jolly "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Width           =   5970
   End
   Begin VB.Label Label1 
      Caption         =   "Partita Iva"
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
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codice Fiscale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   1290
   End
End
Attribute VB_Name = "frmRicercaUtente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cHdrIcons As New cLVHeaderSortIcons


Private Const CMD_CHIUDI = 0
Private Const CMD_SELEZIONA = 1
Private Const CMD_RICERCA = 2

Private rs As ADODB.Recordset

Private Const TXT_CODICEFISCALE = 0
Private Const TXT_PARTITAIVA = 1
Private Const TXT_RAGIONESOCIALE = 3
Private Const TXT_BADGE = 5
Private Const TXT_CODUTENTE = 6

Private txtReturn As TextBox
Private pTipoUtente As String
Private pCodiceFiscaleIva As Boolean

Private RS_UTENTE   As String
Private RS_TRASPORTATORI As String
Private RS_DESTINATARI As String
Private RS_INTERMEDIARI As String


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case CMD_CHIUDI
            Unload Me
        Case CMD_RICERCA
            ricerca
        Case CMD_SELEZIONA
                    
            seleziona
            Unload Me
    End Select
End Sub


Private Sub ricerca()
    Dim item As ListItem
    Dim search As String
    Dim value As String
    
    
    On Error GoTo ricercaERR:
    
    Screen.MousePointer = vbHourglass
    
    rs.Filter = vbNullString
    search = vbNullString
    ListView1.ListItems.Clear
    
    
    If txt(TXT_CODICEFISCALE) <> vbNullString And txt(TXT_CODICEFISCALE) <> "*" Then
        search = " CodFiscale like '" & txt(TXT_CODICEFISCALE) & "%'"
    End If
    
    
    
    
    If txt(TXT_BADGE) <> vbNullString And txt(TXT_BADGE) <> "*" Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " CodBadge like '" & txt(TXT_BADGE) & "%'"
    End If
    
    If txt(TXT_CODUTENTE) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " CodUtente = " & txt(TXT_CODUTENTE)
    End If
    
    
    
    If txt(TXT_PARTITAIVA) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " PartitaIva >= " & left$(txt(TXT_PARTITAIVA) & "00000000000", 11) & " and PartitaIva <" & left$((txt(TXT_PARTITAIVA).Text + 1) & "00000000000", 11)
    End If
    
    If txt(TXT_RAGIONESOCIALE) <> vbNullString And txt(TXT_RAGIONESOCIALE) <> "*" Then
        If search <> vbNullString Then search = search & " AND "
        value = Replace(txt(TXT_RAGIONESOCIALE).Text, "*", "%")
        If InStr(1, value, "%") <= 0 Then value = value & "%"
        
        search = search & " (RagioneSociale like '" & value & "' OR CognomeNome like '" & value & "')"
    End If
    
    
    
    If pTipoUtente = "U" Then
        rs.Filter = IIf(search = "", "", search)
    Else
        rs.Filter = search
    End If
    
    
    Dim ragioneSociale As String
    
    
    Do While Not rs.EOF
            
         ragioneSociale = IIf(nullToBlank(rs!Cognome) = "", nullToBlank(rs!ragioneSociale), nullToBlank(rs!Cognome) & " " & nullToBlank(rs!Nome))



        Select Case pTipoUtente
            Case "U"
                Set item = ListView1.ListItems.Add(, "_" & rs!codUtente, rs!codUtente)
                item.ListSubItems.Add , rs!ragioneSociale.Name, ragioneSociale
                item.ListSubItems.Add , rs!codBadge.Name, nullToBlank(rs!codBadge.value)
                item.ListSubItems.Add , rs!Sesso.Name, nullToBlank(rs!Sesso.value)
            Case "T"
                Set item = ListView1.ListItems.Add(, "_" & rs!idTrasportatore, rs!idTrasportatore)
                item.ListSubItems.Add , rs!ragioneSociale.Name, ragioneSociale
            Case "D"
                Set item = ListView1.ListItems.Add(, "_" & rs!idDestinatario, rs!idDestinatario)
                item.ListSubItems.Add , rs!ragioneSociale.Name, ragioneSociale
            Case "I"
                Set item = ListView1.ListItems.Add(, "_" & rs!idIntermediario, rs!idIntermediario)
                item.ListSubItems.Add , rs!ragioneSociale.Name, ragioneSociale
        End Select
        
        item.ListSubItems.Add , rs!CodFiscale.Name, nullToBlank(rs!CodFiscale.value)
        item.ListSubItems.Add , rs!PartitaIva.Name, nullToBlank(rs!PartitaIva.value)
        item.ListSubItems.Add , rs!DatadiNascita.Name, nullToBlank(rs!DatadiNascita.value)
        item.ListSubItems.Add , rs!Comune.Name, nullToBlank(rs!Comune.value)
        
        Select Case pTipoUtente
            Case "U"
                item.ListSubItems.Add , rs!IndirizzoResidenza.Name, nullToBlank(rs!IndirizzoResidenza.value)
            Case "T"
                item.ListSubItems.Add , rs!IndirizzoSedeLegale.Name, nullToBlank(rs!IndirizzoSedeLegale.value)
            Case "D"
                item.ListSubItems.Add , rs!Indirizzo.Name, nullToBlank(rs!Indirizzo.value)
                
        End Select
        
        rs.MoveNext
    Loop
    
    AllargaLeColonneLVW ListView1, False
    
    Screen.MousePointer = vbNormal
    Exit Sub
ricercaERR:
    MsgBox "Ricerca:" & Error, vbExclamation, App.Title
End Sub

Private Sub seleziona()
    On Error Resume Next
    If Not ListView1.SelectedItem Is Nothing Then
        If pCodiceFiscaleIva Then
            txtReturn.Text = IIf(ListView1.SelectedItem.ListSubItems("CodFiscale") = "", ListView1.SelectedItem.ListSubItems("PartitaIva"), ListView1.SelectedItem.ListSubItems("CodFiscale"))
        Else
            txtReturn.Text = ListView1.SelectedItem.Text
        End If
    End If
    
End Sub


Private Sub Form_Load()
    
    RS_UTENTE = "select utenti.*,utenti.cognome + ' ' + utenti.nome as CognomeNome,codBadge,databadgeinizio,databadgefine,progressivobadge,comuni.comune from ((Utenti LEFT JOIN Comuni ON Utenti.CodComuneResidenza = Comuni.CODISTAT) Left join badge on badge.codutente = utenti.codutente AND badge.codComune = utenti.codComune) where ([default]=true or [default] is null) and utenti.CodComune='" & GetCurrentComune("CodIstat").value & "'"
    RS_TRASPORTATORI = "select trasportatori.*,'' as CognomeNome,comuni.comune from (trasportatori LEFT JOIN Comuni ON trasportatori.CodComuneSedeLegale = Comuni.CODISTAT) where Comuni.CODISTAT='" & GetCurrentComune("CodIstat").value & "'"
    RS_DESTINATARI = "select destinatari.*,'' as CognomeNome,comuni.comune from (destinatari LEFT JOIN Comuni ON destinatari.CodComune = Comuni.CODISTAT) where Comuni.CODISTAT='" & GetCurrentComune("CodIstat").value & "'"
    RS_INTERMEDIARI = "select Intermediari.*,'' as CognomeNome,comuni.comune from (Intermediari LEFT JOIN Comuni ON Intermediari.CodComune = Comuni.CODISTAT) where Comuni.CODISTAT='" & GetCurrentComune("CodIstat").value & "'"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    
    Select Case pTipoUtente
        Case "U"
            rs.Source = RS_UTENTE
        Case "T"
            rs.Source = RS_TRASPORTATORI
        Case "D"
            rs.Source = RS_DESTINATARI
        Case "I"
            rs.Source = RS_INTERMEDIARI
            
    End Select
    
    
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    ListView1.ColumnHeaders.Add , , "Codice"
    ListView1.ColumnHeaders.Add , , "Ragione Sociale"
    ListView1.ColumnHeaders.Add , , "Cod.Badge"
    ListView1.ColumnHeaders.Add , , "Sesso"
    ListView1.ColumnHeaders.Add , , "Cod.Fiscale"
    ListView1.ColumnHeaders.Add , , "Partita IVA"
    ListView1.ColumnHeaders.Add , , "Data Nascita"
    ListView1.ColumnHeaders.Add , , "Comune"
    ListView1.ColumnHeaders.Add , , "Indirizzo"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub


Public Sub schowMe(txt As TextBox, TipoUtente As String, Optional codiceFiscaleIva As Boolean = False)
    Set txtReturn = txt
    pTipoUtente = TipoUtente
    pCodiceFiscaleIva = codiceFiscaleIva
    Set m_cHdrIcons.ListView = ListView1
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 3
    Call m_cHdrIcons.SetHeaderIcons(ListView1.SortKey, ListView1.SortOrder)
    
    
    Me.Show vbModal
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview ListView1, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub

Private Sub ListView1_DblClick()
    Command1_Click (CMD_SELEZIONA)
End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case TXT_PARTITAIVA
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    End Select
End Sub
