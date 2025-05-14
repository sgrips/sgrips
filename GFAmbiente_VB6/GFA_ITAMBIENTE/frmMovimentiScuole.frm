VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMovimentiScuole 
   Caption         =   "Movimenti Scuole"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7335
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   225
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Index           =   2
      Left            =   3720
      Picture         =   "frmMovimentiScuole.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Annulla senza salvare"
      Top             =   5160
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Esporta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Index           =   1
      Left            =   4950
      Picture         =   "frmMovimentiScuole.frx":4282
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Annulla senza salvare"
      Top             =   5160
      Width           =   1065
   End
   Begin MSComctlLib.ListView lvwMovimenti 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Index           =   0
      Left            =   6120
      Picture         =   "frmMovimentiScuole.frx":8504
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Annulla senza salvare"
      Top             =   5160
      Width           =   1065
   End
End
Attribute VB_Name = "frmMovimentiScuole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsMovimenti As Recordset
Private m_cHdrIcons As New cLVHeaderSortIcons

Private Sub CaricaMovimenti()
    Dim sql As String
    Dim cmd As ADODB.Command
    
    Set rsMovimenti = CreateObject("ADODB.Recordset")
    
    sql = "SELECT CodiciCerComunali.Descrizione AS descodcer, Movimenti.*, Utenti.CodUtente, Utenti.TipoAnagrafe, Utenti.RagioneSociale, Utenti_1.CodUtente, Utenti_1.TipoAnagrafe, Utenti_1.RagioneSociale, Utenti_2.CodUtente, Utenti_2.TipoAnagrafe, Utenti_2.RagioneSociale, Utenti_3.CodUtente, Utenti_3.TipoAnagrafe, Utenti_3.RagioneSociale " & _
          "FROM ((((Movimenti INNER JOIN Utenti ON Movimenti.CodUtente = Utenti.CodUtente) " & _
          "LEFT JOIN Utenti AS Utenti_1 ON Utenti.CodPadre = Utenti_1.CodUtente) " & _
          "LEFT JOIN Utenti AS Utenti_2 ON Utenti_1.CodPadre = Utenti_2.CodUtente) " & _
          "LEFT JOIN Utenti AS Utenti_3 ON Utenti_2.CodPadre = Utenti_3.CodUtente) " & _
          "INNER JOIN CodiciCerComunali ON (Movimenti.CodComune = CodiciCerComunali.CodComune) AND (Movimenti.CodSubCer = CodiciCerComunali.CodSubCer) AND (Movimenti.CodCer = CodiciCerComunali.CodCer) AND (Movimenti.CodSottoCategoria = CodiciCerComunali.CodSottoCategoria) AND (Movimenti.CodCategoria = CodiciCerComunali.CodCategoria)" & _
          "WHERE Utenti.TipoAnagrafe=? AND Utenti_1.TipoAnagrafe=? AND (Utenti_2.TipoAnagrafe=? or Utenti_2.TipoAnagrafe is null) AND (Utenti_3.TipoAnagrafe=? or Utenti_3.TipoAnagrafe is null)"

    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = sql
    
    Set cmd.ActiveConnection = gDBC
    cmd.Parameters.Append cmd.CreateParameter("P1", adVarChar, adParamInput, 1)
    cmd.Parameters.Append cmd.CreateParameter("P2", adVarChar, adParamInput, 1)
    cmd.Parameters.Append cmd.CreateParameter("P3", adVarChar, adParamInput, 1)
    cmd.Parameters.Append cmd.CreateParameter("P4", adVarChar, adParamInput, 1)
    
    
    Dim i As Long
    For i = 4 To 1 Step -1
        If i <= gPuntiAssegnatiPer Then
            cmd("P" & (gPuntiAssegnatiPer - i) + 1).Value = Choose(i, "I", "S", "C", "Z")
        Else
            cmd("P" & i).Value = "@"
        End If
    Next
    
    Set rsMovimenti = cmd.Execute
    Set cmd.ActiveConnection = Nothing

End Sub


Private Sub ApplyFilter(Optional filtro As Variant)
    Dim where As String
    Dim i As Long
    
    rsMovimenti.Filter = ""
    
    If Not IsMissing(filtro) Then
        For i = 1 To gPuntiAssegnatiPer
            If filtro(gPuntiAssegnatiPer - i) <> -1 Then
                If where <> "" Then where = where & " and "
                where = where & Choose(i, "Utenti.CodUtente", "Utenti_1.CodUtente", "Utenti_2.CodUtente", "Utenti_3.CodUtente") & " = " & _
                        filtro(gPuntiAssegnatiPer - i)
            End If
        Next
        
        rsMovimenti.Filter = where
    End If
    
    CreaListView
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim ret() As Long
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Esporta
        Case 2
            ApplyFilter frmFiltroMovimentiScuole.ShowMe
    End Select
End Sub

Private Sub Esporta()
    Dim item As ListItem
    Dim ItemH As ColumnHeader
    Dim ItemS As ListSubItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim mChecked As Long
    
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gMovimentiFile & IIf(gMovimentiFileType = 0, "_S.TXT", "_S.CSV")
   
    If lvwMovimenti.ListItems.Count = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = lvwMovimenti.ListItems.Count
    ProgressBar1.Value = 0
    
    For Each ItemH In lvwMovimenti.ColumnHeaders
        txtStream.Write ItemH.Text
        If ItemH.Index < lvwMovimenti.ColumnHeaders.Count Then
            txtStream.Write gMovimentiSeparatore
        End If
    Next
    txtStream.Write vbCrLf
    
    For Each item In lvwMovimenti.ListItems
        ProgressBar1.Value = ProgressBar1.Value + 1
        txtStream.Write Chr$(34) & item.Text & Chr$(34)
        txtStream.Write gMovimentiSeparatore
        For Each ItemS In item.ListSubItems
            
            txtStream.Write Chr$(34) & IIf(ItemS.Tag <> "", ItemS.Tag, ItemS.Text) & Chr$(34)
            If ItemS.Index < item.ListSubItems.Count Then
                txtStream.Write gMovimentiSeparatore
            End If

        Next
        txtStream.Write vbCrLf
    Next
    
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    
    fso.MoveFile fileEXP, gPercorsoEsp & "\" & gMovimentiFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gMovimentiFileType = 0, ".TXT", ".CSV")
    
    
    Set fso = Nothing
    
    
    
    CaricaMovimenti
    
    ProgressBar1.Visible = False
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title

End Sub


Private Sub Form_Load()
    CaricaMovimenti
    ApplyFilter
End Sub


Private Sub CreaListView()
    Dim i As Long
    Dim item As ListItem
    lvwMovimenti.ListItems.Clear
    
    
    If rsMovimenti.EOF Then Exit Sub
    
    If lvwMovimenti.ColumnHeaders.Count = 0 Then
        
        For i = 1 To gPuntiAssegnatiPer
            lvwMovimenti.ColumnHeaders.Add , , Choose(i, "Istituto", "Scuola", "Classe", "Sezione")
        Next
        lvwMovimenti.ColumnHeaders.Add , , "Cod.Cer"
        lvwMovimenti.ColumnHeaders.Add , , "Tipo Rif."
        lvwMovimenti.ColumnHeaders.Add , , "Nr. Mov."
        lvwMovimenti.ColumnHeaders.Add , , "Data Mov."
        lvwMovimenti.ColumnHeaders.Add , , "Quantità"
        lvwMovimenti.ColumnHeaders.Add , , "U.M."
        lvwMovimenti.ColumnHeaders.Add , , "Punti"
        lvwMovimenti.ColumnHeaders.Add , , "Causale"
        lvwMovimenti.ColumnHeaders.Add , , "Id.Ali.Mem"
        lvwMovimenti.ColumnHeaders.Add , , "Utente Ins."
        lvwMovimenti.ColumnHeaders.Add , , "Data Ins."
    End If
    

    Do While Not rsMovimenti.EOF
        Set item = lvwMovimenti.ListItems.Add(, "_" & rsMovimenti!nrMovimento)
        
        For i = gPuntiAssegnatiPer To 1 Step -1
           If i = gPuntiAssegnatiPer Then
                item.Text = rsMovimenti(Choose(i, "Utenti.RagioneSociale", "Utenti_1.RagioneSociale", "Utenti_2.RagioneSociale", "Utenti_3.RagioneSociale"))
            Else
                item.ListSubItems.Add , , rsMovimenti(Choose(i, "Utenti.RagioneSociale", "Utenti_1.RagioneSociale", "Utenti_2.RagioneSociale", "Utenti_3.RagioneSociale"))
            End If
        Next
        item.ListSubItems.Add , , rsMovimenti!codCategoria & rsMovimenti!codSottoCategoria & rsMovimenti!codCer
        item.ListSubItems.Add , , rsMovimenti!DesCodCer
        item.ListSubItems.Add , , rsMovimenti!nrMovimento
        item.ListSubItems.Add , , rsMovimenti!DataMovimento
        item.ListSubItems.Add , , rsMovimenti!quantita
        item.ListSubItems.Add , , rsMovimenti!UnitaMisura
        item.ListSubItems.Add , , rsMovimenti!PuntiDovuti
        item.ListSubItems.Add , , rsMovimenti!codCausale
        item.ListSubItems.Add , , nullToBlank(rsMovimenti!idAlibiMemory)
        item.ListSubItems.Add , , rsMovimenti!codUser
        item.ListSubItems.Add , , rsMovimenti!dataIns
        
        rsMovimenti.MoveNext
    Loop
    
    AllargaLeColonneLVW lvwMovimenti, False
End Sub



Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - 100
    posLeft = Me.ScaleWidth - cmd(i).width - 100
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - 100 - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    lvwMovimenti.Move 0, 0, Me.ScaleWidth - 100, cmd(0).top - 100
    
    
    ProgressBar1.Move 0, Me.ScaleHeight - ProgressBar1.Height - 100, cmd(cmd.Count - 1).left - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lStyle As Long
    
    If gStartMode = 0 Or gStartMode = 1 Then
        lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
        lStyle = lStyle Or MY_WS_DOCK
        SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    
        If gStartMode = 1 Then
            SetParent Me.hWnd, hOldParent
        End If
    End If

    Set pfrmMovimentiScuole = Nothing
End Sub


Public Sub ShowMe()
    
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        Me.Show vbModal
    Else
        Me.Show
    End If
End Sub

Private Sub lvwMovimenti_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListview lvwMovimenti, ColumnHeader.Index, lvsdToggle, True, True, m_cHdrIcons
End Sub
