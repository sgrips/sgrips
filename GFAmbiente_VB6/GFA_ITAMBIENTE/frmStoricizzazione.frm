VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStoricizzazione 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Storicizzazione"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Storicizza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   308
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
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Avanzamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Indicare la data da Storicizzare"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmStoricizzazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Command1(0).enabled = False
            Command1(1).enabled = False
            Storicizza
            Command1(0).enabled = True
            Command1(1).enabled = True
    
    End Select
End Sub

Private Sub Form_Load()
    
    lblInfo.Caption = ""
    
    caricaDati
End Sub

Private Sub caricaDati()
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("select * from Storico where Storicizzato = false")
    If Not rs.EOF Then
        Combo1.AddItem rs!Anno
        Combo1.ItemData(Combo1.ListCount - 1) = rs!Anno
        rs.Close
    End If
End Sub


Private Sub Storicizza()
    Dim TotaleDati As Long
    Dim rs As ADODB.Recordset
    Dim Anno As Integer
    Dim con As Connection
    Dim f As FileSystemObject
    Dim BeginTrans As Boolean
    
    If Forms.Count > 4 Then
        MsgBox "Prima di utilizzare questa funzione occorre uscire da tutte le videate aperte", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    Set f = New FileSystemObject
    
    On Error GoTo StoricizzaERR
    If Combo1.Text = vbNullString Then
        MsgBox "Selezionare un anno per la storicizzazione"
        Exit Sub
    End If
    
    lblInfo.Caption = "Inizializzazione"
    Anno = Combo1.Text
    Set rs = gDBC.Execute("Select count(*) from FormularioRifiuti where Data Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from LegameLottoIngressoLottoUscita where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from LegameLottoRegistrazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from LegameRegistroOperazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from Lotti where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from Movimenti where DataMovimento Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from Registro where dataOperazione Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from Fatturazione where DATA_MOVIMENTO Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    Set rs = gDBC.Execute("Select count(*) from RilevazioneAccessi where DataOraIngresso Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    TotaleDati = TotaleDati + rs(0).Value
    
    If TotaleDati = 0 Then Exit Sub
    
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    ProgressBar1.Max = TotaleDati
    
    lblInfo.Caption = "Inizializzazione"
    
  
    
    lblInfo.Caption = "Genarazione Archivio"
    lblInfo.Refresh
    f.CopyFile f.BuildPath(App.Path, "Storico.mdb"), f.BuildPath(gPercorsoStoricizzazione, "Storico" & Anno & ".mdb")
    
    Set con = CreateObject("ADODB.Connection")
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & f.BuildPath(gPercorsoStoricizzazione, "Storico" & Anno & ".mdb") & ";Jet OLEDB:Database Password=" + gPasswordDB
    con.Open
    
    lblInfo.Caption = "Copia Formulario Rifiuti"
    lblInfo.Refresh
    If Not CopiaFormularioRifiuti(con, Anno) Then
        Exit Sub
    End If
        
        
    lblInfo.Caption = "Copia Legame Lotto Ingresso Lotto Uscita"
    lblInfo.Refresh
    If Not CopiaLegameLottoIngressoLottoUscita(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Legame Lotto Registrazione Carico"
    lblInfo.Refresh
    If Not CopiaLegameLottoRegistrazioneCarico(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Legame Registro Operazione Carico"
    lblInfo.Refresh
    If Not CopiaLegameRegistroOperazioneCarico(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Lotti"
    lblInfo.Refresh
    If Not CopiaLotti(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Movimenti"
    lblInfo.Refresh
    If Not CopiaMovimenti(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Registro"
    lblInfo.Refresh
    If Not CopiaRegistro(con, Anno) Then
        Exit Sub
    End If
    
    lblInfo.Caption = "Copia Fatturazione"
    lblInfo.Refresh
    If Not CopiaFatturazione(con, Anno) Then
        Exit Sub
    End If
                
    
    lblInfo.Caption = "Copia Rilevazione Accessi"
    lblInfo.Refresh
    If Not CopiaRilevazioneAccessi(con, Anno) Then
        Exit Sub
    End If
    'completato
    MsgBox "Operazione Avvenuta con Successo, " & vbCrLf & "ora si procederà all'archiviazione dell'anno scelto " & Anno & " e" & vbCrLf & "i contatori saranno impostati al valore iniziale", vbExclamation + vbOKOnly, App.Title
                
    gDBC.BeginTrans
    BeginTrans = True
    gDBC.Execute ("delete * from FormularioRifiuti where Data Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from LegameLottoIngressoLottoUscita where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from LegameLottoRegistrazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from LegameRegistroOperazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from Lotti where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from Movimenti where DataMovimento Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from Registro where dataOperazione Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from Fatturazione where DATA_MOVIMENTO Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from RilevazioneAccessi where DataOraIngresso Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#")
    gDBC.Execute ("delete * from Contatori where Entita<>'IMMOBILI'")
                        
                        
    gDBC.Execute ("update Storico set Storicizzato=true,Percorso='" & f.BuildPath(gPercorsoStoricizzazione, "Storico" & Anno & ".mdb") & "' where anno=" & Anno)
    gDBC.CommitTrans
                        
                        
    con.Execute ("insert into AnnoStoricizzato (Anno) values(" & Anno & ")")
    con.Close
    Set con = Nothing
    
    ComprimiDB
                
    Set f = Nothing
    Exit Sub
StoricizzaERR:
    If BeginTrans Then gDBC.RollbackTrans
    MsgBox "frmStoricizzazione Storicizza:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Sub


Private Function CopiaRilevazioneAccessi(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaRilevazioneAccessiERR
    
    CopiaRilevazioneAccessi = False
    

    
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "INSERT INTO RilevazioneAccessi (id, CodComune, CodBadge, DataOraIngresso, DataOraUscita, Stato, Ospite, UtenzaDomestica, DaPesare )" & _
            "values (?,?,?,?, ?, ?, ?, ?, ?)"
            
    cmdq.Parameters.Append cmdq.CreateParameter("ID", adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("CODCOMUNE", adVarChar, adParamInput, 6)
    cmdq.Parameters.Append cmdq.CreateParameter("CodBadge", adVarChar, adParamInput, 20)
    cmdq.Parameters.Append cmdq.CreateParameter("DataOraIngresso", adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("DataOraUscita", adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("STATO", adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter("Ospite", adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("UtenzaDomestica", adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("DaPesare", adBoolean, adParamInput)
    
    cmdq.Prepared = True
    Set cmdq.ActiveConnection = con
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    
    rs.Source = "Select * from RilevazioneAccessi where DataOraIngresso Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0).Value = rs!id.Value
        cmdq.Parameters(1).Value = rs!codComune.Value
        cmdq.Parameters(2).Value = rs!codBadge.Value
        cmdq.Parameters(3).Value = rs!DataOraIngresso.Value
        cmdq.Parameters(4).Value = rs!DataOraUscita.Value
        cmdq.Parameters(5).Value = rs!Stato.Value
        cmdq.Parameters(6).Value = rs!ospite.Value
        cmdq.Parameters(7).Value = rs!UtenzaDomestica.Value
        cmdq.Parameters(8).Value = rs!DaPesare.Value

        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    
    
    
    CopiaRilevazioneAccessi = True
    Exit Function
CopiaRilevazioneAccessiERR:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaFatturazione(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaFatturazioneERR
    
    CopiaFatturazione = False
    

    
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "INSERT INTO Fatturazione (id,CODCOMUNE,PROGRESSIVO_INVIO, COD_ANA,  " & _
            "COD_IMMOBILE, COD_RIFIUTO, QUANTITA, DATA_MOVIMENTO, NUMERO_FORMULARIO, STATO_INVIO, DATA_SCARICO, FLAG, [NOTE],STATO,IDRIFMOVIMENTO) " & _
            "values (?,?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)"
            
    cmdq.Parameters.Append cmdq.CreateParameter("ID", adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("CODCOMUNE", adVarChar, adParamInput, 6)
    cmdq.Parameters.Append cmdq.CreateParameter("PROGRESSIVO_INVIO", adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("COD_ANA", adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("COD_IMMOBILE", adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("COD_RIFIUTO", adVarChar, adParamInput, 8)
    
    
    Set P = cmdq.CreateParameter("QUANTITA", adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6
    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter("DATA_MOVIMENTO", adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("NUMERO_FORMULARIO", adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter("STATO_INVIO", adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter("DATA_SCARICO", adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter("FLAG", adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter("NOTE", adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter("STATO", adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter("IDRIFMOVIMENTO", adInteger, adParamInput)
    
    cmdq.Prepared = True
    Set cmdq.ActiveConnection = con
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    
    rs.Source = "Select * from Fatturazione where DATA_MOVIMENTO Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0).Value = rs!id.Value
        cmdq.Parameters(1).Value = rs!codComune.Value
        cmdq.Parameters(2).Value = rs!PROGRESSIVO_INVIO.Value
        cmdq.Parameters(3).Value = rs!COD_ANA.Value
        cmdq.Parameters(4).Value = rs!COD_IMMOBILE.Value
        cmdq.Parameters(5).Value = rs!COD_RIFIUTO.Value
        cmdq.Parameters(6).Value = rs!quantita.Value
        cmdq.Parameters(7).Value = rs!DATA_MOVIMENTO.Value
        cmdq.Parameters(8).Value = rs!NUMERO_FORMULARIO.Value
        cmdq.Parameters(9).Value = rs!STATO_INVIO.Value
        cmdq.Parameters(10).Value = rs!DATA_SCARICO.Value
        cmdq.Parameters(11).Value = rs!Flag.Value
        cmdq.Parameters(12).Value = rs!note.Value
        cmdq.Parameters(13).Value = rs!Stato.Value
        cmdq.Parameters(14).Value = rs!idRifMovimento.Value

        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    
    
    
    CopiaFatturazione = True
    Exit Function
CopiaFatturazioneERR:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
End Function


Private Function CopiaRegistro(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaRegistroERR
    
    Set cmdq = CreateObject("ADODB.Command")
    
    cmdq.CommandText = "INSERT INTO Registro (idRegistro, nrRegistro, flgOperazione, dataOperazione, nrFormulario, dataFormulario, " & _
             " codCategoria, codSottocategoria, codCer, codTipoOpeSmaltRecupero, unitaMisura, quantita, ImpiantoiProduzione, idIntermediario, [note], flgStampato,DataIns,CodUser,classiPericolosita )" & _
             " values(?,?, ?, ?, ?, ?, ?, ?, ?,?,?, ?, ?, ?, ?, ?,?,?,?)"
             

    Set cmdq.ActiveConnection = gDBC

    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    
    cmdq.Prepared = True
    Set cmdq.ActiveConnection = con
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    
    rs.Source = "Select * from Registro where dataOperazione Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0).Value = rs!idRegistro.Value
        cmdq.Parameters(1).Value = rs!nrRegistro.Value
        cmdq.Parameters(2).Value = rs!flgOperazione.Value
        cmdq.Parameters(3).Value = rs!dataOperazione.Value
        cmdq.Parameters(4).Value = rs!nrFormulario.Value
        cmdq.Parameters(5).Value = rs!DataFormulario.Value
        cmdq.Parameters(6).Value = rs!codCategoria.Value
        cmdq.Parameters(7).Value = rs!codSottoCategoria.Value
        cmdq.Parameters(8).Value = rs!codCer.Value
        cmdq.Parameters(9).Value = rs!codTipoOpeSmaltRecupero.Value
        cmdq.Parameters(10).Value = rs!UnitaMisura.Value
        cmdq.Parameters(11).Value = rs!quantita.Value
        cmdq.Parameters(12).Value = rs!ImpiantoiProduzione.Value
        cmdq.Parameters(13).Value = rs!idIntermediario.Value
        cmdq.Parameters(14).Value = rs!note.Value
        cmdq.Parameters(15).Value = rs!flgStampato.Value
        cmdq.Parameters(16).Value = rs!dataIns.Value
        cmdq.Parameters(17).Value = rs!codUser.Value
        cmdq.Parameters(18).Value = rs!classiPericolosita.Value
        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    
    
    
    CopiaRegistro = True
    Exit Function
CopiaRegistroERR:
    MsgBox "CopiaRegistro:" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function


Private Function CopiaMovimenti(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaMovimentiERR
    
    CopiaMovimenti = False
    
    Set cmdq = CreateObject("ADODB.Command")
    'movimenti
    cmdq.CommandText = "Insert into Movimenti (id,CodImmobile,CodUtente,CodCategoria,CodSottoCategoria, CodCer, DataMovimento, Quantita, Stato,Manuale, " & _
        "CodUser,BonusDovuto,UnitaMisura,Targa,NumeroFormulario,DataFormulario,flgForzatura, " & _
        "TargaTrasportatore,CodTrasportatore,nrMovimento,codCausale,flgRegistrato,CodPesa)" & _
        " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adWChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    
    cmdq.Prepared = True
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
        
    rs.Source = "Select * from Movimenti where DataMovimento Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq(0).Value = rs!id.Value
        cmdq(1).Value = rs!codImmobile.Value
        cmdq(2).Value = rs!codUtente.Value
        cmdq(3).Value = rs!codCategoria.Value
        cmdq(4).Value = rs!codSottoCategoria.Value
        cmdq(5).Value = rs!codCer.Value
        cmdq(6).Value = rs!DataMovimento.Value
        cmdq(7).Value = rs!quantita.Value
        cmdq(8).Value = rs!Stato.Value
        cmdq(9).Value = rs!manuale.Value
        cmdq(10).Value = rs!codUser.Value
        cmdq(11).Value = rs!BonusDovuto.Value
        cmdq(12).Value = rs!UnitaMisura.Value
        cmdq(13).Value = rs!Targa.Value
        cmdq(14).Value = rs!NumeroFormulario.Value
        cmdq(15).Value = rs!DataFormulario.Value
        cmdq(16).Value = rs!flgForzatura.Value
        cmdq(17).Value = rs!TargaTrasportatore.Value
        cmdq(18).Value = rs!CodTrasportatore.Value
        cmdq(19).Value = rs!nrMovimento.Value
        cmdq(20).Value = rs!codCausale.Value
        cmdq(21).Value = rs!flgRegistrato.Value
        cmdq(22).Value = rs!codPesa.Value

        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaMovimenti = True
    
    Exit Function
CopiaMovimentiERR:
    MsgBox "CopiaMovimenti" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaLotti(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaLottiERR
    
    CopiaLotti = False
    
    'lotti
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandText = "INSERT INTO Lotti (id, CodLotto,  CodCategoria, CodSottoCategoria, CodCer, " & _
                       "QtaIngresso, QtaUscita,Tipo,Stato, DataMod, CodUserMod, DataIns, CodUserIns ) " & _
                       "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"

    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    cmdq.Prepared = True
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
        
    rs.Source = "Select * from Lotti where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq(0).Value = rs!id.Value
        cmdq(1).Value = rs!codLotto.Value
        cmdq(2).Value = rs!codCategoria.Value
        cmdq(3).Value = rs!codSottoCategoria.Value
        cmdq(4).Value = rs!codCer.Value
        cmdq(5).Value = rs!QtaIngresso.Value
        cmdq(6).Value = rs!QtaUscita.Value
        cmdq(7).Value = rs!tipo.Value
        cmdq(8).Value = rs!Stato.Value
        cmdq(9).Value = rs!DataMod.Value
        cmdq(10).Value = rs!CodUserMod.Value
        cmdq(11).Value = rs!dataIns.Value
        cmdq(12).Value = rs!codUserins.Value

        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaLotti = True
    
    Exit Function
CopiaLottiERR:
    MsgBox "CopiaLotti" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaLegameRegistroOperazioneCarico(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaLegameRegistroOperazioneCaricoERR
    
    CopiaLegameRegistroOperazioneCarico = False
    'formulario
    
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandText = "INSERT INTO LegameRegistroOperazioneCarico ( idRegistro, idoperazionecarico,dataIns)  values(?,?,? )"
    
    Set cmdq.ActiveConnection = gDBC
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    
    cmdq.Prepared = True
    
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
        
    rs.Source = "Select * from LegameRegistroOperazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0) = rs!idRegistro.Value
        cmdq.Parameters(1) = rs!idoperazionecarico.Value
        cmdq.Parameters(2) = rs!dataIns.Value
        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaLegameRegistroOperazioneCarico = True
    
    Exit Function
CopiaLegameRegistroOperazioneCaricoERR:
    MsgBox "LegameRegistroOperazioneCarico" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaLegameLottoRegistrazioneCarico(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaLegameLottoRegistrazioneCaricoERR
    
    CopiaLegameLottoRegistrazioneCarico = False
    'formulario
    
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "INSERT INTO LegameLottoRegistrazioneCarico ( idRegistro,  codLotto ,descrizione, DataIns, CodUserIns ) " & _
                       "VALUES (?,?,?,?,?)"
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    cmdq.Prepared = True
    
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
        
    rs.Source = "Select * from LegameLottoRegistrazioneCarico where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0) = rs!idRegistro.Value
        cmdq.Parameters(1) = rs!codLotto.Value
        cmdq.Parameters(2) = rs!Descrizione.Value
        cmdq.Parameters(3) = rs!dataIns.Value
        cmdq.Parameters(4) = rs!codUserins.Value
        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaLegameLottoRegistrazioneCarico = True
    
    Exit Function
CopiaLegameLottoRegistrazioneCaricoERR:
    MsgBox "CopiaLegameLottoRegistrazioneCarico" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaLegameLottoIngressoLottoUscita(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaLegameLottoIngressoLottoUscitaERR
    
    CopiaLegameLottoIngressoLottoUscita = False
    'formulario
    
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "insert into LegameLottoIngressoLottoUscita (id,LottoUscita,idLottoEntrata,DataIns,CodUser) values (?,?,?,?,?)"
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    cmdq.Prepared = True
    
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
        
    rs.Source = "Select * from LegameLottoIngressoLottoUscita where DataIns Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0) = rs!id.Value
        cmdq.Parameters(1) = rs!LottoUscita.Value
        cmdq.Parameters(2) = rs!idLottoEntrata.Value
        cmdq.Parameters(3) = rs!dataIns.Value
        cmdq.Parameters(4) = rs!codUser.Value
        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaLegameLottoIngressoLottoUscita = True
    
    Exit Function
CopiaLegameLottoIngressoLottoUscitaERR:
    MsgBox "CopiaLegameLottoIngressoLottoUscita" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Private Function CopiaFormularioRifiuti(con As ADODB.Connection, Anno As Integer) As Boolean
    Dim cmdq As Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo CopiaFormularioRifiutiERR
    
    CopiaFormularioRifiuti = False
    'formulario
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandType = adCmdText
    cmdq.CommandText = "INSERT INTO FormularioRifiuti (id,data,nrRegistro,  idMovimento, idImpProduttore, idImpDestinatario, idTrasportatore, idConducente, " & _
                "Annotazioni, codCategoria, codSottoCategoria, codCer, codOpeSmalRecup, carattChimicheFisiche, quantita, percorso, flgNormativaADR_RID, targa, dataInizioTrasporto,numColli, DataIns, CodUser ) " & _
                       "values (?,?,?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "

    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 35)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    cmdq.Prepared = True
    
    Set cmdq.ActiveConnection = con
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
        
    rs.Source = "Select * from FormularioRifiuti where Data Between #" & Anno & "-01-01 00:00:00# and #" & (Anno + 1) & "-12-31 23:59:59#"
    rs.Open

    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        cmdq.Parameters(0).Value = rs!id.Value
        cmdq.Parameters(1).Value = rs!Data.Value
        cmdq.Parameters(2).Value = rs!nrRegistro.Value
        cmdq.Parameters(3).Value = rs!idMovimento.Value
        cmdq.Parameters(4).Value = rs!idImpProduttore.Value
        cmdq.Parameters(5).Value = rs!idImpDestinatario.Value
        cmdq.Parameters(6).Value = rs!idTrasportatore.Value
        cmdq.Parameters(7).Value = rs!idConducente.Value
        cmdq.Parameters(8).Value = rs!Annotazioni.Value
        cmdq.Parameters(9).Value = rs!codCategoria.Value
        cmdq.Parameters(10).Value = rs!codSottoCategoria.Value
        cmdq.Parameters(11).Value = rs!codCer.Value
        cmdq.Parameters(12).Value = rs!codOpeSmalRecup.Value
        cmdq.Parameters(13).Value = rs!carattChimicheFisiche.Value
        cmdq.Parameters(14).Value = rs!quantita.Value
        cmdq.Parameters(15).Value = rs!percorso.Value
        cmdq.Parameters(16).Value = rs!flgNormativaADR_RID.Value
        cmdq.Parameters(17).Value = rs!Targa.Value
        cmdq.Parameters(18).Value = rs!dataInizioTrasporto.Value
        cmdq.Parameters(19).Value = rs!numColli.Value
        cmdq.Parameters(20).Value = rs!dataIns.Value
        cmdq.Parameters(21).Value = rs!codUser.Value
        
        cmdq.Execute
        rs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
    rs.Close
    Set cmdq = Nothing
    CopiaFormularioRifiuti = True
    
    Exit Function
CopiaFormularioRifiutiERR:
    MsgBox "CopiaFormularioRifiuti" & Err.Description, vbCritical + vbOKOnly, App.Title
End Function


