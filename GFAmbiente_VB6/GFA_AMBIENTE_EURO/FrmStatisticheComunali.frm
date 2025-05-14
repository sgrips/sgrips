VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatisticheComunali 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Stampa"
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Calcola"
      Height          =   390
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Chiudi"
      Height          =   390
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      Top             =   4080
      Width           =   1125
   End
   Begin VB.TextBox txtStatistica 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmStatisticheComunali.frx":0000
      Top             =   120
      Width           =   6015
   End
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
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
      CustomFormat    =   "yyyy"
      Format          =   59768835
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
End
Attribute VB_Name = "frmStatisticheComunali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAnnoRiferimento As Integer
Private mCodComune As String
Private sComune As String
Private rsDatiAna As ADODB.Recordset
Private rsDatiTec As ADODB.Recordset
Private rsDatiEco As ADODB.Recordset
Private rsDatiFin As ADODB.Recordset
Private mMaxlen As Integer

Private Type StatisticaRec
    Descrizione As String
    Valore As String
    UM As String
    Decimali As Integer
End Type

Private mStatistica() As StatisticaRec
 

Private Sub Cmd_Click(index As Integer)
    Select Case index
        Case 0
            'Salva
            SalvaStatistiche
            'chiudi
            Unload Me
        Case 1
            txtStatistica = ""
            CalcolaStatistiche
        Case 2
            'stampa
            'frmStatisticheComunali.Visible = False
            frmStampe.Carica mCodComune, sComune, mAnnoRiferimento, "08"
            'MDIMain.WindowState = vbMaximized
            'frmStatisticheComunali.Visible = True
            'frmStatisticheComunali.ZOrder 0
    
    End Select
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.value)
    CaricaDati mAnnoRiferimento
End Sub

Private Sub Form_Load()
    
    MDIMain.mnuComprimi.Enabled = False
    DTPAnno.value = Format$(Now, "yyyy-mm-dd")
    
    
End Sub

Private Sub Form_Resize()
    
    
    Const bordo = 100
    
    Dim NextTop As Integer
    Dim NextLeft As Integer
    
    If WindowState = vbMinimized Then Exit Sub
    
    NextTop = ScaleHeight - Cmd(0).Height - 2 * bordo
    NextLeft = ScaleWidth
    
    txtStatistica.Move 0, 0, NextLeft, NextTop
    NextTop = NextTop + bordo
    NextLeft = ScaleWidth - Cmd(0).Width - bordo
    Cmd(0).Move NextLeft, NextTop
    NextLeft = NextLeft - Cmd(1).Width - bordo
    Cmd(1).Move NextLeft, NextTop
    NextLeft = NextLeft - Cmd(1).Width - bordo
    Cmd(2).Move NextLeft, NextTop
    NextLeft = NextLeft - DTPAnno.Width - bordo
    DTPAnno.Move NextLeft, NextTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub




Public Sub Carica(Optional sCodComune As String = vbNullString)
    
    Set Me.Font = txtStatistica.Font
    mAnnoRiferimento = Year(Now)
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    mCodComune = sCodComune
    
    
    sComune = getDesComune(mCodComune)
    Caption = "Statistiche del comune: " & sComune
    CaricaDati mAnnoRiferimento
    
    
    Me.Show vbModal
End Sub

Property Let Statistica(idStatistica As Integer, s As String)
    
    'mi salvo il valore calcolato
    mStatistica(idStatistica - 1).Valore = s
    'mostro il valore
    txtStatistica = txtStatistica & mStatistica(idStatistica - 1).Descrizione _
        & String((54 - Len(mStatistica(idStatistica - 1).Descrizione)), ".") _
        & String(13 - Len(Format$(s, "#,##0.0" & String(mStatistica(idStatistica - 1).Decimali - 1, "0"))), ".") & "  " _
        & Format$(s, "#,##0.0" & String(mStatistica(idStatistica - 1).Decimali - 1, "0")) _
        & " " & mStatistica(idStatistica - 1).UM _
        & "" & vbCrLf
    If mMaxlen < TextWidth(txtStatistica.Text) Then
        mMaxlen = TextWidth(txtStatistica.Text)
    End If
End Property

Function msgError(ByVal scheda As String, ByVal campo As String) As String
    msgError = "Attenzione non è stato inserito nella" + vbCrLf + "Scheda: " + "[" + scheda + "]" + vbCrLf + "il valore per il campo: " + campo + vbCrLf
End Function
Private Sub CalcolaStatistiche()
    Dim msg As String
    
    On Error GoTo CalcolaStatisticheERR
    
    
    'validazione dati
    If rsDatiTec("NRuoliUtenzaAttivita") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Numero Ruoli Utenza Attività]")
    End If
    
    If rsDatiTec("TotSuperficieIscrRuolo") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Totale Superficie Iscritti Ruolo]")
    End If
    
    If rsDatiTec("NRuoliUtenzaDomestica") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Numero Ruoli Utenze Domestica]")
    End If
    
    
    If rsDatiTec("QRifiutiTotali") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Quantità Rifiuti Totali]")
    End If

    If rsDatiTec("NRuoliTot") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Numero Ruoli Totali]")
    End If

    If rsDatiTec("NRuoliTot") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Numero Ruoli Totali]")
    End If

    If rsDatiEco("Tot_Servizio") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Economici", "[Totale Servizio]")
    End If

    If rsDatiAna("NumAbitantiResidenti") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Anagrafici", "[Numero Abitanti Residenti]")
    End If

    If rsDatiTec("TotSuperficieUtenzaAttivita") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Tot Superficie Utenza Attività]")
    End If

    If rsDatiTec("TotSuperficieUtenzaDomestica") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Dati Tecnici", "[Tot Superficie Utenza Domestica]")
    End If

    If rsDatiFin("QVDOM") + rsDatiFin("QVNDOM") = 0 Then
        Err.Raise 999, "CalcoloStatistiche", msgError("Calcolo Tariffa", "Quota Unitaria Parte Variabile," + vbCrLf + " relativa a Utenza Domestica o Attività")
    End If

    
   
    'Statistica 1
    Statistica(1) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiTec("NRuoliUtenzaDomestica")
    'Statistica 2
    Statistica(2) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiTec("NRuoliUtenzaAttivita")
    'Statistica 3
    Statistica(3) = rsDatiEco("Tot_Servizio") / rsDatiTec("TotSuperficieIscrRuolo")
    'Statistica 4
    Statistica(4) = rsDatiTec("CTotKgRifiutiGestiti")
    'Statistica 5
    Statistica(5) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiTec("QRifiutiTotali")
    'Statistica 6
    Statistica(6) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiTec("QRifiutiTotali")
    'Statistica 7
    Statistica(7) = rsDatiTec("QRifiutiRaccoltaDifferenziata") / rsDatiTec("QRifiutiTotali")
    'Statistica 8
    Statistica(8) = rsDatiTec("QRifiutiTotali") / rsDatiTec("TotSuperficieIscrRuolo")
    'Statistica 9
    Statistica(9) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiTec("TotSuperficieIscrRuolo")
    'Statistica 10
    Statistica(10) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiTec("TotSuperficieIscrRuolo")
    'Statistica 11
    Statistica(11) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiTec("NRuoliTot")
    'Statistica 12
    Statistica(12) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiTec("NRuoliTot")
    'Statistica 13
    Statistica(13) = rsDatiEco("ParteFissa") / rsDatiEco("Tot_Servizio")
    'Statistica 14
    Statistica(14) = rsDatiEco("ParteVariabile") / rsDatiEco("Tot_Servizio")
    'Statistica 15
    Statistica(15) = rsDatiEco("CSL") / rsDatiEco("Tot_Servizio")
    'Statistica 16
    Statistica(16) = rsDatiEco("CRT") / rsDatiEco("Tot_Servizio")
    'Statistica 17
    Statistica(17) = rsDatiEco("CRD") / rsDatiEco("Tot_Servizio")
    'Statistica 18
    Statistica(18) = rsDatiEco("CTS") / rsDatiEco("Tot_Servizio")
    'Statistica 19
    Statistica(19) = rsDatiEco("CTR") / rsDatiEco("Tot_Servizio")
    'Statistica 20
    Statistica(20) = rsDatiEco("CARC") / rsDatiEco("Tot_Servizio")
    'Statistica 21
    Statistica(21) = rsDatiEco("CGG") / rsDatiEco("Tot_Servizio")
    'Statistica 22
    Statistica(22) = rsDatiEco("AC") / rsDatiEco("Tot_Servizio")
    'Statistica 23
    Statistica(23) = rsDatiEco("CCD") / rsDatiEco("Tot_Servizio")
    'Statistica 24
    Statistica(24) = rsDatiEco("CK") / rsDatiEco("Tot_Servizio")
    'Statistica 25
    Statistica(25) = rsDatiTec("QRifiutiTotali") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 26
    Statistica(26) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 27
    Statistica(27) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 28
    Statistica(28) = rsDatiTec("QRifiutiRaccoltaDifferenziata") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 29
    Statistica(29) = rsDatiTec("QRifiutiUtenzaDomestica") / rsDatiTec("TotSuperficieUtenzaDomestica")
    'Statistica 30
    Statistica(30) = rsDatiTec("QRifiutiCategorieAttivita") / rsDatiTec("TotSuperficieUtenzaAttivita")
    'Statistica 31
    Statistica(31) = rsDatiEco("ParteFissa") / rsDatiTec("QRifiutiTotali")
    'Statistica 32
    Statistica(32) = rsDatiEco("ParteVariabile") / rsDatiTec("QRifiutiTotali")
    'Statistica 33
    Statistica(33) = rsDatiFin("QVDOM")
    'Statistica 34
    Statistica(34) = rsDatiFin("QVNDOM")
    'Statistica 35
    Statistica(35) = rsDatiFin("QFDOM")
    'Statistica 36
    Statistica(36) = rsDatiFin("QFNDOM")
    'Statistica 37
    Statistica(37) = rsDatiFin("QVDOM") / (rsDatiFin("QVDOM") + rsDatiFin("QVNDOM"))
    'Statistica 38
    Statistica(38) = rsDatiFin("QVNDOM") / (rsDatiFin("QVDOM") + rsDatiFin("QVNDOM"))
    'Statistica 39
    Statistica(39) = rsDatiFin("QFDOM") / (rsDatiFin("QFDOM") + rsDatiFin("QFNDOM"))
    'Statistica 40
    Statistica(40) = rsDatiFin("QFNDOM") / (rsDatiFin("QFDOM") + rsDatiFin("QFNDOM"))
    'Statistica 41
    Statistica(41) = rsDatiFin("QVDOM") / rsDatiEco("Tot_Servizio")
    'Statistica 42
    Statistica(42) = rsDatiFin("QVNDOM") / rsDatiEco("Tot_Servizio")
    'Statistica 43
    Statistica(43) = rsDatiFin("QFDOM") / rsDatiEco("Tot_Servizio")
    'Statistica 44
    Statistica(44) = rsDatiFin("QFNDOM") / rsDatiEco("Tot_Servizio")
    'Statistica 45
    Statistica(45) = rsDatiEco("Gettito") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 46
    Statistica(46) = (rsDatiFin("QVDOM") + rsDatiFin("QFDOM")) / rsDatiAna("NumAbitantiResidenti")
    'Statistica 47
    Statistica(47) = rsDatiFin("QFDOM") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 48
    Statistica(48) = rsDatiFin("QVDOM") / rsDatiAna("NumAbitantiResidenti")
    'Statistica 49
    Statistica(49) = (rsDatiFin("QVDOM") + rsDatiFin("QFDOM")) / rsDatiTec("NRuoliUtenzaDomestica")
    'Statistica 50
    Statistica(50) = rsDatiFin("QFDOM") / rsDatiTec("NRuoliUtenzaDomestica")
    'Statistica 51
    Statistica(51) = rsDatiFin("QVDOM") / rsDatiTec("NRuoliUtenzaDomestica")
    'Statistica 52
    Statistica(52) = (rsDatiFin("QVNDOM") + rsDatiFin("QFNDOM")) / rsDatiTec("NRuoliUtenzaAttivita")
    'Statistica 53
    Statistica(53) = rsDatiFin("QFNDOM") / rsDatiTec("NRuoliUtenzaAttivita")
    'Statistica 54
    Statistica(54) = rsDatiFin("QVNDOM") / rsDatiTec("NRuoliUtenzaAttivita")

    
    
    Me.Width = mMaxlen + 2 * 23 * 15
    Exit Sub
CalcolaStatisticheERR:
    If Err.Number = 999 Then
        MsgBox Err.description, vbExclamation + vbOKOnly, App.Title
    Else
        MsgBox "Verificare che tutti i dati per il comune siano inseriti e che il calcolo tariffa si stato fatto", vbExclamation, App.Title
    End If
End Sub

Private Sub SalvaStatistiche()
   
    'salvo le statistiche
    gDBC.Execute "Delete from StatisticheComunali where CodComune ='" & mCodComune & "' and AnnoRiferimento='" & DTPAnno.Year & "'"
    
    
    Dim i As Integer
    i = -1
    On Error Resume Next
    i = LBound(mStatistica)
    On Error GoTo 0
    
    If i > -1 Then
        For i = LBound(mStatistica) To UBound(mStatistica)
            gDBC.Execute "Insert Into StatisticheComunali  (CodStatistica,CodComune,AnnoRiferimento,Valore) values(" & i + 1 & ",'" & mCodComune & "','" & DTPAnno.Year & "','" & FormatFLD(mStatistica(i).Valore, adCurrency) & "')"
        Next
    End If
End Sub

Private Sub CaricaDati(ByVal Anno As String)
    Dim rsDesStat As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    On Error GoTo CaricaDatiERR
    
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento)
    
    If Not rs.EOF Then
        gDBC.Execute ("INSERT INTO StatisticheComunali ( CodStatistica, CodComune, AnnoRiferimento, Valore ) " & _
                        "SELECT ST.CodStatistica, '" & mCodComune & "','" & mAnnoRiferimento & "',0  " & _
                        "FROM Statistiche AS ST " & _
                        "WHERE not Exists (select 1 from StatisticheComunali where codStatistica=ST.CodStatistica and CodComune='" & mCodComune & "' and AnnoRiferimento='" & mAnnoRiferimento & "') " & _
                        "ORDER BY ST.CodStatistica")
    End If
        
    'se non esistono le statistiche le inserisco a ZERO
    mMaxlen = 5000
    txtStatistica = ""
    Set rsDesStat = CreateObject("ADODB.Recordset")
    rsDesStat.CursorLocation = adUseClient
    rsDesStat.CursorType = adOpenStatic
    rsDesStat.LockType = adLockBatchOptimistic
    rsDesStat.Source = DatiStatisticiComune(mCodComune, Anno).Source
    
    Set rsDesStat.ActiveConnection = gDBC
    rsDesStat.Open
    Set rsDesStat.ActiveConnection = Nothing
    
    ReDim mStatistica(rsDesStat.RecordCount - 1)
    'metto le Descrizioni in una Collection
    Do While Not rsDesStat.EOF
        mStatistica(rsDesStat("ST.CodStatistica").value - 1).Descrizione = rsDesStat("Descrizione").value
        mStatistica(rsDesStat("ST.CodStatistica").value - 1).Valore = 0
        mStatistica(rsDesStat("ST.CodStatistica").value - 1).UM = rsDesStat("UnitaMisura").value
        mStatistica(rsDesStat("ST.CodStatistica").value - 1).Decimali = rsDesStat("Decimali").value
        If Not IsNull(rsDesStat("Valore").value) Then
            Statistica(rsDesStat("ST.CodStatistica").value) = Replace(rsDesStat("Valore").value, ".", ",")
        End If
        rsDesStat.MoveNext
    Loop
    
    Set rsDesStat = Nothing


    Me.Width = mMaxlen + 1000
    
    Set rsDatiAna = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento)
    Set rsDatiTec = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    Set rsDatiEco = DatiCostiComune(mCodComune, mAnnoRiferimento)
    Set rsDatiFin = DatiFinaliTariffa(mCodComune, mAnnoRiferimento, TariffaNormalizzata)
    
    
    Exit Sub
    
CaricaDatiERR:
End Sub

