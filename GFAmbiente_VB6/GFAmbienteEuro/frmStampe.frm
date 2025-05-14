VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStampe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampe"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox DTPAnno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmStampe.frx":0000
      Left            =   5190
      List            =   "frmStampe.frx":0002
      TabIndex        =   4
      Top             =   3097
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stampa"
      Default         =   -1  'True
      Height          =   390
      Index           =   1
      Left            =   4530
      TabIndex        =   2
      Top             =   3690
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   390
      Index           =   0
      Left            =   5895
      TabIndex        =   1
      Top             =   3690
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2895
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Stampa"
         Text            =   "Stampa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Descrizione"
         Text            =   "Descrizione"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Stampata"
         Text            =   "Stampata"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Anno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4425
      TabIndex        =   5
      Top             =   3127
      Width           =   735
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   3120
      Width           =   4065
   End
End
Attribute VB_Name = "frmStampe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim mCodComune As String
Dim mAnno As String

'documenti word
Dim cWord As New Collection
'oggetto word
Dim oWord As Object

Private Const TAG_PWD_FLD = SEPARATORE & "PWD"

'recordset
Private rsDatiFinali As ADODB.Recordset
Private rsSvuotamentiPerTipo As ADODB.Recordset
Private rsListino507UD As ADODB.Recordset
Private rsListino507UA As ADODB.Recordset
Private rsUTDomentiche As ADODB.Recordset
Private rsUTNonDomentiche As ADODB.Recordset
Private rsStatistiche As ADODB.Recordset
Private mdbStampe As New ADODB.Connection
Private soloStampa As Boolean

'costanti word
Private Const wdLine = 5
Private Const wdExtend = 1
Private Const wdAlignParagraphCenter = 1
Private Const wdAlignParagraphLeft = 0
Private Const wdAlignParagraphRight = 2
Private Const wdCharacter = 1
Private Const wdAllowOnlyFormFields = 2
Private Const wdNoProtection = -1
Private Const wdAdjustNone = 0
Private Const wdLineStyleNone = 0
Private Const wdBorderBottom = -3
Private Const wdBorderDiagonalDown = -7
Private Const wdBorderDiagonalUp = -8
Private Const wdBorderHorizontal = -5
Private Const wdBorderLeft = -2
Private Const wdBorderRight = -4
Private Const wdBorderTop = -1
Private Const wdBorderVertical = -6
Private Const wdCell = 12

Sub Carica(Optional ByVal sCodComune As String = vbNullString, Optional ByVal sComune As String = vbNullString, Optional ByVal sAnnoRiferimento As String = vbNullString, Optional ByVal CodStampa As String = "")
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim SubItem As ListSubItem
    
    
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
    
        sAnnoRiferimento = IIf(sAnnoRiferimento = vbNullString, Year(Now), sAnnoRiferimento)

        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
        
    End If
    
    
    If sComune = vbNullString Then
        Set rs = gDBC.Execute("Select * from Comuni where CodIstat='" & sCodComune & "'")
        sComune = rs(0)
    End If
    
    mCodComune = sCodComune
    mAnno = sAnnoRiferimento
    
    
    Load frmStampe

    lvw.ListItems.Clear
    If CodStampa <> vbNullString Then
        Set rs = gDBC.Execute("Select * from Stampe Where CodStampa = '" & CodStampa & "'")
        soloStampa = True
    Else
        soloStampa = False
        Set rs = gDBC.Execute("Select * from Stampe order by CodStampa ")
    End If
    lbl = sComune & " - " & sAnnoRiferimento
    DTPAnno_click
    
    lvw.ColumnHeaders(1).Width = lvw.Width * 10 / 100
    lvw.ColumnHeaders(2).Width = lvw.Width * 80 / 100
    lvw.ColumnHeaders(3).Width = 0
    lvw.ColumnHeaders(4).Width = 0
    lvw.ColumnHeaders(5).Width = 0
    Set lvw.Icons = MDIMain.imglist
    Set lvw.SmallIcons = MDIMain.imglist
    
    Do While Not rs.EOF
        Set Item = lvw.ListItems.Add(, "_" & rs("CodStampa"), rs("CodStampa"))
        Item.ForeColor = vbBlue
        Set SubItem = Item.ListSubItems.Add(, , rs("Descrizione"))
        SubItem.ForeColor = &H4000&
        Set SubItem = Item.ListSubItems.Add(, "NomeFile", rs("NomeFile"))
        SubItem.Tag = SetValueInTag(TAG_PWD_FLD, IIf(IsNull(rs("pwd")), "", rs("pwd")), SubItem.Tag)
        Item.ListSubItems.Add , "Caricata", "0"
        Item.ListSubItems.Add , "Stampata", "0", "Icon_Null"
        rs.MoveNext
    Loop

    If soloStampa Then
        'Apro documento e stampo
        Command1_Click (1)
        'Chiudo Documento
        Command1_Click (1)
        Unload frmStampe
    Else
        Me.Show
    End If
End Sub

Private Sub Command1_Click(index As Integer)
    Dim i As Integer
    Dim NomeDoc As String
    Dim pwd As String
    
    Select Case index
        Case 0
            Unload Me
        Case 1
            For i = 1 To lvw.ListItems.Count
                With lvw.ListItems(i)
                    If .Selected Then
                        If Not CaricaDatiPerStampa(mCodComune, mAnno, .Key, lvw) Then Exit For
                        NomeDoc = App.Path & "\Stampe\" & .ListSubItems("NomeFile").Text
                        pwd = GetValueInTag(TAG_PWD_FLD, .ListSubItems("NomeFile").Tag)
                        
                        If lvw.ListItems(.Key).ListSubItems("Stampata") = "0" Then
                            If ApriDoc(NomeDoc, .Key, pwd) Then
                                If Not MDIMain.WindowState = vbMinimized And Not soloStampa Then
                                    MDIMain.WindowState = vbMinimized
                                    MDIMain.Visible = False
                                End If
                            End If
                        Else
                            ChiudiDoc NomeDoc, .Key, pwd
                        End If
                    End If
                End With
            Next
    End Select
End Sub


Private Function CaricaDatiPerStampa(ByVal mCodComune As String, ByVal mAnnoRiferimento As String, ByVal IDStampa As String, lvw As ListView) As Boolean
    Dim mdbAdox As New adox.Catalog
    Dim mParAbitanti As String
    Dim qry As ADODB.Command
    Dim rsANA As Recordset
    Dim rsTEC As Recordset
    Dim rsCC As Recordset
    Dim rsDTFIN As Recordset
    Dim rsDTFIN_1 As Recordset
    Dim rsKB As Recordset
    Dim rsKA As Recordset
    Dim rsKC As Recordset
    Dim rsKD As Recordset
    Dim rsUD As Recordset
    Dim rsUD_1 As Recordset
    Dim rsUA As Recordset
    Dim rsUA_1 As Recordset
    Dim rsDTFIN507 As Recordset
    Dim rsDTFIN507UA As Recordset
    Dim rsDTFIN507UD As Recordset
    Dim rsTotSvuotamenti As Recordset
    
    Dim i As Integer
    
    CaricaDatiPerStampa = False
    
    On Error GoTo CaricaDatiPerStampaERR
    
    'connessione al MDB per le stampe
    If Not mdbStampe Is Nothing Then
        Set mdbStampe = New ADODB.Connection
        mdbStampe.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\stampe\stampe.mdb"
        mdbStampe.IsolationLevel = adXactReadUncommitted
    End If
    Set mdbAdox.ActiveConnection = mdbStampe
        
        
        
    'Dati Tecnici
    Set rsTEC = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    If rsTEC.EOF Then Exit Function
    'Dati Anagrafici
    Set rsANA = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
    'Dati Economici
    Set rsCC = DatiCostiComune(mCodComune, mAnnoRiferimento)
    'Dati Finali Calcolo Tariffa
    Set rsDTFIN = DatiFinaliTariffa(mCodComune, mAnnoRiferimento, TariffaNormalizzata)
    'Dati Finali Calcolo Tariffa anno prec
    Set rsDTFIN_1 = DatiFinaliTariffa(mCodComune, mAnnoRiferimento - 1, TariffaNormalizzata)
    'Dati Finali Calcolo Tariffa 507
    Set rsDTFIN507 = DatiFinaliTariffa(mCodComune, mAnnoRiferimento, Tariffa507)
    'Dati Finali Calcolo Tariffa 507 - Utenza Domestica
    Set rsDTFIN507UD = SelezionaDatiFinali507UD(mCodComune, mAnnoRiferimento)
    'Dati Finali Calcolo Tariffa 507 - Utenza Attivita
    Set rsDTFIN507UA = SelezionaDatiFinali507UA(mCodComune, mAnnoRiferimento)
    'KB
    Set rsKB = SelezionaCoeffK(CoeffKB, mCodComune, mAnnoRiferimento, 0, 0)
    'KA
    Set rsKA = SelezionaCoeffK(CoeffKA, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    'Utenze Domestiche
    Set rsUD = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento)
    'Utenze Domestiche - anno precedente
    Set rsUD_1 = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento - 1)
    
    'Utenze Attivita
    Set rsUA = SelezionaUtenzeAttivita(mCodComune, mAnnoRiferimento)
    'Utenze Attivita - anno precedente
    Set rsUA_1 = SelezionaUtenzeAttivita(mCodComune, mAnnoRiferimento - 1)
    
    'KC
    Set rsKC = SelezionaCoeffK(CoeffKC, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    'KD
    Set rsKD = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    'Statistiche
    Set rsStatistiche = DatiStatisticiComune(mCodComune, mAnnoRiferimento)
    
    'Statistiche
    Set rsTotSvuotamenti = SelezionaTotaliSvuotamentiPerContenitore(mCodComune, mAnnoRiferimento)
    
    
    
    If lvw.ListItems(IDStampa).ListSubItems("Caricata") = "0" Then
        mdbStampe.BeginTrans
        mdbStampe.Execute ("Delete * from DatiFinali")
        mdbStampe.Execute ("Delete * from LISTINO507UA")
        mdbStampe.Execute ("Delete * from LISTINO507UD")
        mdbStampe.Execute ("Delete * from UTDOMESTICHE")
        mdbStampe.Execute ("Delete * from UTNONDOMESTICHE")
        mdbStampe.Execute ("Delete * from CONTENITORI")
        mdbStampe.CommitTrans
    End If
    
    If rsTEC.EOF Or rsANA.EOF Or rsCC.EOF Then
        MsgBox "Attenzione occorre inserire i dati del comune"
        Exit Function
    End If
    
    
    If IDStampa = "_01" Then
    
    
        If lvw.ListItems(IDStampa).ListSubItems("Caricata") = "0" Then
                
            If rsDTFIN507.EOF Then
                MsgBox "Calcolare la tariffa"
                Exit Function
            End If
                
                
            'Inserimento dati Utenze Domestiche e Categorie attivita
            Set qry = mdbAdox.Procedures("InsDati507UD").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            Do While Not rsDTFIN507UD.EOF
                'utenze domestiche
                qry.Parameters("[pCOMUNE]") = rsDTFIN507UD("Comune")
                qry.Parameters("[pCOSTOTOT]") = rsDTFIN507("CostoTotale")
                qry.Parameters("[pGETTITO]") = rsDTFIN507("Gettito")
                qry.Parameters("[pPERCCOPERT]") = rsDTFIN507("Copertura")
                qry.Parameters("[pCMG507]") = Format$(rsDTFIN507("cmg") / rsDTFIN507("CoeffCorrTariffa"), FORMAT5DEC)
                qry.Parameters("[pPRODMED]") = Format$(rsDTFIN507("ProdMedia"), FORMAT5DEC)
                qry.Parameters("[pClasse]") = rsDTFIN507UD("IdNumeriComponenti")
                qry.Parameters("[pDescrizione]") = rsDTFIN507UD("Descrizione")
                qry.Parameters("[pips]") = rsDTFIN507UD("Valore") * rsDTFIN507("CoeffCorrettivoIPS")  'Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), "#0.00")
                qry.Parameters("[piqs]") = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), "#0.00")
                qry.Parameters("[pTariffa]") = rsDTFIN507UD("TariffaSuperficie")
                qry.Execute
                rsDTFIN507UD.MoveNext
            Loop
            
            
            Set qry = mdbAdox.Procedures("InsDati507UA").Command
            qry.Prepared = True
            qry.Parameters.Refresh
            Do While Not rsDTFIN507UA.EOF
                qry.Parameters("[pCOMUNE]") = rsDTFIN507UA("Comune")
                qry.Parameters("[pCOSTOTOT]") = rsDTFIN507("CostoTotale")
                qry.Parameters("[pGETTITO]") = rsDTFIN507("Gettito")
                qry.Parameters("[pPERCCOPERT]") = rsDTFIN507("Copertura")
                qry.Parameters("[pCMG507]") = Format$(rsDTFIN507("cmg") / rsDTFIN507("CoeffCorrTariffa"), FORMAT5DEC)
                qry.Parameters("[pPRODMED]") = Format$(rsDTFIN507("ProdMedia"), FORMAT5DEC)
                qry.Parameters("[pClasse]") = rsDTFIN507UA("idCategoriaAttivita")
                qry.Parameters("[pDescrizione]") = rsDTFIN507UA("Descrizione")
                qry.Parameters("[pips]") = Format$(rsDTFIN507UA("ips") * rsDTFIN507("CoeffCorrettivoIPS"), "#0.00")
                qry.Parameters("[piqs]") = Format$(rsDTFIN507UA("iqs"), "#0.00")
                qry.Parameters("[pTariffa]") = rsDTFIN507UA("tariffa")
                qry.Execute
                
                rsDTFIN507UA.MoveNext
            Loop
                
                        
            lvw.ListItems(IDStampa).ListSubItems("Caricata") = "1"
        End If
    Else
        If lvw.ListItems(IDStampa).ListSubItems("Caricata") = "0" Then
            If rsDTFIN.EOF Then
                MsgBox "Calcolare la tariffa"
                Exit Function
            End If
            
            If rsKB.EOF Or rsKA.EOF Then
                MsgBox "Attenzione occorre inserire i coefficienti KB e KA"
                Exit Function
            End If
            
            If rsKC.EOF Or rsKD.EOF Then
                MsgBox "Attenzione occorre inserire i coefficienti KC e KD"
                Exit Function
            End If
        
            
            'Inserimento dati Utenze Domestiche
            Set qry = mdbAdox.Procedures("insDatiUtenzeDomestiche").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            Do While Not rsKA.EOF
                qry.Parameters("pCOMUNE") = rsANA("Comune")
                qry.Parameters("pComp") = rsKA("IdNumeroComponenti")
                qry.Parameters("pmq") = rsUD("Superficie")
                qry.Parameters("pnumeroRuoli") = rsUD("NumeroUtenze")
                qry.Parameters("pKA") = Round(rsKA("valore"), 2)
                qry.Parameters("pKB") = Round(rsKB("valore"), 2)
                qry.Parameters("pTF") = rsUD("TariffaSuperficie")
                qry.Parameters("pTV") = rsUD("TariffaNComp")
                
                If Not rsUD_1.EOF Then
                    qry.Parameters("pTF_1") = rsUD_1("TariffaSuperficie")
                    qry.Parameters("pTV_1") = rsUD_1("TariffaNComp")
                Else
                    qry.Parameters("pTF_1") = 0
                    qry.Parameters("pTV_1") = 0
                End If
                
                qry.Execute
                
                
                rsKA.MoveNext
                rsUD.MoveNext
                If Not rsUD_1.EOF Then rsUD_1.MoveNext
                
                rsKB.MoveNext
            Loop
                
            'Inserimento dati Utenze Attivita
            Set qry = mdbAdox.Procedures("insDatiUtenzeAttivita").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            Do While Not rsKC.EOF
                If rsUA("Superficie") > 0 Then
                    qry.Parameters("pCOMUNE") = rsANA("Comune")
                    qry.Parameters("pCAT") = rsKC("IdCategoriaAttivita")
                    qry.Parameters("pCLASS") = rsKC("Descrizione")
                    qry.Parameters("pKC") = Round(rsKC("valore"), 2)
                    qry.Parameters("pKD") = Round(rsKD("valore"), 2)
                    qry.Parameters("pTF") = rsUA("TariffaPF")
                    qry.Parameters("pmq") = rsUA("Superficie")
                    qry.Parameters("pnumeroRuoli") = rsUA("NumRuoli")
                    qry.Parameters("pTV") = rsUA("TariffaPV")
                    qry.Parameters("pTARIFFA") = rsUA("Tariffa")
                    
                    If Not rsUA_1.EOF Then
                        qry.Parameters("pTARIFFA_1") = rsUA_1("Tariffa")
                    Else
                        qry.Parameters("pTARIFFA_1") = 0
                    End If
                    qry.Execute
                End If
                rsKC.MoveNext
                rsUA.MoveNext
                rsKD.MoveNext
                If Not rsUA_1.EOF Then rsUA_1.MoveNext
            Loop
                
                
            Set qry = mdbAdox.Procedures("InsDatiFinali").Command
            qry.Parameters.Refresh
            qry.Parameters("pCOMUNE") = rsANA("Comune")
            qry.Parameters("pCOSTOTOT") = rsCC("Tot_Servizio")
            qry.Parameters("pGETTITO") = rsCC("gettito")
            qry.Parameters("pPERCCOPERT") = 100
            qry.Parameters("pNUMABIT") = rsANA("NumAbitantiResidenti")
            qry.Parameters("pTARREGIME") = rsCC("Tot_Servizio")
            qry.Parameters("pCSL") = rsCC("CSL")
            qry.Parameters("pAC") = rsCC("AC")
            qry.Parameters("pCARC") = rsCC("CARC")
            qry.Parameters("pCGG") = rsCC("CGG")
            qry.Parameters("pCCD") = rsCC("CCD")
            qry.Parameters("pCK") = rsCC("CK")
            qry.Parameters("pCRT") = rsCC("CRT")
            qry.Parameters("pCTS") = rsCC("CTS")
            qry.Parameters("pCRD") = rsCC("CRD")
            qry.Parameters("pCTR") = rsCC("CTR")
            qry.Parameters("pTASSOINFL") = GetParamValue(mAnnoRiferimento, mCodComune, "INFLAZIO")
            qry.Parameters("pPERCREC") = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFRIP")
            qry.Parameters("pTASSOREM") = rsCC("R")
            qry.Parameters("pPVTAR") = rsCC("ParteVariabile")
            qry.Parameters("pPFTAR") = rsCC("ParteFissa")
            qry.Parameters("pQVDOM") = rsDTFIN("QVDOM")
            qry.Parameters("pQFDOM") = rsDTFIN("QFDOM")
            qry.Parameters("pQUVDOM") = rsDTFIN("QUVDOM")
            qry.Parameters("pQUFDOM") = rsDTFIN("QUFDOM")
            qry.Parameters("pCUDOM") = rsDTFIN("CUDOM")
            qry.Parameters("pQVNDOM") = rsDTFIN("QVNDOM")
            qry.Parameters("pQFNDOM") = rsDTFIN("QFNDOM")
            qry.Parameters("pCUNDOM") = rsDTFIN("CUNDOM")
            qry.Parameters("pQAPFNDOM") = rsDTFIN("QAPFNDOM")
            qry.Parameters("pTotSuperficeUA") = rsTEC("TotSuperficieUtenzaAttivita")
            qry.Parameters("pTotSuperficeUD") = rsTEC("TotSuperficieUtenzaDomestica")
            qry.Parameters("pTotRuoliUA") = rsTEC("NRuoliUtenzaAttivita")
            qry.Parameters("pTotRuoliUD") = rsTEC("NRuoliUtenzaDomestica")
            qry.Parameters("pQRifiutiTotali") = rsTEC("QRifiutiTotali")
            qry.Parameters("pQRifiutiRaccoltaDifferenziata") = rsTEC("QRifiutiRaccoltaDifferenziata")
            qry.Parameters("pQRifiutiCategorieAttivita") = rsTEC("QRifiutiCategorieAttivita")
            qry.Parameters("pQRifiutiUtenzaDomestica") = rsTEC("QRifiutiUtenzaDomestica")
            qry.Parameters("pProduzioneMediaRifiuti") = rsTEC("ProduzioneMediaRifiuti")
            qry.Parameters("pQRifiutiPerAbitante") = rsTEC("QRifiutiPerAbitante")
            qry.Parameters("pCostoPerGestioneKGSmaltiti") = rsTEC("CostoPerGestioneKGSmaltiti")
            qry.Parameters("pProduzionePerUtenzaDomestica") = rsTEC("ProduzionePerUtenzaDomestica")
            qry.Parameters("pProduzionePerUtenzaAttivita") = rsTEC("ProduzionePerUtenzaAttivita")
            qry.Parameters("pCONAI") = rsCC("CON")
            qry.Parameters("pY") = rsCC("Y")
            qry.Parameters("pACC") = rsCC("ACC")
            qry.Parameters("pAMM") = rsCC("AMM")
            qry.Parameters("pIntroitoDaUA_PF") = rsDTFIN("IntroitoDaUA_PF")
            qry.Parameters("pIntroitoDaUA_PV") = rsDTFIN("IntroitoDaUA_PV")
            qry.Parameters("pIntroitoDaUD_PF") = rsDTFIN("IntroitoDaUD_PF")
            qry.Parameters("pIntroitoDaUD_PV") = rsDTFIN("IntroitoDaUD_PV")
            qry.Parameters("pGettitoUD") = rsDTFIN("GettitoUD")
            qry.Parameters("pGettitoUA") = rsDTFIN("GettitoUA")
            If Not rsDTFIN_1.EOF Then
                qry.Parameters("pGettitoUD_1") = rsDTFIN_1("GettitoUD")
                qry.Parameters("pGettitoUA_1") = rsDTFIN_1("GettitoUA")
            Else
                qry.Parameters("pGettitoUD_1") = 0
                qry.Parameters("pGettitoUA_1") = 0
            End If
            
            
            qry.Execute
        
        
            'Inserimento dati Svuotamento
            Set qry = mdbAdox.Procedures("insContenitori").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            Do While Not rsTotSvuotamenti.EOF
                qry.Parameters("pCOMUNE") = rsTotSvuotamenti("CodComune")
                qry.Parameters("pID") = rsTotSvuotamenti("idContenitore")
                qry.Parameters("pSVUOTAMENTI") = rsTotSvuotamenti("Svuotamenti")
                qry.Parameters("pCOEFF") = 1
                

                qry.Execute
                rsTotSvuotamenti.MoveNext
            Loop

        
        
            'metto in caricate tutte le stampe
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Key <> "_01" Then lvw.ListItems(i).ListSubItems("Caricata") = "1"
            Next
        
        
        End If
        

    End If
    
    If Err = 0 Then
        CaricaDatiPerStampa = True
        SelezionoDatiStampe mCodComune, mAnnoRiferimento, IDStampa
    End If

    Set mdbAdox = Nothing
    Set qry = Nothing
    Set rsANA = Nothing
    Set rsTEC = Nothing
    Set rsCC = Nothing
    Set rsDTFIN = Nothing
    Set rsDTFIN_1 = Nothing
    Set rsKB = Nothing
    Set rsKA = Nothing
    Set rsUD = Nothing
    Set rsUA = Nothing
    Set rsUA_1 = Nothing
    Set rsUD_1 = Nothing

    Exit Function
CaricaDatiPerStampaERR:
    MsgBox Error$, vbCritical, App.Title
End Function

Private Sub ChiudiDoc(ByVal s As String, ByVal CodiceDoc As String, ByVal pwd As String)
    Dim adoc As Object
    
    On Error Resume Next
    'se il doc esiste non lo carico
    For Each adoc In cWord
        If InStr(1, s & ".doc", adoc.Name) > 0 Then
            Exit For
        End If
    Next
    
    cWord.Remove adoc.Name
    adoc.Close False
    lvw.ListItems(CodiceDoc).ListSubItems("Stampata") = "0"
    lvw.ListItems(CodiceDoc).ListSubItems("Stampata").ReportIcon = "Icon_Null"
    
    If cWord.Count = 0 Then
        oWord.Quit False
        Set oWord = Nothing
        Set cWord = Nothing
    End If
End Sub


Private Function ApriDoc(ByVal s As String, ByVal CodiceDoc As String, ByVal pwd As String) As Boolean
    Dim conStampe As ADODB.Connection
    Dim adoc As Object

    'controllo se WORD è apero
    VerifaDocumentiWord
    
    'se il doc esiste non lo carico
    For Each adoc In cWord
        If InStr(1, s & ".doc", adoc.Name) > 0 Then Exit Function
    Next
    
    
    
    'apro la connessione al mdb Stampe
    Set conStampe = CreateObject("ADODB.Connection")
    conStampe.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\stampe\stampe.mdb" & ";Jet OLEDB:Database Password="
    
    ApriDoc = False
    On Error GoTo ApriDocERR
    
    If oWord Is Nothing Then
        Set oWord = CreateObject("Word.application")
'        If Not soloStampa Then
'            oWord.Visible = True
'            oWord.Activate
'        End If
    End If
    
    
    Set adoc = oWord.Documents.Open(s, , , , pwd)
    
    lvw.ListItems(CodiceDoc).ListSubItems("Stampata") = "1"
    lvw.ListItems(CodiceDoc).ListSubItems("Stampata").ReportIcon = "Icon_PV"
    
    adoc.Activate
    cWord.Add oWord.ActiveDocument, oWord.ActiveDocument.Name
    oWord.Visible = False
    If oWord.ActiveDocument.ProtectionType = wdNoProtection Then
        oWord.ActiveDocument.Protect password:="gfa", NoReset:=False, Type:=wdAllowOnlyFormFields
    End If
    
    
    oWord.ActiveDocument.CommandBars("Mail Merge").Visible = False
    oWord.ActiveDocument.CommandBars("Forms").Visible = False
    

    'Compila il documento con i dati dal MDB
    CompilaIlDocumento CodiceDoc, pwd
    
    
    'stampo il documento se sono in solo stampa
    If soloStampa Then
        oWord.ActiveDocument.PrintOut False
    Else
        oWord.Visible = True
    End If
    
    
    
    ApriDoc = True

    Exit Function
ApriDocERR:
    If Not oWord Is Nothing Then
        oWord.Quit False: Set oWord = Nothing
    End If
    
    MsgBox Error$, vbCritical, App.Title

End Function

Private Sub Form_Load()
    If DTPAnno.ListCount <= 0 Then
        CaricaComboAnniTariffazione
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim adoc As Object
    
    
    'scarico dalla memoria i documenti word aperti
    On Error Resume Next
    For Each adoc In cWord
        cWord.Remove adoc.Name
        adoc.Close False
    Next
    
    
    'chiudo word
    Set cWord = Nothing
    If Not oWord Is Nothing Then
        oWord.Quit False
        Set oWord = Nothing
    End If
    
    On Error GoTo 0
    'connessioni
    Set rsDatiFinali = Nothing
    Set rsListino507UD = Nothing
    Set rsListino507UA = Nothing
    Set rsSvuotamentiPerTipo = Nothing
    Set rsUTDomentiche = Nothing
    Set rsUTNonDomentiche = Nothing
    Set rsStatistiche = Nothing
    
    Set mdbStampe = Nothing
    
    'rimetto fuori il main
    MDIMain.Visible = True
    DoEvents
End Sub

Private Sub Lvw_DblClick()
    If Not lvw.SelectedItem Is Nothing Then
        Command1_Click 1
    End If
End Sub


Private Sub CompilaIlDocumento(ID As String, pwd As String)
    Dim i As Integer
    Dim PrimaRiga As Boolean
    
    Screen.MousePointer = vbHourglass

    With oWord.ActiveDocument
        Select Case ID
            Case "_01"
                'metto il comune
                .FormFields("ComuneID").Result = rsListino507UD("comune")
               
                For i = 1 To .FormFields.Count
                    'metto l'anno
                    If InStr(1, .FormFields(i).Name, "anno") > 0 Then .FormFields(i).Result = mAnno
                Next
            
                'costo totale
                .FormFields("costotot").Result = rsListino507UD("costotot")
                'gettito
                .FormFields("gettito").Result = rsListino507UD("gettito")
                'percentuale copertura
                .FormFields("perccopert").Result = rsListino507UD("perccopert")
                'cmg507
                .FormFields("cmg507").Result = rsListino507UD("cmg507")
                'produzione media
                .FormFields("prodmed").Result = rsListino507UD("prodmed")
            
                i = 1
                Do While Not rsListino507UD.EOF
                    'Tariffa
                    .FormFields("tariffaUD" & CStr(i)).Result = rsListino507UD("tariffa")
                    'Classe
                    '.FormFields("classeUD" & CStr(i)).Result = rsListino507UD("classe")
                    'Des Classe
                    .FormFields("DescrizioneUD" & CStr(i)).Result = rsListino507UD("descrizione")
                    'iqs
                    .FormFields("iqsUD" & CStr(i)).Result = rsListino507UD("iqs")
                    'ips
                    .FormFields("ipsUD" & CStr(i)).Result = rsListino507UD("ips")
                    i = i + 1
                    rsListino507UD.MoveNext
                Loop
            
            
                i = 1
                Do While Not rsListino507UA.EOF
                    'Tariffa
                    .FormFields("tariffa" & CStr(i)).Result = rsListino507UA("tariffa")
                    'Classe
                    .FormFields("classe" & CStr(i)).Result = rsListino507UA("classe")
                    'Des Classe
                    .FormFields("descrizione" & CStr(i)).Result = rsListino507UA("descrizione")
                    'iqs
                    .FormFields("iqs" & CStr(i)).Result = rsListino507UA("iqs")
                    'ips
                    .FormFields("ips" & CStr(i)).Result = rsListino507UA("ips")
                    i = i + 1
                    rsListino507UA.MoveNext
                Loop
                    
            Case "_02"
                'metto il comune
                .FormFields("ComuneID").Result = rsDatiFinali("comune")
            Case "_03"
                'metto il comune
                .FormFields("ComuneID").Result = rsDatiFinali("comune")
                For i = 1 To .FormFields.Count
                    'metto l'anno
                    If .FormFields(i).Name = "anno" Then .FormFields(i).Result = mAnno
                Next
            Case "_04"
                'metto il comune
                .FormFields("ComuneID").Result = rsDatiFinali("comune")
                For i = 1 To .FormFields.Count
                    'metto l'anno
                    If .FormFields(i).Name = "anno" Then .FormFields(i).Result = mAnno
                Next
                'costo totale
                .FormFields("costototale").Result = rsDatiFinali("costotot")
                'gettito
                .FormFields("gettito").Result = rsDatiFinali("gettito")
                'percentuale copertura
                .FormFields("perccopert").Result = rsDatiFinali("perccopert")
                'numero abitanti residenti
                .FormFields("numabitanti").Result = rsDatiFinali("numabit")

            Case "_05"
                For i = 1 To .FormFields.Count
                    'metto il comune
                    If InStr(1, .FormFields(i).Name, "ComuneID") > 0 Then .FormFields(i).Result = rsDatiFinali("comune")
                Next
                'costo totale
                .FormFields("costototale").Result = rsDatiFinali("costotot")
                'tariffa riferimento a regime
                .FormFields("tarregime").Result = rsDatiFinali("tarregime")
                'CSL
                .FormFields("csl").Result = rsDatiFinali("csl")
                'CRT
                .FormFields("crt").Result = rsDatiFinali("crt")
                'AC
                .FormFields("ac").Result = rsDatiFinali("ac")
                'CTS
                .FormFields("cts").Result = rsDatiFinali("cts")
                'CARC
                .FormFields("carc").Result = rsDatiFinali("carc")
                'CRD
                .FormFields("crd").Result = rsDatiFinali("crd")
                'CGG
                .FormFields("cgg").Result = rsDatiFinali("cgg")
                'CTR
                .FormFields("ctr").Result = rsDatiFinali("ctr")
                'CCD
                .FormFields("ccd").Result = rsDatiFinali("ccd")
                'CK
                .FormFields("ck").Result = rsDatiFinali("ck")
                'Tasso inflazione
                .FormFields("tassoinfla").Result = rsDatiFinali("tassoinfl")
                'Percentuale recupero
                .FormFields("percrec").Result = rsDatiFinali("percrec")
                'Tasso Rem
                .FormFields("tassorem").Result = rsDatiFinali("tassorem")
                'Parte Variabile tariffa
                .FormFields("pvtar").Result = rsDatiFinali("pvtar")
                'Parte fissa tariffa
                .FormFields("pftar").Result = rsDatiFinali("pftar")
                'QVDOM
                .FormFields("qvdom").Result = rsDatiFinali("qvdom")
                'QFDOM
                .FormFields("qfdom").Result = rsDatiFinali("qfdom")
                'QUVDOM
                .FormFields("quvdom").Result = rsDatiFinali("quvdom")
                'QUFDOM
                .FormFields("qufdom").Result = rsDatiFinali("qufdom")
                'CUDOM
                .FormFields("cudom").Result = rsDatiFinali("cudom")
                'QVNDOM
                .FormFields("qvndom").Result = rsDatiFinali("qvndom")
                'QFNDOM
                .FormFields("qfndom").Result = rsDatiFinali("qfndom")
                'CUNDOM
                .FormFields("cundom").Result = rsDatiFinali("cundom")
                'QAPFNDOM
                .FormFields("qapfndom").Result = rsDatiFinali("qapfndom")
            Case "_06"
                For i = 1 To .FormFields.Count
                    'metto il comune
                    If InStr(1, .FormFields(i).Name, "ComuneID") > 0 Then .FormFields(i).Result = rsDatiFinali("comune")
                Next
        
                Do While Not rsUTDomentiche.EOF
                    'KA
                    .FormFields("ka" & CStr(rsUTDomentiche("comp"))).Result = rsUTDomentiche("ka")
                    'KB
                    .FormFields("kb" & CStr(rsUTDomentiche("comp"))).Result = rsUTDomentiche("kb")
                    'tf
                    .FormFields("tf" & CStr(rsUTDomentiche("comp"))).Result = rsUTDomentiche("tf")
                    'tv
                    .FormFields("tv" & CStr(rsUTDomentiche("comp"))).Result = rsUTDomentiche("tv")
                    
                    rsUTDomentiche.MoveNext
                Loop
        
            Case "_07"
                For i = 1 To .FormFields.Count
                    'metto il comune
                    If InStr(1, .FormFields(i).Name, "ComuneID") > 0 Then .FormFields(i).Result = rsDatiFinali("comune")
                Next
                i = 1
                Do While Not rsUTNonDomentiche.EOF
                    'categoria
                    .FormFields("cat" & CStr(i)).Result = rsUTNonDomentiche("cat")
                    'des categoria
                    .FormFields("class" & CStr(i)).Result = rsUTNonDomentiche("class")
                    'KC
                    .FormFields("kc" & CStr(i)).Result = rsUTNonDomentiche("kc")
                    'KD
                    .FormFields("kd" & CStr(i)).Result = rsUTNonDomentiche("kd")
                    'TF
                    .FormFields("tf" & CStr(i)).Result = rsUTNonDomentiche("tf")
                    'TV
                    .FormFields("tv" & CStr(i)).Result = rsUTNonDomentiche("tv")
                    'Tariffa
                    .FormFields("tariffa" & CStr(i)).Result = rsUTNonDomentiche("tariffa")
                    
                    i = i + 1
                    rsUTNonDomentiche.MoveNext
                Loop
            Case "_09"
                'metto il comune
                .FormFields("ComuneID").Result = rsDatiFinali("comune")
                .FormFields("ComuneID1").Result = rsDatiFinali("comune")
                                    'metto l'anno
                .FormFields("ANNO").Result = mAnno

                
                'UA
                rsUTNonDomentiche.MoveFirst
                i = 1
                
                Do While Not rsUTNonDomentiche.EOF
                    .FormFields("CADES" & CStr(i)).Result = rsUTNonDomentiche("class")
                    .FormFields("CAMQ" & CStr(i)).Result = rsUTNonDomentiche("mq")
                    .FormFields("CANR" & CStr(i)).Result = rsUTNonDomentiche("numeroRuoli")
                    .FormFields("CATA" & CStr(i)).Result = rsUTNonDomentiche("tariffa")
                    .FormFields("CAPF" & CStr(i)).Result = rsUTNonDomentiche("tf")
                    .FormFields("CAPV" & CStr(i)).Result = rsUTNonDomentiche("tv")
                    .FormFields("CATA" & CStr(i) & "_1").Result = rsUTNonDomentiche("tariffa_1")
                    .FormFields("KC" & CStr(i)).Result = rsUTNonDomentiche("kc")
                    .FormFields("KD" & CStr(i)).Result = rsUTNonDomentiche("kd")
                    
                    i = i + 1
                    rsUTNonDomentiche.MoveNext
                Loop
                
                rsUTDomentiche.MoveFirst
                i = 1
                Do While Not rsUTDomentiche.EOF
                    .FormFields("UDMQ" & CStr(i)).Result = rsUTDomentiche("mq")
                    .FormFields("UDNR" & CStr(i)).Result = rsUTDomentiche("numeroRuoli")
                    .FormFields("UDTF" & CStr(i) & "_1").Result = rsUTDomentiche("tf_1")
                    .FormFields("UDTV" & CStr(i) & "_1").Result = rsUTDomentiche("tv_1")
                    .FormFields("UDTF" & CStr(i)).Result = rsUTDomentiche("tf")
                    .FormFields("UDTV" & CStr(i)).Result = rsUTDomentiche("tv")
                    .FormFields("KA" & CStr(i)).Result = rsUTDomentiche("ka")
                    .FormFields("KB" & CStr(i)).Result = rsUTDomentiche("kb")
                    i = i + 1
                    rsUTDomentiche.MoveNext
                Loop
                
                rsDatiFinali.MoveFirst
                .FormFields("DT_SMQ_UD").Result = rsDatiFinali("TotSuperficeUD")
                .FormFields("DT_SMQ_UA").Result = rsDatiFinali("TotSuperficeUA")
                .FormFields("DT_NR_UD").Result = rsDatiFinali("TotRuoliUD")
                .FormFields("DT_NR_UA").Result = rsDatiFinali("TotRuoliUA")
                .FormFields("QRifiutiTotali").Result = rsDatiFinali("QRifiutiTotali")
                .FormFields("QRifiutiDaRacDiffe").Result = rsDatiFinali("QRifiutiRaccoltaDifferenziata")
                .FormFields("NAbitantiResidenti").Result = rsDatiFinali("NUMABIT")
                .FormFields("NumRuoliUA").Result = rsDatiFinali("TotRuoliUA")
                .FormFields("NumRuoliUD").Result = rsDatiFinali("TotRuoliUD")
                .FormFields("ProduzTotRifUD").Result = rsDatiFinali("QRifiutiUtenzaDomestica")
                .FormFields("ProduzTotRifUA").Result = rsDatiFinali("QRifiutiCategorieAttivita")
                .FormFields("ProduzMediaRifiuti").Result = rsDatiFinali("ProduzioneMediaRifiuti")
                .FormFields("QRifiutiPerAbitante").Result = rsDatiFinali("QRifiutiPerAbitante")
                .FormFields("CostoGestionePerKG").Result = rsDatiFinali("CostoPerGestioneKGSmaltiti")
                .FormFields("ProduzUnitariaUD").Result = rsDatiFinali("ProduzionePerUtenzaDomestica")
                .FormFields("ProduzUnitariaUA").Result = rsDatiFinali("ProduzionePerUtenzaAttivita")
            
                .FormFields("CostoTotale").Result = rsDatiFinali("COSTOTOT")
                .FormFields("CSL").Result = rsDatiFinali("CSL")
                .FormFields("CRT").Result = rsDatiFinali("CRT")
                .FormFields("CTS").Result = rsDatiFinali("CTS")
                .FormFields("AC").Result = rsDatiFinali("AC")
                .FormFields("CRD").Result = rsDatiFinali("CRD")
                .FormFields("CTR").Result = rsDatiFinali("CTR")
                
                
                .FormFields("CARC").Result = rsDatiFinali("CARC")
                .FormFields("CGG").Result = rsDatiFinali("CGG")
                .FormFields("CCD").Result = rsDatiFinali("CCD")
                
                .FormFields("AMM").Result = rsDatiFinali("AMM")
                .FormFields("ACC").Result = rsDatiFinali("ACC")
                .FormFields("R").Result = rsDatiFinali("tassorem")
                .FormFields("CONAI").Result = rsDatiFinali("CONAI")
                .FormFields("Y").Result = rsDatiFinali("Y")
                .FormFields("QFDOM").Result = rsDatiFinali("QFDOM")
                .FormFields("QFNDOM").Result = rsDatiFinali("QFNDOM")
                .FormFields("QuotaTotPF").Result = rsDatiFinali("QFNDOM") + rsDatiFinali("QFDOM")
                .FormFields("QuotaTotPV").Result = rsDatiFinali("QVNDOM") + rsDatiFinali("QVDOM")
                .FormFields("MANC_INTROITO_UD_PV").Result = rsDatiFinali("IntroitoDaUD_PV")
                .FormFields("MANC_INTROITO_UD_PF").Result = rsDatiFinali("IntroitoDaUD_PF")
                .FormFields("MANC_INTROITO_UA_PV").Result = rsDatiFinali("IntroitoDaUA_PV")
                .FormFields("MANC_INTROITO_UA_PF").Result = rsDatiFinali("IntroitoDaUA_PF")
                .FormFields("GettitoUD").Result = rsDatiFinali("GettitoUD")
                .FormFields("GettitoUA").Result = rsDatiFinali("GettitoUA")
                '.FormFields("GettitoUD_1").Result = rsDatiFinali("GettitoUD_1")
                '.FormFields("GettitoUA_1").Result = rsDatiFinali("GettitoUA_1")
                
                For i = 1 To .fields.Count
                    If .fields(i).Type <> 70 Then
                        .fields(i).Update
                    End If
                Next
                
            Case "_10"
                'metto il comune
                .FormFields("ComuneID").Result = rsDatiFinali("comune")
                .FormFields("anno").Result = mAnno
                .FormFields("pvtar").Result = rsDatiFinali("pvtar")
                
                Dim idriga As Integer
                Dim totLitri As Long
                Dim totPeso As Double
                Dim totSvuotamenti As Long
                
                Do While Not rsSvuotamentiPerTipo.EOF
                
                    idriga = idriga + 1
                    .FormFields("TIPOCONT" & CStr(idriga)).Result = rsSvuotamentiPerTipo("id")
                    .FormFields("NUMSVUOTA" & CStr(idriga)).Result = rsSvuotamentiPerTipo("Svuotamenti")
                    .FormFields("LT" & CStr(idriga)).Result = rsSvuotamentiPerTipo("Svuotamenti") * rsSvuotamentiPerTipo("id")
                    .FormFields("COEFF" & CStr(idriga)).Result = 1
                    .FormFields("PESO" & CStr(idriga)).Result = rsSvuotamentiPerTipo("Svuotamenti") * rsSvuotamentiPerTipo("id")
                    
                    totPeso = totPeso + rsSvuotamentiPerTipo("Svuotamenti") * rsSvuotamentiPerTipo("id")
                    totLitri = totLitri + rsSvuotamentiPerTipo("Svuotamenti") * rsSvuotamentiPerTipo("id")
                    totSvuotamenti = totSvuotamenti + rsSvuotamentiPerTipo("Svuotamenti")
                    
                    rsSvuotamentiPerTipo.MoveNext
                Loop
                
                .FormFields("NUMSVUOTA22").Result = totSvuotamenti
                .FormFields("LT22").Result = totLitri
                .FormFields("PESO22").Result = totPeso
                
                .FormFields("totLitri").Result = totLitri
                .FormFields("PESOTOT").Result = totPeso
                rsUTDomentiche.MoveFirst
                .FormFields("Tariffa").Result = rsUTDomentiche("tv")
                
            Case "_08"
                'tolgo protezione al documento per compilarlo
                oWord.ActiveDocument.Unprotect IIf(pwd = vbNullString, "gfa", pwd)
                
                With oWord.Selection
                    'stampa statistiche
                    .Font.size = 26
                    .TypeText "STATISTICHE"
                    .HomeKey wdLine, wdExtend
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    
                    .MoveRight wdCharacter, 1
                    .TypeParagraph
                    .Font.size = 14
                    .TypeText "Del Comune di " & rsDatiFinali("comune") & " del " & mAnno
                    
                    .MoveRight wdCharacter, 1
                    .TypeParagraph
                    .TypeParagraph
                    .Font.size = 10
                    .ParagraphFormat.Alignment = wdAlignParagraphLeft
                    PrimaRiga = True
                    Do While Not rsStatistiche.EOF
                        
                        If PrimaRiga Then
                            oWord.ActiveDocument.Tables.Add Range:=oWord.Selection.Range, NumRows:=1, NumColumns:=3
                            PrimaRiga = False
                            
                            .Tables(1).Select
                            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
                            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
                            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
                                
                            .TypeText Text:=CStr(rsStatistiche("Descrizione"))
                            .MoveRight Unit:=wdCell
                            .ParagraphFormat.Alignment = wdAlignParagraphRight
                            .TypeText Text:=InValutaEst(Replace(rsStatistiche("Valore").value, ".", ","), "", rsStatistiche("decimali"))
                            .MoveRight Unit:=wdCell
                            .TypeText Text:=CStr(rsStatistiche("UnitaMisura"))
                            With .ParagraphFormat
                                .LeftIndent = oWord.CentimetersToPoints(2.16)
                                .SpaceBeforeAuto = False
                                .SpaceAfterAuto = False
                            End With
                            With .ParagraphFormat
                                .SpaceBeforeAuto = False
                                .SpaceAfterAuto = False
                                .FirstLineIndent = oWord.CentimetersToPoints(-2.16)
                            End With
                            
                            .Tables(1).Columns(1).SetWidth ColumnWidth:=310, RulerStyle:=wdAdjustNone
                            .Tables(1).Columns(2).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustNone
                            .Tables(1).Columns(3).SetWidth ColumnWidth:=80, RulerStyle:=wdAdjustNone
                        Else
                            .InsertRowsBelow 1
                            .MoveRight Unit:=wdCell
                            .TypeText Text:=CStr(rsStatistiche("Descrizione"))
                            .MoveRight Unit:=wdCell
                            .TypeText Text:=InValutaEst(Replace(rsStatistiche("Valore").value, ".", ","), "", rsStatistiche("decimali"))
                            .MoveRight Unit:=wdCell
                            .TypeText Text:=CStr(rsStatistiche("UnitaMisura"))
                    
                        End If
                        rsStatistiche.MoveNext
                    Loop
                    
                    
                    
                End With
                
                'non permetto modifiche al documento
                oWord.ActiveDocument.Protect wdAllowOnlyFormFields, False, IIf(pwd = vbNullString, "gfa", pwd)
        End Select
    End With
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub SelezionoDatiStampe(ByVal mCodComune As String, ByVal mAnnoRiferimento As String, IDStampa As String)
    
    
    If IDStampa = "_01" Then
        'caso 507
        If rsListino507UD Is Nothing Then
            Set rsListino507UD = CreateObject("ADODB.Recordset")
        End If
        If rsListino507UD.State <> adStateOpen Then
            rsListino507UD.CursorLocation = adUseClient
            rsListino507UD.CursorType = adOpenStatic
            rsListino507UD.LockType = adLockBatchOptimistic
            Set rsListino507UD.ActiveConnection = mdbStampe
            rsListino507UD.Source = "select * from Listino507UD"
            rsListino507UD.Open
            Set rsListino507UD.ActiveConnection = Nothing
        Else
            rsListino507UD.MoveFirst
        End If
        
        If rsListino507UA Is Nothing Then
            Set rsListino507UA = CreateObject("ADODB.Recordset")
        End If
        If rsListino507UA.State <> adStateOpen Then
            rsListino507UA.CursorLocation = adUseClient
            rsListino507UA.CursorType = adOpenStatic
            rsListino507UA.LockType = adLockBatchOptimistic
            Set rsListino507UA.ActiveConnection = mdbStampe
            rsListino507UA.Source = "select * from Listino507UA"
            rsListino507UA.Open
            Set rsListino507UA.ActiveConnection = Nothing
        Else
            rsListino507UA.MoveFirst
        End If
    Else
        'seleziono i dati e li metto in recordset
        If rsDatiFinali Is Nothing Then
            Set rsDatiFinali = CreateObject("ADODB.Recordset")
        End If
        If rsDatiFinali.State <> adStateOpen Then
            rsDatiFinali.CursorLocation = adUseClient
            rsDatiFinali.CursorType = adOpenStatic
            rsDatiFinali.LockType = adLockBatchOptimistic
            Set rsDatiFinali.ActiveConnection = mdbStampe
            rsDatiFinali.Source = "select * from DatiFinali"
            rsDatiFinali.Open
            Set rsDatiFinali.ActiveConnection = Nothing
        Else
            rsDatiFinali.MoveFirst
        End If
        
        If rsUTDomentiche Is Nothing Then
            Set rsUTDomentiche = CreateObject("ADODB.Recordset")
        End If
        If rsUTDomentiche.State <> adStateOpen Then
            rsUTDomentiche.CursorLocation = adUseClient
            rsUTDomentiche.CursorType = adOpenStatic
            rsUTDomentiche.LockType = adLockBatchOptimistic
            Set rsUTDomentiche.ActiveConnection = mdbStampe
            rsUTDomentiche.Source = "select * from UTDomestiche"
            rsUTDomentiche.Open
            Set rsUTDomentiche.ActiveConnection = Nothing
        Else
            rsUTDomentiche.MoveFirst
        
        End If
        
        If rsUTNonDomentiche Is Nothing Then
            Set rsUTNonDomentiche = CreateObject("ADODB.Recordset")
        End If
        If rsUTNonDomentiche.State <> adStateOpen Then
            rsUTNonDomentiche.CursorLocation = adUseClient
            rsUTNonDomentiche.CursorType = adOpenStatic
            rsUTNonDomentiche.LockType = adLockBatchOptimistic
            Set rsUTNonDomentiche.ActiveConnection = mdbStampe
            rsUTNonDomentiche.Source = "select * from UTNonDomestiche"
            rsUTNonDomentiche.Open
            Set rsUTNonDomentiche.ActiveConnection = Nothing
        Else
            rsUTNonDomentiche.MoveFirst
        End If
    
    
        'seleziono i dati e li metto in recordset
        If rsSvuotamentiPerTipo Is Nothing Then
            Set rsSvuotamentiPerTipo = CreateObject("ADODB.Recordset")
        End If
        If rsSvuotamentiPerTipo.State <> adStateOpen Then
            rsSvuotamentiPerTipo.CursorLocation = adUseClient
            rsSvuotamentiPerTipo.CursorType = adOpenStatic
            rsSvuotamentiPerTipo.LockType = adLockBatchOptimistic
            Set rsSvuotamentiPerTipo.ActiveConnection = mdbStampe
            rsSvuotamentiPerTipo.Source = "select * from Contenitori"
            rsSvuotamentiPerTipo.Open
            Set rsSvuotamentiPerTipo.ActiveConnection = Nothing
        Else
            On Error Resume Next
            rsSvuotamentiPerTipo.MoveFirst
        End If
    
    
    
    End If
    
End Sub


Private Sub VerifaDocumentiWord()
    Dim docName As String
    Dim i As Integer
    On Error Resume Next
    
    For i = 1 To lvw.ListItems.Count
        Err.Clear
        If lvw.ListItems(i).ListSubItems("Stampata").ReportIcon <> "Icon_Null" Then
            docName = cWord(lvw.ListItems(i).ListSubItems("NomeFile").Text & ".doc").Name
            If Err <> 0 Then
                lvw.ListItems(i).ListSubItems("Stampata").ReportIcon = "Icon_Null"
                lvw.ListItems(i).ListSubItems("Stampata").Text = "0"
            End If
        End If
    Next
    
    For i = cWord.Count To 1 Step -1
        Err.Clear
        docName = cWord(i).Name
        If Err <> 0 Then cWord.Remove (i)
    Next
    
    If cWord.Count = 0 Then Set oWord = Nothing
End Sub


Private Sub CaricaComboAnniTariffazione()
    Dim rs As Recordset
    
    Set rs = gDBC.Execute("SELECT DISTINCT AnagraficaComunale.AnnoRiferimento, AnagraficaComunale.flgSimulazione " & _
                         " From AnagraficaComunale " & _
                         " WHERE AnagraficaComunale.CodComune='" & mCodComune & "' AND " & _
                         " AnagraficaComunale.NumAbitantiResidenti>0 AND AnagraficaComunale.flgSimulazione=False ")
                         
    DTPAnno.Clear
    Do While Not rs.EOF
        
        
        DTPAnno.AddItem rs(0)
        If rs(0).value = Year(Now) Then
            DTPAnno.ListIndex = DTPAnno.ListCount - 1
        End If
        
        rs.MoveNext
    Loop
        
    If DTPAnno.ListCount > 0 And DTPAnno.ListIndex = -1 Then
        DTPAnno.ListIndex = DTPAnno.ListCount - 1
    End If
End Sub


Private Sub DTPAnno_click()
    mAnno = DTPAnno.List(DTPAnno.ListIndex)
    If lbl.Caption <> "" Then
        lbl.Caption = Left$(lbl.Caption, InStrRev(lbl.Caption, "- ")) & " " & mAnno
    End If
End Sub
