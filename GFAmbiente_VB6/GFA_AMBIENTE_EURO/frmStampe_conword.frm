VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStampe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampe"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Stampa"
      Default         =   -1  'True
      Height          =   390
      Index           =   1
      Left            =   4530
      TabIndex        =   2
      Top             =   3240
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   390
      Index           =   0
      Left            =   5895
      TabIndex        =   1
      Top             =   3240
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
      FlatScrollBar   =   -1  'True
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
      NumItems        =   2
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
      Width           =   4290
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
'Dim WithEvents cWords As clsWords
'Dim oAccess As Object

Private Const TAG_PWD_FLD = SEPARATORE & "PWD"

Sub Carica(ByVal CodComune As String, ByVal Comune As String, ByVal Anno As String)
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim SubItem As ListSubItem
    mCodComune = CodComune
    mAnno = Anno
    
    Set cWords = New clsWords
    Load frmStampe
    
    lvw.ListItems.Clear
    Set rs = gDBC.Execute("Select * from Stampe order by CodStampa")
    lbl = Comune & " - " & Anno
    
    lvw.ColumnHeaders(1).Width = lvw.Width * 0.15
    lvw.ColumnHeaders(2).Width = lvw.Width * 0.85
    
    Do While Not rs.EOF
        Set Item = lvw.ListItems.Add(, "_" & rs("CodStampa"), rs("CodStampa"))
        Item.ForeColor = vbBlue
        Set SubItem = Item.ListSubItems.Add(, , rs("Descrizione"))
        SubItem.ForeColor = &H4000&
        Set SubItem = Item.ListSubItems.Add(, "NomeFile", rs("NomeFile"))
        SubItem.Tag = SetValueInTag(TAG_PWD_FLD, IIf(IsNull(rs("pwd")), "", rs("pwd")), SubItem.Tag)
        Item.ListSubItems.Add , "Caricata", False
        rs.MoveNext
    Loop

    Me.Show
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer
    Dim NomeDoc As String
    Dim pwd As String
    Select Case Index
        Case 0
            Unload Me
        Case 1
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Selected Then
                    If Not CaricaDatiPerStampa(mCodComune, mAnno, lvw.ListItems(i).Key, lvw) Then Exit For
                    NomeDoc = App.Path & "\Stampe\" & lvw.ListItems(i).ListSubItems("NomeFile").Text
                    pwd = GetValueInTag(TAG_PWD_FLD, lvw.ListItems(i).ListSubItems("NomeFile").Tag)
                    If ApriDoc(NomeDoc, pwd) Then
                        If Not MDIMain.WindowState = vbMinimized Then
                            MDIMain.WindowState = vbMinimized
                            Me.Visible = False
                            DoEvents
                        End If
                        
                    End If
                End If
            Next
    End Select
End Sub


Function CaricaDatiPerStampa(ByVal mCodComune As String, ByVal mAnnoRiferimento As String, ByVal idStampa As String, lvw As ListView) As Boolean
    Dim mdbStampe As New ADODB.Connection
    Dim mdbAdox As New adox.Catalog
    Dim mParAbitanti As String
    Dim qry As ADODB.Command
    Dim rsANA As Recordset
    Dim rsTEC As Recordset
    Dim rsCC As Recordset
    Dim rsDTFIN As Recordset
    Dim rsKB As Recordset
    Dim rsKA As Recordset
    Dim rsKC As Recordset
    Dim rsKD As Recordset
    Dim rsUD As Recordset
    Dim rsUA As Recordset
    Dim rsDTFIN507 As Recordset
    
    Dim i As Integer
    
    CaricaDatiPerStampa = False
    
    On Error GoTo CaricaDatiPerStampaERR
    
    'connessione al MDB per le stampe
    Set mdbStampe = New ADODB.Connection
    mdbStampe.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\stampe\stampe.mdb"
    mdbStampe.IsolationLevel = adXactReadUncommitted
    Set mdbAdox.ActiveConnection = mdbStampe
        
    'Dati Tecnici
    Set rsTEC = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    'Dati Anagrafici
    Set rsANA = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
    'Dati Economici
    Set rsCC = CostiComune(mCodComune, mAnnoRiferimento)
    'Dati Finali Calcolo Tariffa
    Set rsDTFIN = DatiFinaliTariffa(mCodComune, mAnnoRiferimento, TariffaNormalizzata)
    'Dati Finali Calcolo Tariffa 507
    Set rsDTFIN507 = DatiFinaliTariffa(mCodComune, mAnnoRiferimento, Tariffa507)
    'KB
    Set rsKB = SelezionaCoeffK(CoeffKB, mCodComune, mAnnoRiferimento, 0, 0)
    'KA
    Set rsKA = SelezionaCoeffK(CoeffKA, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    'Utenze Domestiche
    Set rsUD = SelezionaUtenzeDomestiche(mCodComune, mAnnoRiferimento)
    'Utenze Attivita
    Set rsUA = SelezionaUtenzeAttivita(mCodComune, mAnnoRiferimento)
    'KC
    Set rsKC = SelezionaCoeffK(CoeffKC, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    'KD
    Set rsKD = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
    
    If lvw.ListItems(idStampa).ListSubItems("Caricata") = "0" Then
        mdbStampe.BeginTrans
        mdbStampe.Execute ("Delete * from DatiFinali")
        mdbStampe.Execute ("Delete * from LISTINO507")
        mdbStampe.Execute ("Delete * from UTDOMESTICHE")
        mdbStampe.Execute ("Delete * from UTNONDOMESTICHE")
        mdbStampe.CommitTrans
    End If
    
    If rsTEC.EOF Or rsANA.EOF Or rsCC.EOF Then
        MsgBox "Attenzione occorre inserire i dati del comune"
        Exit Function
    End If
    
    
    If idStampa = "_01" Then
    
    
        If lvw.ListItems(idStampa).ListSubItems("Caricata") = "0" Then
                
            If rsDTFIN507.EOF Then
                MsgBox "Calcolare la tariffa"
                Exit Function
            End If
                
                
            'Inserimento dati Utenze Domestiche
            Set qry = mdbAdox.Procedures("InsDati507").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            'utenze domestiche
            qry.Parameters("[pCOMUNE]") = rsDTFIN507("Comune")
            qry.Parameters("[pCOSTOTOT]") = rsDTFIN507("CostoTotale")
            qry.Parameters("[pGETTITO]") = rsDTFIN507("Gettito")
            qry.Parameters("[pPERCCOPERT]") = rsDTFIN507("Copertura")
            qry.Parameters("[pCMG507]") = Format$(rsDTFIN507("cmg"), "#0.00")
            qry.Parameters("[pPRODMED]") = rsDTFIN507("ProdMedia")
            qry.Parameters("[pClasse]") = "UD"
            qry.Parameters("[pDescrizione]") = "Utenza Domestica"
            qry.Parameters("[pips]") = Format$(GetParamValue(mCodComune, "COEFFIPSUD"), "#0.00")
            qry.Parameters("[piqs]") = Format$(GetParamValue(mCodComune, "COEFFIQSUD"), "#0.00")
            qry.Parameters("[pTariffa]") = rsUD("TariffaSuperficie")
            qry.Execute
            
            
            Do While Not rsDTFIN507.EOF
                qry.Parameters("[pCOMUNE]") = rsDTFIN507("Comune")
                qry.Parameters("[pCOSTOTOT]") = rsDTFIN507("CostoTotale")
                qry.Parameters("[pGETTITO]") = rsDTFIN507("Gettito")
                qry.Parameters("[pPERCCOPERT]") = rsDTFIN507("Copertura")
                qry.Parameters("[pCMG507]") = Format$(rsDTFIN507("cmg"), "#0.00")
                qry.Parameters("[pPRODMED]") = rsDTFIN507("ProdMedia")
                qry.Parameters("[pClasse]") = rsDTFIN507("idCategoriaAttivita")
                qry.Parameters("[pDescrizione]") = rsDTFIN507("Descrizione")
                qry.Parameters("[pips]") = Format$(rsDTFIN507("ips"), "#0.00")
                qry.Parameters("[piqs]") = Format$(rsDTFIN507("iqs"), "#0.00")
                qry.Parameters("[pTariffa]") = rsDTFIN507("tariffa")
                qry.Execute
                
                rsDTFIN507.MoveNext
            Loop
                
                        
            lvw.ListItems(idStampa).ListSubItems("Caricata") = "1"
        End If
    Else
        If lvw.ListItems(idStampa).ListSubItems("Caricata") = "0" Then
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
        
            'metto in caricate tutte le stampe
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Key <> "_01" Then lvw.ListItems(i).ListSubItems("Caricata") = "1"
            Next
            
            'Inserimento dati Utenze Domestiche
            Set qry = mdbAdox.Procedures("insDatiUtenzeDomestiche").Command
            qry.Prepared = True
            qry.Parameters.Refresh
        
            Do While Not rsKA.EOF
                qry.Parameters("pCOMUNE") = rsANA("Comune")
                qry.Parameters("pComp") = rsKA("IdNumeroComponenti")
                qry.Parameters("pKA") = rsKA("valore")
                qry.Parameters("pKB") = rsKB("valore")
                qry.Parameters("pTF") = rsUD("TariffaSuperficie")
                qry.Parameters("pTV") = rsUD("TariffaNComp")
                qry.Execute
                
                rsKA.MoveNext
                rsUD.MoveNext
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
                    qry.Parameters("pKC") = rsKC("valore")
                    qry.Parameters("pKD") = rsKD("valore")
                    qry.Parameters("pTF") = rsUA("TariffaPF")
                    qry.Parameters("pTV") = rsUA("TariffaPV")
                    qry.Parameters("pTARIFFA") = rsUA("Tariffa")
                    qry.Execute
                End If
                rsKC.MoveNext
                rsUA.MoveNext
                rsKD.MoveNext
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
            qry.Parameters("pTASSOINFL") = GetParamValue(mCodComune, "INFLAZIO")
            qry.Parameters("pPERCREC") = GetParamValue(mCodComune, "COEFFRIP")
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
            qry.Execute
        End If
    End If
    
    If Err = 0 Then CaricaDatiPerStampa = True

    Set mdbStampe = Nothing
    Set mdbAdox = Nothing
    Set qry = Nothing
    Set rsANA = Nothing
    Set rsTEC = Nothing
    Set rsCC = Nothing
    Set rsDTFIN = Nothing
    Set rsKB = Nothing
    Set rsKA = Nothing
    Set rsUD = Nothing
    Set rsUA = Nothing


    Exit Function
CaricaDatiPerStampaERR:
    MsgBox Error$, vbCritical, App.Title
End Function


Private Function ApriDoc(ByVal s As String, ByVal pwd As String) As Boolean
    Dim oWord As Word.Application
    
    'apro Access così quando il documento viene aperto non deve crearne l'istanza
    If oAccess Is Nothing Then
        Set oAccess = CreateObject("Access.Application")
        oAccess.OpenCurrentDatabase (App.Path & "\stampe\stampe.mdb")
    End If
    
    ApriDoc = False
    On Error GoTo ApriDocERR
    Err.Clear
    Set oWord = CreateObject("Word.Application")
        
    cWords.Add oWord, SEPARATORE & CStr(cWords.Count)
    
    oAccess.Visible = True
    oAccess.RunCommand 11
    
    
    oWord.Visible = True
    oWord.Activate
    oWord.StartupPath = App.Path & "\stampe"
    oWord.Documents.Open s, , , , pwd
    DoEvents
    oAccess.Visible = False
    DoEvents
    oWord.ActiveDocument.CommandBars("Mail Merge").Visible = False
    ApriDoc = True

    Exit Function
ApriDocERR:
    If Not oWord Is Nothing Then
        oWord.Quit: Set oWord = Nothing
    End If
    
    If Not oAccess Is Nothing Then
        oAccess.Quit 2: Set oAccess = Nothing
    End If
    
    MsgBox Error$, vbCritical, App.Title

End Function


Private Sub cWords_DocQuit()
    On Error Resume Next
    MDIMain.WindowState = vbMaximized
    Me.Visible = True
    Me.SetFocus
    Set oAccess = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cWords = Nothing
End Sub

Private Sub Lvw_DblClick()
    If Not lvw.SelectedItem Is Nothing Then
        Command1_Click 1
    End If
End Sub


