Attribute VB_Name = "GestioneFile"
Option Explicit

Public Const SCHEMAINI = "Structur.Ini"

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long



Public Type FLTR200F
    TipoLavorazione As String * 1
    CodceFoglioDiViaggio As String * 5
End Type

Public Type FLTR300F
    TipoLavorazione As String * 1
    CodicePiano As String * 2
    LineaDiPartenza As String * 3
    LineaDiArrivo As String * 3
    Zona As String * 2
    CodiceBaia As String * 3
    Quantit‡Baie As String * 1
    Descrizione As String * 30
    Inviato As Boolean
    Key  As String * 8
End Type

Public Type MsgTipo10
    TipoRecord As String * 2
    NProgressivo As String * 5
    LetturaBarCode As String * 33
    LetturaBarCodeManuale As String * 33
End Type

Public Type MsgTipo11
    TipoRecord As String * 2
    NProgressivo As String * 5
    CodiceBaia As String * 3
    Stato As String * 1
End Type

Public Type Barcode
    LineaPartenza As String * 3
    LineaArrivo As String * 3
    Serie As String * 2
    Segnacollo As String * 7
    Zona As String * 2
    crc As String * 1
End Type


Public gFLTR200F As FLTR200F
Public gFLTR300F() As FLTR300F
Public colKey As Collection
Public MsgTipo10() As MsgTipo10
Public MsgTipo11() As MsgTipo11


Static Sub GestioneMSG(MSG As String)
    Dim sBarcodeHANG As String
    Dim sBarcodeSCANNER As String
    Dim sBarcode As String
    Dim sPosPistola As String
    Dim bTestCRC As Boolean
    Dim sIDParcel As String
    Dim bCercaInPiano As Boolean
    Dim idCom As Long
    Dim Barcode As RecordBarCode
    Dim SalvaSUDB As Boolean
    Dim sDestinazione  As String
    
    sBarcodeHANG = vbNullString
    sBarcodeSCANNER = vbNullString
    sBarcode = vbNullString
    bCercaInPiano = True
    bTestCRC = True
    
        
    '           0                  SCANNER              PISTOLA
    '                1         2         3         4         5         6         7
    '       123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
    '
    'msg = "10|00001|123          #123456789012345678#|123          123456789012345678##|UPD1_BF111"
    'msg = "10|90001|123          #123456789012345678#|                                 |UPD1_BF111"
    'msg = "10|90001|                                 |123          123456789012345678##|UPD1_BF111"
    sIDParcel = "0"
    If Left$(MSG, 2) = "10" Or Right$(Left$(MSG, 2), 1) = "9" Then
        'Arrivato pacco
        If Right$(Left$(MSG, 2), 1) = "9" Then
            'Messaggio per gestione senza Vanderland proveniente da soluzione (Metrologic)
            'STX(1) + IDPISTOLA(1) + IDMSG(1) + MSG(50 variabili)  + CS(1) + ETX(1)
            
            sIDParcel = Format$(Val(sIDParcel) + 1, "00000")
            sBarcodeHANG = Trim$(Left$(Mid$(MSG, 3), Len(Mid$(MSG, 3)) - 2)) & "##"
            sPosPistola = "UPD1_BF" & Format$(Asc(UCase$(Left$(MSG, 1))) - 64, "110")
            MSG = "0" & Right$(Left$(MSG, 2), 1) & "|" & sIDParcel & "|" & String(33, " ") & "|" & String$(33 - Len(sBarcodeHANG), " ") & sBarcodeHANG & "|" & sPosPistola
        End If
        
        
        sBarcodeHANG = Mid$(MSG, 57, 18)
        sBarcodeSCANNER = Right$(Mid$(MSG, 10, 32), 18)
        If Trim$(sBarcodeHANG) = vbNullString Then
            gTotali.Pacchi_Spuntati_Dal_FSC = gTotali.Pacchi_Spuntati_Dal_FSC + 1
            sBarcode = sBarcodeSCANNER
            SalvaSUDB = False
        Else
            'se sono attive alcune pistole per la stampa delle etichette invio
            'il BARCODE al PGM che etichetta
            sBarcode = Trim$(sBarcodeHANG)
            sPosPistola = Mid$(MSG, 78)
            sIDParcel = Mid$(MSG, 4, 5)
            SalvaSUDB = True
            If Trim$(sBarcodeSCANNER) = vbNullString Then
            
            
                gTotali.Pacchi_Spuntati_a_Mano = gTotali.Pacchi_Spuntati_a_Mano + 1
                
                If gStampaEtichetta = 1 Then
                    bTestCRC = Not PistolaAttivaComeEtichetta(sPosPistola, idCom)
                    
                    If Not bTestCRC Then
                        bCercaInPiano = False
                        'Se primo Messaggio quindi non presente BARCODE letto da scanner
                        'Reperisco flg quadratura e flg stampa
                        DatiBarCode sBarcode, Barcode
                        If Barcode.DaStampare = DaStampare Then
                            'invio il barcode al PGM di etichettazione
                            InvioBARCODEAlEtichetta Barcode, sPosPistola, sIDParcel, idCom
                            DoEvents
                        ElseIf EUnBarcodeBartolini(sBarcode) Then
                                'BARCODE bartolini
                                bTestCRC = False
                                bCercaInPiano = True
                            ElseIf Barcode.DaStampare = daNONStampare Or Barcode.DaStampare = NonTrovato Then
                                'BARCODE Non bartolini da non STAMPARE
                                InviaAllaBaia "996", "996", "99", 2, Mid$(MSG, 4, 5), SalvaSUDB, sBarcode
                        End If
                    End If
                End If
            
                gTotali.Pacchi = gTotali.Pacchi_Spuntati_a_Mano + gTotali.Pacchi_Spuntati_Dal_FSC
            Else
                'Ricevuto secondo MESSAGGIO da FSC
                bTestCRC = False
                bCercaInPiano = False
                
                sDestinazione = GetDestinazionePacco(sIDParcel)
                If sDestinazione = vbNullString Then
                    InviaAllaBaia "998", "998", "99", 2, sIDParcel, False, sBarcode
                Else
                    InvioMsg "", "", "", False, sDestinazione
                End If
            End If
        
        End If
        
        If bTestCRC Then
            'verifico la validit‡ del BARCODE da SCANNER o pistola non collegata all'etichetta
            bCercaInPiano = ControlloCRCBarCode(sBarcode)
            If Not bCercaInPiano Then
                InviaAllaBaia "998", "998", "99", 2, Mid$(MSG, 4, 5), SalvaSUDB, sBarcode
            Else
                'test quadratura solo da SCANNER
                DatiBarCode sBarcode, Barcode
            End If
        End If
        
        If bCercaInPiano Then
            If Barcode.Quadratura = Squadrato Then
                gTotali.ColliSquadrati = gTotali.ColliSquadrati + 1
                InviaAllaBaia "997", "997", "99", 2, Mid$(MSG, 4, 5), SalvaSUDB, sBarcode
            Else
                If Barcode.Quadratura = Quadrato Then gTotali.ColliQuadrati = gTotali.ColliQuadrati + 1
                InviaAllaBaia Left$(sBarcode, 3), Mid$(sBarcode, 4, 3), Mid$(sBarcode, 16, 2), 1, Mid$(MSG, 4, 5), SalvaSUDB, sBarcode
            End If
        End If
        
    
    ElseIf Left$(MSG, 2) = "30" Then
            DoEvents
            'Messaggio dal PGM stampa etichetta per trovare la destinazione del PARCEL
            '30
            'MS|LNP|LNA|ZO|Id   |BarCode           |Quadratura
            '30|001|001|02|12345|123456789012345678|1
            SalvaSUDB = True
            sBarcode = Left$(Mid$(MSG, 21), Len(Mid$(MSG, 21)) - 2)
            sIDParcel = Mid$(MSG, 15, 5)
            Select Case Right$(MSG, 1)
                Case NonTrovato
                    InviaAllaBaia "996", "996", "99", 2, sIDParcel, SalvaSUDB, sBarcode
                    gTotali.PalletCardNONStampate = gTotali.PalletCardNONStampate + 1
                Case Squadrato
                    gTotali.PalletCardStampate = gTotali.PalletCardStampate + 1
                    gTotali.ColliSquadrati = gTotali.ColliSquadrati + 1
                    InviaAllaBaia "997", "997", "99", 2, sIDParcel, SalvaSUDB, sBarcode
                Case Quadrato
                    gTotali.PalletCardStampate = gTotali.PalletCardStampate + 1
                    gTotali.ColliQuadrati = gTotali.ColliQuadrati + 1
                    InviaAllaBaia Mid$(MSG, 4, 3), Mid$(MSG, 8, 3), Mid$(MSG, 12, 2), 1, sIDParcel, SalvaSUDB, sBarcode
            End Select
        
        Else
            DoEvents
            If Left$(MSG, 2) = "20" Then
                ScriviLog TPL_INFORAMITIVO, "Attenzione sono stati invertiti i cavi: (Request/message con Sort/Message)", "GestioneMSG"
                FrmMonitoraggio.Label1 = "Il cavo del RQST/MSG non Ë nella COM giusta"
            End If
            
            ScriviLog TPL_ERRORE, MSG, "GestioneMSG"
            gTotali.Messaggi_RQST_Errati = gTotali.Messaggi_RQST_Errati + 1
    End If
End Sub

Public Function LeggiChiave(Chiave As String) As Long
    LeggiChiave = -1
    On Error Resume Next
    LeggiChiave = colKey(Chiave)
End Function

Function LeggoFileDaAS400oPC(sCodicePiano As String, sTipoLavorazione As String, sNFoglioDiViaggio As String) As Boolean
    Dim rdoRs1 As rdoResultset
    Dim rdoRs2 As rdoResultset
    Dim dbc As rdoConnection
    Dim i As Long
    

    Screen.MousePointer = vbHourglass
    DoEvents
    On Error GoTo Errore
    LeggoFileDaAS400oPC = False
    Set dbc = rdoEnv.OpenConnection("CollegamentoFileTesto")
    
    If gConnesso Then
        'FLTR200F
        FrmInstrada.LblStatoTrasferimento = "Importazione Fogli di Viaggio " & gNomeFileFogliDiViaggio & " da AS400 in corso ..."
        DoEvents
        Set rdoRs1 = rdoCntToAS400.OpenResultset("Select * from " & gNomeFileFogliDiViaggio)
        
        FrmInstrada.LblStatoTrasferimento = "Generazione Fogli di Viaggio " & gNomeFileFogliDiViaggio & " su PC in corso ..."
        DoEvents
        GeneraFileTesto rdoRs1, App.Path & "\" & gNomeFileFogliDiViaggio & ".TXT", gNomeFileFogliDiViaggio
        'FLTR300F
        FrmInstrada.LblStatoTrasferimento = "Importazione Piani Smistamento " & gNomeFilePianiSmistamento & " da AS400 in corso ..."
        DoEvents
        Set rdoRs2 = rdoCntToAS400.OpenResultset("Select * from " & gNomeFilePianiSmistamento)
        FrmInstrada.LblStatoTrasferimento = "Generazione Piani Smistamento " & gNomeFilePianiSmistamento & " su PC in corso ..."
        DoEvents
        GeneraFileTesto rdoRs2, App.Path & "\" & gNomeFilePianiSmistamento & ".TXT", gNomeFilePianiSmistamento
        
        If Not rdoRs1 Is Nothing Then
            rdoRs1.Close
            Set rdoRs1 = Nothing
        End If
        If Not rdoRs2 Is Nothing Then
            rdoRs2.Close
            Set rdoRs2 = Nothing
        End If
        
    End If
       
    'Da file testo seleziono Tipo Lavorazione e Codice Piano per verificarne la corettezza
    If GeneraSchemaINI("FLTR300F.TXT", gNomeFilePianiSmistamento & ".txt") Then
        Set rdoRs1 = dbc.OpenResultset("select * from " & gNomeFilePianiSmistamento & ".txt where F1=" & "'" & sTipoLavorazione & "'" & " and " & "F2=" & "'" & sCodicePiano & "'" & " ORDER BY F3,F4,F5,F7 DESC")
    End If
    'Da file testo seleziono Tipo Lavorazione e N∞ Foglio viaggio per verificarne la corettezza
    If GeneraSchemaINI("FLTR200F.TXT", gNomeFileFogliDiViaggio & ".txt") Then
        Set rdoRs2 = dbc.OpenResultset("select * from " & gNomeFileFogliDiViaggio & ".txt where F1=" & "'" & sTipoLavorazione & "'" & " and " & "F2=" & "'" & sNFoglioDiViaggio & "'")
    End If
    
    If Not rdoRs1 Is Nothing And Not rdoRs2 Is Nothing Then
        LeggoFileDaAS400oPC = (Not rdoRs1.EOF) And (Not rdoRs2.EOF)
        If rdoRs1.EOF Then MsgBox "Tipo Lavorazione e Codice Piano errato", vbCritical
        If rdoRs2.EOF Then MsgBox "Tipo Lavorazione e Numero Foglio di Viaggio errato", vbCritical
    End If
    
    If LeggoFileDaAS400oPC Then

        gFLTR200F.TipoLavorazione = rdoRs2.rdoColumns(0)
        gFLTR200F.CodceFoglioDiViaggio = rdoRs2.rdoColumns(1)
        
        Set colKey = New Collection
        i = 0
        FrmInstrada.LblStatoTrasferimento = "Inizializzazione sistema in corso ..."
        DoEvents
        Do While Not rdoRs1.EOF
            ReDim Preserve gFLTR300F(i)
            gFLTR300F(i).TipoLavorazione = IIf(IsNull(rdoRs1.rdoColumns(0)), vbNullString, rdoRs1.rdoColumns(0))
            gFLTR300F(i).CodicePiano = IIf(IsNull(rdoRs1.rdoColumns(1)), vbNullString, rdoRs1.rdoColumns(1))
            gFLTR300F(i).LineaDiPartenza = IIf(IsNull(rdoRs1.rdoColumns(2)), vbNullString, rdoRs1.rdoColumns(2))
            gFLTR300F(i).LineaDiArrivo = IIf(IsNull(rdoRs1.rdoColumns(3)), vbNullString, rdoRs1.rdoColumns(3))
            gFLTR300F(i).Zona = IIf(IsNull(rdoRs1.rdoColumns(4)), vbNullString, rdoRs1.rdoColumns(4))
            gFLTR300F(i).CodiceBaia = IIf(IsNull(rdoRs1.rdoColumns(5)), vbNullString, rdoRs1.rdoColumns(5))
            gFLTR300F(i).Quantit‡Baie = IIf(IsNull(rdoRs1.rdoColumns(6)), vbNullString, rdoRs1.rdoColumns(6))
            gFLTR300F(i).Descrizione = IIf(IsNull(rdoRs1.rdoColumns(7)), vbNullString, rdoRs1.rdoColumns(7))
            gFLTR300F(i).Key = gFLTR300F(i).LineaDiPartenza & gFLTR300F(i).LineaDiArrivo & gFLTR300F(i).Zona
            If LeggiChiave(gFLTR300F(i).Key) = -1 Then
                colKey.Add i, gFLTR300F(i).Key
            End If

            i = i + 1
            rdoRs1.MoveNext
        Loop
    End If
    
    If Not rdoRs1 Is Nothing Then
        rdoRs1.Close
        Set rdoRs1 = Nothing
    End If
    If Not rdoRs2 Is Nothing Then
        rdoRs2.Close
        Set rdoRs2 = Nothing
    End If
    If Not dbc Is Nothing Then
        dbc.Close
        Set dbc = Nothing
    End If
        
    FrmInstrada.LblStatoTrasferimento = ""
    Screen.MousePointer = vbDefault
    DoEvents
    Exit Function
    
Errore:
    Screen.MousePointer = vbDefault
    If Not rdoEngine.rdoErrors Is Nothing Then
        If rdoEngine.rdoErrors.Item(0).Number = 0 Then
            MsgBox rdoEngine.rdoErrors.Item(0).Description & " DSN=CollegamentoFileTesto", vbCritical + vbOKOnly
            ScriviLog TPL_ERRORE, rdoEngine.rdoErrors.Item(0).Description & " DSN=CollegamentoFileTesto", "LeggoFileDaAS400oPC"
            Exit Function
        End If
        ScriviLog TPL_ERRORE, rdoEngine.rdoErrors.Item(0).Description, "LeggoFileDaAS400oPC"
    Else
        ScriviLog TPL_ERRORE, Error$, "LeggoFileDaAS400oPC"
    End If
   Resume Next
End Function


Function GeneraFileTesto(rs As rdoResultset, sNomeFile As String, sTable As String) As Boolean
    Dim c() As Long
    Dim Col As Long
    Dim sString As String
    Dim n  As Long
    Dim i  As Long
    Dim nFile As Long
    
    GeneraFileTesto = False
    If rs.EOF Then Exit Function
    
    n = rs.rdoColumns.Count
    ReDim c(n)
    For Col = 0 To n - 1
        If rs.rdoColumns.Item(Col).Type = rdTypeCHAR Then
            c(Col) = rdoCntToAS400.rdoTables(sTable).rdoColumns.Item(Col).Size
        Else
            c(Col) = rdoCntToAS400.rdoTables(sTable).rdoColumns.Item(Col).Size - 2
        End If
    Next

    nFile = FreeFile
    
    On Error Resume Next
    Kill sNomeFile
    On Error GoTo 0
    Open sNomeFile For Append As #nFile
    
    Do While Not rs.EOF
        sString = vbNullString
        For i = 0 To n - 1
            If rs.rdoColumns.Item(i).Type = rdTypeCHAR Then
                sString = sString & Left$(rs.rdoColumns.Item(i) & Space$(c(i)), c(i))
            Else
                sString = sString & Format$(rs.rdoColumns.Item(i), String(c(i), "0"))
            End If
        Next
        Print #nFile, sString
        DoEvents
        rs.MoveNext
    Loop
        
    Close #nFile
    GeneraFileTesto = True
End Function

Function GeneraSchemaINI(sNomeFile As String, Optional ByVal sFile As String, Optional ByVal sPathFile As String) As Boolean
    Dim sValue As String
    Dim sEntryValue As String
    Dim sEntry As String
    Dim ret As Long
    Dim iOld As Long
    Dim iNew As Long

    sValue = Space$(1024)
    ret = GetPrivateProfileSection(sNomeFile, sValue, Len(sValue), App.Path & "\" & SCHEMAINI)
    
    If ret <> 0 Then
        On Error Resume Next
        Kill App.Path & "\Schema.ini"
        On Error GoTo 0
        
        If sPathFile = vbNullString Then
            sPathFile = App.Path & "\Schema.ini"
        Else
            sPathFile = sPathFile & "\Schema.ini"
        End If

        If sFile = vbNullString Then
            sEntry = sNomeFile
        Else
            sEntry = sFile
        End If

        ret = WritePrivateProfileSection(sEntry, sValue, sPathFile)
        GeneraSchemaINI = True
    Else
        ScriviLog TPL_ERRORE, "Non trovato Entry o " & App.Path & "\" & SCHEMAINI, "GeneraSchemaINI"
        GeneraSchemaINI = False
    End If
End Function


Sub InviaAllaBaia(ByVal LNP As String, ByVal LNA As String, ByVal ZO As String, ByVal err As String, ByVal sNumProgPOLISORTER As String, ByVal SalvaSUDB As Boolean, ByVal sBarcode As String)
    Dim Esci As Boolean, flgInviato As Boolean
    Dim i As Long, ii As Long, iii As Long
    Dim sChiave As String
    Dim msg20 As String
    Dim ret As Long

    'se non vengono gestiti gli instradamenti devo creare solo il MSG 20 per generare le spunte
    'da inviare all'AS400
    
    If gAttivaInstradamenti = 1 Then
    
        For iii = 1 To 4
            Select Case iii
                Case 1: sChiave = LNP & LNA & ZO
                Case 2: sChiave = "999" & LNA & ZO
                Case 3: sChiave = "999" & LNA & "99"
                Case 4: sChiave = "99999999"
            End Select
            i = LeggiChiave(sChiave)
            If i >= 0 Then
                If gFLTR300F(i).Quantit‡Baie > 1 Then
                    For ii = i To i + gFLTR300F(i).Quantit‡Baie - 1
                        If gFLTR300F(ii).Key = sChiave And Not gFLTR300F(ii).Inviato Then
                            InvioMsg gFLTR300F(ii).CodiceBaia, err, sNumProgPOLISORTER, SalvaSUDB
                            gFLTR300F(ii).Inviato = True
                            If iii <> 4 Then
                                If err <> 2 Then
                                    gTotali.Pacchi_Inviati_OK = gTotali.Pacchi_Inviati_OK + 1
                                Else
                                    gTotali.Pacchi_Inviati_Non_Letti = gTotali.Pacchi_Inviati_Non_Letti + 1
                                End If
                            Else
                                gTotali.Pacchi_Inviati_Non_Trovati = gTotali.Pacchi_Inviati_Non_Trovati + 1
                            End If
                            Exit Sub
                        End If
                    Next
                    
                    InvioMsg gFLTR300F(i).CodiceBaia, err, sNumProgPOLISORTER, SalvaSUDB
                    For ii = i To i + gFLTR300F(i).Quantit‡Baie - 1
                        gFLTR300F(ii).Inviato = False
                    Next
                    gFLTR300F(i).Inviato = True
                    If iii <> 4 Then
                        If err <> 2 Then
                            gTotali.Pacchi_Inviati_OK = gTotali.Pacchi_Inviati_OK + 1
                        Else
                            gTotali.Pacchi_Inviati_Non_Letti = gTotali.Pacchi_Inviati_Non_Letti + 1
                        End If
                    Else
                        gTotali.Pacchi_Inviati_Non_Trovati = gTotali.Pacchi_Inviati_Non_Trovati + 1
                    End If
                    Exit Sub
                Else
                    InvioMsg gFLTR300F(i).CodiceBaia, err, sNumProgPOLISORTER, SalvaSUDB
                    If iii <> 4 Then
                        If err <> 2 Then
                            gTotali.Pacchi_Inviati_OK = gTotali.Pacchi_Inviati_OK + 1
                        Else
                            gTotali.Pacchi_Inviati_Non_Letti = gTotali.Pacchi_Inviati_Non_Letti + 1
                        End If
                    Else
                        gTotali.Pacchi_Inviati_Non_Trovati = gTotali.Pacchi_Inviati_Non_Trovati + 1
                    End If
                    Exit Sub
                End If
            End If
        Next
    Else
        If err = "2" Then
            'Errore
            sBarcode = String$(18, "?")
        End If
        sBarcode = "000" & String$(28 - Len(sBarcode), " ") & sBarcode & "##"
        msg20 = "20|" & sNumProgPOLISORTER & "|" & String(6, "0") & "|" & String(3, "0") & "|" & String(1, "0") & "|" & String(3, "0") & "|" & String(1, "0") & "|" & String(33, "0") & _
                "|" & sBarcode & "|" & String(33, "0") & "|" & String(33, "0") & "|" & String(14, "0") & "|" & String(1, "0") & "|" & String(14, "0") & "|" & String(1, "0") & "|00"
        ret = SendMessage(ByVal gHndTxtWaitSortMsg20, ByVal WM_SETTEXT, ByVal 0&, ByVal msg20)
    End If
End Sub

Function TrasferisciFilediSpunta(sFile As String) As Boolean
    Dim dbc As rdoConnection
    Dim rs As rdoResultset
    Dim qry As rdoQuery
    Dim lParam As Long
    Dim Chiamata   As String
    Dim fine   As String
    

    TrasferisciFilediSpunta = False
    frmCopyFile.Label1 = "Trasferimento di " & sFile & " in corso ...."
    Screen.MousePointer = vbHourglass
    DoEvents
    If GeneraSchemaINI("SPUNTA", sFile, gPathWrkFileSpunta) Then
        On Error GoTo Errore
        Set dbc = rdoEnv.OpenConnection("TrasferimentoFileDiSpuntaWRK")
        Set rs = dbc.OpenResultset("select * from " & sFile, rdOpenForwardOnly)
        
        If Not rs.EOF Then
            Set qry = rdoCntToAS400.CreateQuery("", SQLstring:="insert into " & gNomeFileSpunta & " values(" & "?" & ")")
        End If
        Do While Not rs.EOF
            With qry
                For lParam = 0 To .rdoParameters.Count - 1
                    If .rdoParameters(lParam).Type = rdTypeNUMERIC Then
                        .rdoParameters(lParam) = Val("0" & rs.rdoColumns(lParam))
                    Else
                        .rdoParameters(lParam) = "" & rs.rdoColumns(lParam)
                    End If
                Next
                .Execute
            End With
            rs.MoveNext
        Loop
        
        Screen.MousePointer = vbDefault
        
        If Not rs Is Nothing Then
            rs.Close
            Set rs = Nothing
        End If
        
        If Not qry Is Nothing Then
            qry.Close
            Set qry = Nothing
        End If
        
        If Not dbc Is Nothing Then
            dbc.Close
            Set dbc = Nothing
        End If
       
        TrasferisciFilediSpunta = True
    End If
    
    Exit Function
Errore:
    
    Screen.MousePointer = vbDefault
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    If Not qry Is Nothing Then
        qry.Close
        Set qry = Nothing
    End If
    
    If Not dbc Is Nothing Then
        dbc.Close
        Set dbc = Nothing
    End If
    
    
    If rdoErrors.Count > 0 Then
        If rdoErrors.Item(0).Number = 0 Then
            MsgBox rdoEngine.rdoErrors.Item(0).Description & " DSN=TrasferimentoFileDiSpuntaWRK", vbCritical + vbOKOnly
        End If
        ScriviLog TPL_ERRORE, rdoEngine.rdoErrors.Item(0).Description, "TrasferisciFileDiSpunta"
    Else
        ScriviLog TPL_ERRORE, Error$, "TrasferisciFileDiSpunta"
    End If

End Function



Function InvioLogAllAs400(sNomeFile As String) As Boolean
    Dim dbc As rdoConnection
    Dim rs As rdoResultset
    
    Screen.MousePointer = vbHourglass
    InvioLogAllAs400 = False
    DoEvents
    
    On Error GoTo Errore
    
    Set dbc = rdoEnv.OpenConnection("TrasferimentoFileDiSpuntaWRK")
    
    'Creao connessione su file POSITROVA.LOG
    If GeneraSchemaINI("FILELOG", sNomeFile) Then
        Set rs = dbc.OpenResultset("select * from " & sNomeFile)
    End If
       
    Do While Not rs.EOF
        If vbNullString & rs.rdoColumns(0) <> vbNullString Then
            rdoCntToAS400.Execute "Insert into FLTRLOG  Values('" & RaddoppiaApice(rs.rdoColumns(0)) & "')"
        End If
        rs.MoveNext
    Loop
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    If Not dbc Is Nothing Then
        dbc.Close
        Set dbc = Nothing
    End If
        
    Screen.MousePointer = vbDefault
    DoEvents
    InvioLogAllAs400 = True
    Exit Function
    
Errore:
    
    InvioLogAllAs400 = False
    Screen.MousePointer = vbDefault
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    If Not rdoEngine.rdoErrors Is Nothing Then
        If rdoEngine.rdoErrors.Item(0).Number = 0 Then
            MsgBox rdoEngine.rdoErrors.Item(0).Description & " DSN=CollegamentoFileTesto", vbCritical + vbOKOnly
            ScriviLog TPL_ERRORE, rdoEngine.rdoErrors.Item(0).Description & " DSN=CollegamentoFileTesto", "InvioLogPCtoAS400"
        End If
        ScriviLog TPL_ERRORE, rdoEngine.rdoErrors.Item(0).Description, "InvioLogPCtoAS400"
    Else
        ScriviLog TPL_ERRORE, Error$, "InvioLogPCtoAS400"
    End If
    
    If Not dbc Is Nothing Then
        dbc.Close
        Set dbc = Nothing
    End If
    Exit Function

End Function

