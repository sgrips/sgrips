Attribute VB_Name = "ModAnagrafica"
Option Explicit



Public Function GetRetribuzione(CodComune As String, AnnoRiferimento As String, TipoPersonale As Long, Livello As String) As Double
    Dim rs As Recordset
    
    'Dato in codice del comune , la tipologia del personale e il suo livello
    'reperisce la retribuzione annua
    GetRetribuzione = 0
    Set rs = gDBC.Execute("Select RetribuzioneAnnuale from RetribuzioneDelPersonale where CodComune='" & CodComune & "' and AnnoRiferimento='" & AnnoRiferimento & "' and CodTipoPersonale=" & TipoPersonale & " and CodLivelloPersonale='" & Livello & "'")
    
    If Not rs.EOF Then GetRetribuzione = rs("RetribuzioneAnnuale")
End Function

Sub MettivaloreInPKEY(pKey As String, ByVal Key As String, ByVal value As Variant, ByVal TipoDato As ADODB.DataTypeEnum)
    'Inserisce un valore nella PKEY
    Dim i As Integer
    Dim iUguale As Integer
    
    i = InStr(1, pKey, Key)
    
    If i > 0 Then
         i = InStr(i + 1, pKey, "=")
         iUguale = i
         If TipoDato = adChar Or TipoDato = adVarChar Or TipoDato = adVarWChar Then
            pKey = Left$(pKey, iUguale) & "'" & value & "'" & Mid$(pKey, iUguale + 1)
         ElseIf TipoDato = adDBTimeStamp Or TipoDato = adDate Then
                pKey = Left$(pKey, iUguale) & "#" & Format$(value, "YYYY/MM/DD") & "#" & Mid$(pKey, iUguale + 1)
             ElseIf TipoDato = adBoolean Then
                    pKey = Left$(pKey, iUguale) & IIf(value = False, "false", "true") & Mid$(pKey, iUguale + 1)
             Else
                    pKey = Left$(pKey, iUguale) & value & Mid$(pKey, iUguale + 1)
            End If
    End If
    
End Sub


Function GetPkeyForWhereCondition(ByVal s As String) As String
    'Crea la where condition dalla PKEY
    Dim i As Integer
    Dim ii As Integer
    
    
    'cancello i campi di PKEY non valorizzati
    For i = 2 To Len(s)
        If Mid$(s, i, 1) = SEPARATORE Then
            If Mid$(s, i - 1, 1) = "=" Then
            'Togli Campo
                ii = InStrRev(s, SEPARATORE, i - 1)
                s = Left$(s, ii - 1) & Mid$(s, i)
            End If
        End If
    Next
    
    s = Mid$(s, 2, Len(s) - 2)
    s = Replace(s, "&", " AND ")
    
    GetPkeyForWhereCondition = s

    
End Function


Function getCampiDapKeyPerInsert(ByVal s As String) As String
    Dim i As Integer
    Dim ii As Integer
    
    i = InStr(1, s, SEPARATORE)
    ii = InStr(i, s, "=")
    Do While i > 0 And ii > 0
        getCampiDapKeyPerInsert = getCampiDapKeyPerInsert + Mid$(s, i + 1, ii - i - 1) & ","
        i = InStr(i + 1, s, SEPARATORE)
        ii = InStr(i, s, "=")
    Loop
    
    getCampiDapKeyPerInsert = Left$(getCampiDapKeyPerInsert, Len(getCampiDapKeyPerInsert) - 1)
End Function


Function getValoriDapKeyPerInsert(ByVal s As String) As String
    Dim i As Integer
    Dim ii As Integer
    
    i = InStr(1, s, "=")
    
    Do While i > 0
        ii = InStr(i, s, SEPARATORE)
        getValoriDapKeyPerInsert = getValoriDapKeyPerInsert + Mid$(s, i + 1, ii - i - 1) & ","
        i = InStr(i + 1, s, "=")
    Loop
    
    getValoriDapKeyPerInsert = Left$(getValoriDapKeyPerInsert, Len(getValoriDapKeyPerInsert) - 1)
End Function


Function CercaKeyInListItems(lvwItem As ListItems, Key As String) As Integer
    'cerca la key nella collezzione dei list item
    Dim i As Integer
    CercaKeyInListItems = 0
    For i = 1 To lvwItem.Count
        If lvwItem(i).Key = Key Then
            CercaKeyInListItems = i
            Exit For
        End If
    Next
End Function


Function ScritturaSuDB(lvw As ListView, tabella As String) As Boolean
    Dim i As Integer
    Dim Item As ListItem
    Dim rs As ADODB.Recordset
    Dim sWhere As String
    Dim sCampiEvalori As String
    'Scrivo sul MDB le modifiche al dati
    
    On Error GoTo ScritturaSuDBERR
    ScritturaSuDB = False
    For i = 1 To lvw.ListItems.Count
        Set Item = lvw.ListItems(i)
        
        sWhere = " Where " & GetPkeyForWhereCondition(Item.Key)
        Set rs = gDBC.Execute("Select * from " & tabella & sWhere)
        'Se il Record Esiste faccio l'UPDATE altrimenti lo inserisco
        

        If Not rs.EOF Then
            If Item.Ghosted Then
            'se Ghosted cancello record da MDB
                If tabella = "AnagraficaComunale" Then gDBC.Execute ("Delete from ParametriSistema where AnnoRiferimento=" & GetValueInTag("AnnoRiferimento", Item.Key) & " and CODISTAT=" & GetValueInTag("CodComune", Item.Key))
                gDBC.Execute ("Delete from " & tabella & sWhere)
            Else
            'UPDATE
                sCampiEvalori = GetCampiEValoriDaLVWItem(lvw, i, False)
                gDBC.Execute ("Update " & tabella & " set " & sCampiEvalori & sWhere)
            End If
        Else
            'INSERT
            If Not Item.Ghosted Then
            End If
        End If
    Next
    
    ScritturaSuDB = True
    Exit Function
ScritturaSuDBERR:
    MsgBox Error, vbCritical, App.Title
End Function


Function GetCampiEValoriDaLVWItem(lvw As ListView, IdItem, TipoInserimento As Boolean) As String
    'Crea una stringa con i campi per Inserimento o Update nel database
    Dim i As Integer
    Dim s As String
    Dim sVal As String
    Dim TipoFld As Long
    Dim sValore As String
    
    With lvw
        If TipoInserimento Then
            s = "("
            sVal = s
        
        'INSERT
            For i = 1 To .ColumnHeaders.Count
                s = s & .ColumnHeaders(i).Key & ","
                
                If InStr(.ColumnHeaders(i).Key, "_NDB") = 0 Then
                    TipoFld = GetValueInTag(TAG_TIPO_FLD, .ColumnHeaders(i).Tag)
                    If i = 1 Then
                        sValore = .ListItems(IdItem).Text
                    Else
                        sValore = .ListItems(IdItem).ListSubItems(i - 1)
                    End If
                    sVal = sVal & FormatFLD(sValore, TipoFld) & ","
                End If
            Next
            s = Left$(s, Len(s) - 1) & " " & Left$(sVal, Len(sVal) - 1) & ")"
        Else

        'UPDATE
            For i = 1 To .ColumnHeaders.Count
                If InStr(.ColumnHeaders(i).Key, "_NDB") = 0 Then
                    TipoFld = GetValueInTag(TAG_TIPO_FLD, .ColumnHeaders(i).Tag)
                        
                    
                    If i = 1 Then
                        sValore = .ListItems(IdItem).Text
                    Else
                        sValore = .ListItems(IdItem).ListSubItems(i - 1)
                    End If
                    
                    s = s & .ColumnHeaders(i).Key & "=" & FormatFLD(sValore, TipoFld) & ","
                End If

            Next
            

            s = Left$(s, Len(s) - 1)
            
        End If

    GetCampiEValoriDaLVWItem = s
    End With
End Function

Function FormatFLD(ByVal s As String, ByVal Tipo As DataTypeEnum) As String
    'Formatta il campo a seconda del tipo
    Dim i As Integer

    If (s = vbNullString) Then
        FormatFLD = "NULL"
    Else
        Select Case Tipo
            Case adChar, adVarWChar, adVarChar
                FormatFLD = "'" & Replace(s, "'", "''") & "'"
            Case adDate
                FormatFLD = "#" & s & "#"
            Case adBoolean
                s = UCase$(s)
                FormatFLD = IIf(s = "TRUE" Or s = "VERO", "True", "False")
            Case Else
                i = InStr(s, ",")
                If i > 0 Then Mid$(s, i) = "."
                FormatFLD = s
        End Select
    
    End If
    
    
End Function


Sub CaricaInListView(lvw As ListView, Icona As String, tabella As TipoDatoLvwEnum, Optional DatiSimulazioneCompresi As Boolean = False)
    'Carica le righe riguardanti i dati anagrafici del comune
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim sPKey As String
    Dim Item As ListItem
    Dim sKEYLVW As String
    Dim l As Long
    Dim bColNoDB As Boolean
    Dim svalue As String
    Dim oHeader As ColumnHeader
    Dim sTable As String
    Dim s As String
    Dim where As String
    
    
    Set rs = CreateObject("ADODB.Recordset")
    where = gWhereSuComune & gWhereFiltroSuComune
    If where = vbNullString Then
        where = " where "
    Else
        where = where & " and "
    End If
    
    
    Select Case tabella
        Case Is = Lvw_AnagraficaComuni
            where = where & " AnagraficaComunale.flgSimulazione=false "
            rs.Source = "Select CodComune,Comuni.Comune as Comune_NDB,AnnoRiferimento,Regioni.Regione as Regione_NDB,Provincie.Provincia as Provincia_NDB,Comuni.Cap as Cap_NDB,NumAbitantiResidenti , NumNucleoFamiliare, FamigliaCon1Componente, FamigliaCon2Componenti, FamigliaCon3Componenti, FamigliaCon4Componenti, FamigliaCon5Componenti " & _
                ", FamigliaConPiuDi5Componenti, Immigrati, Emmigrati, NumMedioResidenti, NumDomiciliatii, NumMedioTuristi, SuperficieDelTerrComunale, AmbitoComunaleOttimale, PRComunaleOttimale " _
                & "from  ((AnagraficaComunale INNER JOIN Comuni on  AnagraficaComunale.CodComune=Comuni.CodIstat ) INNER JOIN   Provincie on Comuni.Provincia =Provincie.Provincia) INNER JOIN Regioni on Regioni.CodRegione=Provincie.CodRegione " & where & " Order by CodComune,AnnoRiferimento "
                
        Case Is = Lvw_CostiComunali
            where = where & IIf(DatiSimulazioneCompresi, "1=1", " CostoComunale.flgSimulazione=false ")
            rs.Source = "Select CodComune,Comuni.Comune as Comune_NDB,AnnoRiferimento,Regioni.Regione as Regione_NDB,Provincie.Provincia as Provincia_NDB,Comuni.Cap as Cap_NDB" _
                & ",Tot_Servizio , Gettito, CG, CSL, CRT, CTS, AC,CGIND,CRD,CTR_COSTI,CTR_RICAVI,CTR,CON,Y,CGD,CC,CARC,CGG,CCD,CK,AMM,ACC,R,Rn,Kn_Anno_Prec,Investimenti_Programmati,Fn,Tot_CSL,Tot_CRT  ,Tot_CRD  ,Tot_CTS ,Tot_CTR  " _
                & "from ((CostoComunale INNER JOIN Comuni on  CostoComunale.CodComune=Comuni.CodIstat ) INNER JOIN   Provincie on Comuni.Provincia =Provincie.Provincia) INNER JOIN Regioni on Regioni.CodRegione=Provincie.CodRegione " & where & " Order by CodComune,AnnoRiferimento "
        Case Is = Lvw_TipiAmmortamenti
            rs.Source = "Select CodTipoAmmortamento  ,TipiAmmortamento.Descrizione as Descrizione , FamigliaAmmortamenti.Descrizione as Tipologia from  TipiAmmortamento , FamigliaAmmortamenti where  CodFamigliaAmmortamento =FamigliaAmmortamento Order by FamigliaAmmortamenti.Descrizione,TipiAmmortamento.Descrizione"
        Case Is = Lvw_Personale
            rs.Source = "Select CodTipoPersonale  ,Descrizione,CGG as CostoGeneraleGestione from  TipiPersonaleImpiegato Order by Descrizione"
        Case Is = Lvw_DatiTecnici
            where = where & " DatiTecniciComune.flgSimulazione=false "
            rs.Source = "Select CodComune,Comuni.Comune as Comune_NDB,AnnoRiferimento,Regioni.Regione as Regione_NDB,Provincie.Provincia as Provincia_NDB,Comuni.Cap as Cap_NDB, " _
            & "QRifiutiTotali,QRifiutiRaccoltaDifferenziata,NRuoliUtenzaDomestica,NRuoliUtenzaAttivita,TotSuperficieUtenzaDomestica,TotSuperficieUtenzaAttivita,TotSuperficieNonIscritti,TotGettitoDaUtenzaDomestica,TotGettitoDaUtenzaAttivita , QRifiutiCategorieAttivita,NRuoliSecondaCasa,Id_NumCompSecondaCasa " _
            & "from  ((DatiTecniciComune INNER JOIN Comuni on  DatiTecniciComune.CodComune=Comuni.CodIstat ) INNER JOIN   Provincie on Comuni.Provincia =Provincie.Provincia) INNER JOIN Regioni on Regioni.CodRegione=Provincie.CodRegione " & where & " Order by CodComune,AnnoRiferimento "
    
        Case Is = Lvw_AnagraficaRaccolta
            where = where & " AnagraficaRaccolta.flgSimulazione=false "
            rs.Source = "Select CodComune,Comuni.Comune as Comune_NDB,AnnoRiferimento,Regioni.Regione as Regione_NDB,Provincie.Provincia as Provincia_NDB,Comuni.Cap as Cap_NDB, " _
                & "ModalitaCalcoloPV, ModalitaCalcoloPF, UnitaDiMisuraPV " _
                & "from  ((AnagraficaRaccolta INNER JOIN Comuni on  AnagraficaRaccolta.CodComune=Comuni.CodIstat ) INNER JOIN   Provincie on Comuni.Provincia =Provincie.Provincia) INNER JOIN Regioni on Regioni.CodRegione=Provincie.CodRegione " & where & " Order by CodComune,AnnoRiferimento "
                    
    End Select
    
    
    Set rs.ActiveConnection = gDBC
    rs.Open

    'creo testata ListView
    Screen.MousePointer = vbHourglass
    lvw.Visible = False
    lvw.ColumnHeaders.Clear
    lvw.ListItems.Clear
    

    For i = 0 To rs.fields.Count - 1
        With rs.fields(i)
            bColNoDB = InStr(1, .Name, "_NDB") > 0
            
            If bColNoDB Then
                svalue = Left$(.Name, Len(.Name) - 4)
            Else
                svalue = .Name
            End If
                
            Set oHeader = lvw.ColumnHeaders.Add(, .Name, svalue)
            
            oHeader.Tag = SetValueInTag(TAG_TIPO_FLD, .Type, oHeader.Tag)
            oHeader.Tag = SetValueInTag(TAG_SIZE_FLD, .DefinedSize, oHeader.Tag)
        End With
    Next
    

    

    i = InStr(1, UCase$(rs.Source), "FROM ")
    sTable = Mid$(rs.Source, i + Len("FROM "))
    For i = 1 To Len(sTable)
        If Mid$(sTable, i, 1) <> " " And Mid$(sTable, i, 1) <> "(" Then Exit For
    Next
    sTable = Mid$(sTable, i)
    
    For i = 1 To Len(sTable)
        If Mid$(sTable, i, 1) = " " Or Mid$(sTable, i, 1) = "(" Or Mid$(sTable, i, 1) = ")" Then Exit For
    Next
    
    sTable = Left$(sTable, i - 1)
    
    sPKey = CreaLaPkey(sTable)
    
    Do While Not rs.EOF
        sKEYLVW = sPKey
        For i = 0 To rs.fields.Count - 1
            If InStr(sPKey, SEPARATORE & rs.fields(i).Name) Then
                MettivaloreInPKEY sKEYLVW, SEPARATORE & rs.fields(i).Name, rs.fields(i).value, rs.fields(i).Type
            End If
        Next
        

        For i = 0 To rs.fields.Count - 1
            If i = 0 Then
                Set Item = lvw.ListItems.Add(, sKEYLVW, rs.fields(i).value, , Icona)
            Else
                Item.ListSubItems.Add , rs.fields(i).Name, vbNullString & rs.fields(i).value
            End If
            
        Next
        
        rs.MoveNext
    Loop
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    Set rs = Nothing
    
    
    AllargaLeColonneLVW lvw, False
    
    Screen.MousePointer = vbDefault
End Sub


Sub AllargaLeColonneLVW(lvw As ListView, AllNumCarAgg As Boolean, ParamArray NumCarAggiunti() As Variant)
    Dim i As Long
    Dim l As Long
    Dim Item  As ListItem
    Dim NumCar As Integer
    
    
    For i = 0 To lvw.ColumnHeaders.Count - 1
        'Aggiungo una riga con i nomi dei titoli per allargare le colonne per
        'incaso che non vi siano record da visualizzare
        
        On Error Resume Next
        NumCar = 0
        NumCar = Abs(NumCarAggiunti(i))

        If AllNumCarAgg Then NumCar = Abs(NumCarAggiunti(0))
        If i = 0 Then
            Set Item = lvw.ListItems.Add(, "TESTATACOLONNE", lvw.ColumnHeaders(i + 1).Text & String(NumCar, "W"))
        Else
            Item.SubItems(i) = lvw.ColumnHeaders(i + 1).Text & String(NumCar, "W")
        End If
    
    Next

    For i = 0 To lvw.ColumnHeaders.Count - 1
        If lvw.ColumnHeaders(i + 1).Width <> 0 And Sgn(NumCarAggiunti(i)) = 1 Then
            l = SendMessage(lvw.hwnd, 4126, i, -1)
        End If
    Next
    
    
    'Rimuovo la riga con le intestazioni
    lvw.ListItems.Remove ("TESTATACOLONNE")
    lvw.Visible = True
End Sub



Function CreaLaPkey(ByVal sTable As String) As String
    Dim i As Integer
    Dim iSelect As Integer
    'crea la pKey
    'Reperisco la PKEY
    iSelect = -1
    With gADOX.Tables(sTable)
        For i = 0 To .Indexes.Count - 1
            If .Indexes(i).PrimaryKey Then iSelect = i: Exit For
        Next
        
        If iSelect <> -1 Then
            CreaLaPkey = vbNullString
            For i = 0 To .Indexes(iSelect).Columns.Count - 1
                CreaLaPkey = CreaLaPkey & SEPARATORE & .Indexes(iSelect).Columns(i) & "="
            Next
        End If
    End With
    
    CreaLaPkey = CreaLaPkey & SEPARATORE
End Function



Function getDesComune(codIstat As String)
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("Select * from Comuni where CodIstat='" & codIstat & "'")
    getDesComune = rs("Comune")
    Set rs = Nothing

End Function

Function getProvinciaComune(codIstat As String)
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("Select * from Comuni where CodIstat='" & codIstat & "'")
    getProvinciaComune = rs("Provincia")
    Set rs = Nothing

End Function

Function ConvertiUnitaDiMisura(ByVal CodiceCer As String, UM1 As String, UM2 As String, Valore As Double) As Double
    'UM1 unità di misura
    'UM2 unità di misura
    'Calcola il valore in un altra unità di misura
    
    Dim Cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim CategoriaAttivita As String
    Dim SottoCategoriaAttivita As String
    
    Set Cmd = CreateObject("ADODB.Command")
    
    CategoriaAttivita = "'" & Left$(CodiceCer, 2) & "'"
    SottoCategoriaAttivita = "'" & Mid$(CodiceCer, 3, 2) & "'"
    CodiceCer = "'" & Right$(CodiceCer, 2) & "'"
    
    Cmd.CommandText = "select Coefficiente from UnitaDiMisuraConvesioni where  UnitaDiMisuraDA ='" & UM1 & "' and UnitaDiMisuraA ='" & UM2 & "' and " & _
        " CategoriaAttivita = '" & CategoriaAttivita & "' and SottoCategoriaAttivita = '" & SottoCategoriaAttivita & "' and '" & _
        " CodiceCER = '" & CodiceCer & "'"
    Set Cmd.ActiveConnection = gDBC
    Cmd.CommandType = adCmdText
    Set rs = Cmd.Execute
    
    
    If rs.EOF Then
        ConvertiUnitaDiMisura = Valore * rs!Coefficiente
    Else
        ConvertiUnitaDiMisura = -1
    End If
    
    
    rs.Close
    Set Cmd.ActiveConnection = Nothing
End Function
